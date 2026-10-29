"""
银行对账单解析器
支持 PDF（文本型 + OCR 扫描件）、Excel/CSV。
扫描件 PDF 使用 Tesseract OCR（中英）兜底。
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import pdfplumber
from dateutil import parser as date_parser

from .config import BANK_KEYWORDS, EXPENSE_KEYWORDS, INCOME_KEYWORDS

# OCR 可选依赖
_OCR_AVAILABLE = False
try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image
    _OCR_AVAILABLE = True
except ImportError:
    pass


def detect_bank(text: str) -> str:
    text_upper = text.upper()
    for bank, keywords in BANK_KEYWORDS.items():
        if bank == "通用":
            continue
        for kw in keywords:
            if kw.upper() in text_upper or kw in text:
                return bank
    return "通用"


def clean_amount(val: Any) -> float | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().replace(",", "").replace("，", "").replace("¥", "").replace("￥", "")
    s = re.sub(r"[^\d.\-]", "", s)
    if not s or s in ("-", "."):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(val: Any) -> str | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if not s:
        return None
    # 常见中文格式
    for fmt in (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y.%m.%d",
        "%Y年%m月%d日",
        "%m-%d",
        "%m/%d",
        "%Y%m%d",
    ):
        try:
            dt = datetime.strptime(s[:20], fmt)
            if dt.year < 2000:
                dt = dt.replace(year=datetime.now().year)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue
    try:
        dt = date_parser.parse(s, fuzzy=True)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None


def classify_type(desc: str, amount: float | None) -> str:
    if amount is None:
        return "未知"
    desc = str(desc or "")
    for kw in INCOME_KEYWORDS:
        if kw in desc:
            return "收入"
    for kw in EXPENSE_KEYWORDS:
        if kw in desc:
            return "支出"
    # 按金额正负兜底
    if amount > 0:
        return "收入"
    if amount < 0:
        return "支出"
    return "未知"


def ocr_pdf_to_text(path: Path, max_pages: int = 15) -> str:
    """扫描件 PDF → 文本（中英混合 OCR）"""
    if not _OCR_AVAILABLE:
        return ""
    try:
        images = convert_from_path(str(path), dpi=200, first_page=1, last_page=max_pages)
        texts = []
        for img in images:
            # chi_sim+eng 对中文银行流水效果较好
            txt = pytesseract.image_to_string(img, lang="chi_sim+eng")
            texts.append(txt)
        return "\n".join(texts)
    except Exception:
        return ""


def parse_ocr_text_to_rows(text: str) -> list[dict]:
    """
    从 OCR 纯文本里用正则尽量抽出「日期 + 金额」流水行。
    不保证 100% 准确，仅作兜底。
    """
    rows = []
    # 常见：2026-10-01 或 2026/10/01 或 10-01，后面跟金额
    line_pat = re.compile(
        r"(?P<date>\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?|\d{1,2}[-/.]\d{1,2})\s+"
        r"(?P<desc>.{0,40}?)\s*"
        r"(?P<amount>-?[\d,]+\.?\d*)"
    )
    for line in text.splitlines():
        line = line.strip()
        if len(line) < 8:
            continue
        m = line_pat.search(line)
        if m:
            rows.append({
                "交易日期": m.group("date"),
                "摘要": m.group("desc").strip() or "OCR提取",
                "金额": m.group("amount"),
                "余额": "",
                "对方信息": "",
            })
    return rows


def extract_tables_from_pdf(path: Path) -> tuple[list[pd.DataFrame], str, bool]:
    """
    返回 (tables, full_text, used_ocr)
    """
    tables = []
    full_text = ""
    used_ocr = False
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"
            page_tables = page.extract_tables() or []
            for t in page_tables:
                if not t or len(t) < 2:
                    continue
                header = [str(c).strip() if c else f"col_{i}" for i, c in enumerate(t[0])]
                rows = []
                for row in t[1:]:
                    if not any(cell for cell in row):
                        continue
                    rows.append([str(c).strip() if c else "" for c in row])
                if rows:
                    df = pd.DataFrame(rows, columns=header[: len(rows[0])])
                    tables.append(df)

    # 文本极少 → 尝试 OCR
    if len(full_text.strip()) < 80 and _OCR_AVAILABLE:
        ocr_text = ocr_pdf_to_text(path)
        if ocr_text and len(ocr_text.strip()) > 30:
            full_text = ocr_text
            used_ocr = True
            ocr_rows = parse_ocr_text_to_rows(ocr_text)
            if ocr_rows:
                df = pd.DataFrame(ocr_rows)
                tables.append(df)

    return tables, full_text, used_ocr


def normalize_dataframe(df: pd.DataFrame, source_name: str = "") -> pd.DataFrame:
    """把各种乱七八糟的列名统一成标准字段"""
    if df.empty:
        return pd.DataFrame(columns=["交易日期", "摘要", "金额", "余额", "交易类型", "对方信息", "来源"])

    # 列名清洗
    col_map = {}
    for col in df.columns:
        c = str(col).strip().lower()
        if any(k in c for k in ["日期", "date", "交易日", "记账日"]):
            col_map[col] = "交易日期"
        elif any(k in c for k in ["摘要", "说明", "备注", "用途", "desc", "narration", "交易明细"]):
            col_map[col] = "摘要"
        elif any(k in c for k in ["金额", "发生额", "amount", "交易金额"]):
            col_map[col] = "金额"
        elif any(k in c for k in ["余额", "balance", "账户余额"]):
            col_map[col] = "余额"
        elif any(k in c for k in ["收入", "贷方", "credit", "转入"]):
            col_map[col] = "收入"
        elif any(k in c for k in ["支出", "借方", "debit", "转出"]):
            col_map[col] = "支出"
        elif any(k in c for k in ["对方", "对手", "户名", "对方账户", "counter"]):
            col_map[col] = "对方信息"

    df = df.rename(columns=col_map)

    # 如果有单独的收入/支出列，合并成金额
    if "金额" not in df.columns:
        if "收入" in df.columns and "支出" in df.columns:
            def merge_amount(row):
                inc = clean_amount(row.get("收入"))
                exp = clean_amount(row.get("支出"))
                if inc and inc != 0:
                    return abs(inc)
                if exp and exp != 0:
                    return -abs(exp)
                return None
            df["金额"] = df.apply(merge_amount, axis=1)
        elif "收入" in df.columns:
            df["金额"] = df["收入"].apply(lambda x: abs(clean_amount(x) or 0) or None)
        elif "支出" in df.columns:
            df["金额"] = df["支出"].apply(lambda x: -abs(clean_amount(x) or 0) or None)

    # 标准化列
    result = pd.DataFrame()
    result["交易日期"] = df.get("交易日期", pd.Series([None] * len(df))).apply(parse_date)
    result["摘要"] = df.get("摘要", pd.Series([""] * len(df))).astype(str).str.strip()
    result["金额"] = df.get("金额", pd.Series([None] * len(df))).apply(clean_amount)
    result["余额"] = df.get("余额", pd.Series([None] * len(df))).apply(clean_amount)
    result["对方信息"] = df.get("对方信息", pd.Series([""] * len(df))).astype(str).str.strip()
    result["来源"] = source_name

    # 过滤完全空行
    result = result.dropna(how="all", subset=["交易日期", "金额", "摘要"])
    result = result[result["金额"].notna() | (result["摘要"].str.len() > 0)]

    # 交易类型
    result["交易类型"] = result.apply(
        lambda r: classify_type(r["摘要"], r["金额"]), axis=1
    )

    # 金额统一：收入为正，支出为负（如果原来是正数支出则翻转）
    def normalize_sign(row):
        amt = row["金额"]
        if amt is None:
            return None
        t = row["交易类型"]
        if t == "支出" and amt > 0:
            return -amt
        if t == "收入" and amt < 0:
            return abs(amt)
        return amt

    result["金额"] = result.apply(normalize_sign, axis=1)

    # 排序
    result = result.sort_values("交易日期", na_position="last").reset_index(drop=True)
    return result[
        ["交易日期", "摘要", "金额", "余额", "交易类型", "对方信息", "来源"]
    ]


def parse_file(path: Path, original_name: str = "") -> tuple[pd.DataFrame, dict]:
    """
    解析单个文件，返回标准化 DataFrame + 元信息
    """
    suffix = path.suffix.lower()
    meta = {
        "filename": original_name or path.name,
        "bank": "通用",
        "page_count": 0,
        "raw_rows": 0,
        "parse_method": "",
        "warnings": [],
    }

    if suffix == ".pdf":
        try:
            tables, full_text, used_ocr = extract_tables_from_pdf(path)
            meta["bank"] = detect_bank(full_text)
            meta["page_count"] = len(tables)
            meta["parse_method"] = "pdf_ocr" if used_ocr else "pdf_table"
            if used_ocr:
                meta["warnings"].append("已启用 OCR 识别扫描件（结果请人工复核）")

            if not tables:
                meta["warnings"].append("未检测到标准表格或 OCR 有效行，建议导出 Excel 后重试")
                df = pd.DataFrame(columns=["交易日期", "摘要", "金额", "余额", "交易类型", "对方信息", "来源"])
                return df, meta

            dfs = []
            for i, t in enumerate(tables):
                norm = normalize_dataframe(t, source_name=f"{original_name}#表{i+1}")
                if not norm.empty:
                    dfs.append(norm)
            if not dfs:
                meta["warnings"].append("表格/OCR 解析后无有效交易记录")
                return pd.DataFrame(columns=["交易日期", "摘要", "金额", "余额", "交易类型", "对方信息", "来源"]), meta

            df = pd.concat(dfs, ignore_index=True)
            meta["raw_rows"] = len(df)
            return df, meta

        except Exception as e:
            meta["warnings"].append(f"PDF 解析失败: {str(e)[:100]}")
            return pd.DataFrame(columns=["交易日期", "摘要", "金额", "余额", "交易类型", "对方信息", "来源"]), meta

    elif suffix in (".xlsx", ".xls"):
        try:
            # 读取所有 sheet
            xls = pd.ExcelFile(path)
            dfs = []
            for sheet in xls.sheet_names:
                raw = pd.read_excel(path, sheet_name=sheet, dtype=str)
                norm = normalize_dataframe(raw, source_name=f"{original_name}#{sheet}")
                if not norm.empty:
                    dfs.append(norm)
            if not dfs:
                meta["warnings"].append("Excel 中未找到有效交易数据")
                return pd.DataFrame(columns=["交易日期", "摘要", "金额", "余额", "交易类型", "对方信息", "来源"]), meta
            df = pd.concat(dfs, ignore_index=True)
            meta["raw_rows"] = len(df)
            meta["parse_method"] = "excel"
            meta["bank"] = "Excel导入"
            return df, meta
        except Exception as e:
            meta["warnings"].append(f"Excel 解析失败: {str(e)[:100]}")
            return pd.DataFrame(columns=["交易日期", "摘要", "金额", "余额", "交易类型", "对方信息", "来源"]), meta

    elif suffix == ".csv":
        try:
            for enc in ("utf-8", "gbk", "gb2312", "utf-8-sig"):
                try:
                    raw = pd.read_csv(path, encoding=enc, dtype=str)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                raise ValueError("无法识别 CSV 编码")
            df = normalize_dataframe(raw, source_name=original_name)
            meta["raw_rows"] = len(df)
            meta["parse_method"] = "csv"
            meta["bank"] = "CSV导入"
            return df, meta
        except Exception as e:
            meta["warnings"].append(f"CSV 解析失败: {str(e)[:100]}")
            return pd.DataFrame(columns=["交易日期", "摘要", "金额", "余额", "交易类型", "对方信息", "来源"]), meta

    else:
        meta["warnings"].append(f"不支持的文件类型: {suffix}")
        return pd.DataFrame(columns=["交易日期", "摘要", "金额", "余额", "交易类型", "对方信息", "来源"]), meta


def find_anomalies(df: pd.DataFrame) -> pd.DataFrame:
    """简单异常检测：大额、重复、缺失日期等"""
    if df.empty:
        return pd.DataFrame(columns=["交易日期", "摘要", "金额", "异常类型", "说明"])

    anomalies = []
    # 大额（超过均值 5 倍或绝对值 > 5万）
    amounts = df["金额"].dropna().abs()
    if len(amounts) > 3:
        mean = amounts.mean()
        threshold = max(mean * 5, 50000)
        for idx, row in df.iterrows():
            if row["金额"] is not None and abs(row["金额"]) >= threshold:
                anomalies.append({
                    "交易日期": row["交易日期"],
                    "摘要": row["摘要"],
                    "金额": row["金额"],
                    "异常类型": "大额交易",
                    "说明": f"金额超过阈值 {threshold:,.0f}",
                })

    # 重复交易（同日期+同金额+同摘要前20字）
    df2 = df.copy()
    df2["_key"] = (
        df2["交易日期"].astype(str)
        + "|"
        + df2["金额"].astype(str)
        + "|"
        + df2["摘要"].str[:20]
    )
    dupes = df2[df2.duplicated("_key", keep=False)]
    for idx, row in dupes.iterrows():
        anomalies.append({
            "交易日期": row["交易日期"],
            "摘要": row["摘要"],
            "金额": row["金额"],
            "异常类型": "疑似重复",
            "说明": "相同日期+金额+摘要",
        })

    # 缺失日期
    missing_date = df[df["交易日期"].isna()]
    for idx, row in missing_date.iterrows():
        anomalies.append({
            "交易日期": None,
            "摘要": row["摘要"],
            "金额": row["金额"],
            "异常类型": "日期缺失",
            "说明": "无法解析交易日期",
        })

    if not anomalies:
        return pd.DataFrame(columns=["交易日期", "摘要", "金额", "异常类型", "说明"])
    return pd.DataFrame(anomalies).drop_duplicates()


def reconcile_two(df_a: pd.DataFrame, df_b: pd.DataFrame) -> dict:
    """
    两份流水简单对账：按 日期+金额 匹配，找出只在一边出现的记录
    """
    if df_a.empty or df_b.empty:
        return {
            "matched": 0,
            "only_in_a": len(df_a),
            "only_in_b": len(df_b),
            "diff_a": df_a,
            "diff_b": df_b,
        }

    a = df_a.copy()
    b = df_b.copy()
    a["_key"] = a["交易日期"].astype(str) + "|" + a["金额"].round(2).astype(str)
    b["_key"] = b["交易日期"].astype(str) + "|" + b["金额"].round(2).astype(str)

    keys_a = set(a["_key"])
    keys_b = set(b["_key"])
    matched_keys = keys_a & keys_b

    only_a = a[~a["_key"].isin(matched_keys)].drop(columns=["_key"], errors="ignore")
    only_b = b[~b["_key"].isin(matched_keys)].drop(columns=["_key"], errors="ignore")

    return {
        "matched": len(matched_keys),
        "only_in_a": len(only_a),
        "only_in_b": len(only_b),
        "diff_a": only_a,
        "diff_b": only_b,
    }