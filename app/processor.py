"""
处理与 Excel 生成
隐私：所有临时文件在函数结束后立即删除
"""
from __future__ import annotations

import io
import uuid
from pathlib import Path
from typing import BinaryIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

from .parser import find_anomalies, parse_file, reconcile_two


def style_header(ws, row=1):
    fill = PatternFill("solid", fgColor="1F4E79")
    font = Font(bold=True, color="FFFFFF")
    for cell in ws[row]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def auto_width(ws, max_width=40):
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if len(str(cell.value or "")) > max_length:
                    max_length = len(str(cell.value or ""))
            except Exception:
                pass
        adjusted = min(max_length + 2, max_width)
        ws.column_dimensions[column_letter].width = max(adjusted, 10)


def build_excel(
    transactions: pd.DataFrame,
    anomalies: pd.DataFrame,
    summary: dict,
    reconcile_result: dict | None = None,
    meta_list: list[dict] | None = None,
) -> bytes:
    """生成带多 Sheet 的标准化 Excel，返回 bytes"""
    wb = Workbook()

    # Sheet 1: 标准化流水
    ws1 = wb.active
    ws1.title = "标准化流水"
    if not transactions.empty:
        for r in dataframe_to_rows(transactions, index=False, header=True):
            ws1.append(r)
        style_header(ws1)
        auto_width(ws1)
        # 金额列格式
        for row in ws1.iter_rows(min_row=2, min_col=3, max_col=4):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'
    else:
        ws1.append(["无有效交易记录"])

    # Sheet 2: 汇总
    ws2 = wb.create_sheet("汇总报告")
    ws2.append(["项目", "数值"])
    style_header(ws2)
    for k, v in summary.items():
        ws2.append([k, v])
    auto_width(ws2)

    # Sheet 3: 异常提示
    ws3 = wb.create_sheet("异常提示")
    if not anomalies.empty:
        for r in dataframe_to_rows(anomalies, index=False, header=True):
            ws3.append(r)
        style_header(ws3)
        auto_width(ws3)
        # 高亮异常
        red_fill = PatternFill("solid", fgColor="FFC7CE")
        for row in ws3.iter_rows(min_row=2):
            for cell in row:
                cell.fill = red_fill
    else:
        ws3.append(["未发现明显异常"])
        ws3["A1"].font = Font(color="006600")

    # Sheet 4: 对账差异（如果有）
    if reconcile_result:
        ws4 = wb.create_sheet("对账差异")
        ws4.append(["匹配笔数", reconcile_result.get("matched", 0)])
        ws4.append(["仅在文件A出现", reconcile_result.get("only_in_a", 0)])
        ws4.append(["仅在文件B出现", reconcile_result.get("only_in_b", 0)])
        ws4.append([])
        ws4.append(["—— 仅在文件A的记录 ——"])
        diff_a = reconcile_result.get("diff_a")
        if diff_a is not None and not diff_a.empty:
            for r in dataframe_to_rows(diff_a, index=False, header=True):
                ws4.append(r)
        else:
            ws4.append(["无"])
        ws4.append([])
        ws4.append(["—— 仅在文件B的记录 ——"])
        diff_b = reconcile_result.get("diff_b")
        if diff_b is not None and not diff_b.empty:
            for r in dataframe_to_rows(diff_b, index=False, header=True):
                ws4.append(r)
        else:
            ws4.append(["无"])
        auto_width(ws4)

    # Sheet 5: 解析元信息
    if meta_list:
        ws5 = wb.create_sheet("解析信息")
        ws5.append(["文件名", "识别银行", "解析方式", "有效行数", "警告"])
        style_header(ws5)
        for m in meta_list:
            ws5.append([
                m.get("filename", ""),
                m.get("bank", ""),
                m.get("parse_method", ""),
                m.get("raw_rows", 0),
                "; ".join(m.get("warnings", [])),
            ])
        auto_width(ws5)

    # 隐私声明
    ws6 = wb.create_sheet("隐私说明")
    ws6.append(["本工具隐私承诺"])
    ws6.append(["1. 您上传的文件仅在内存/临时目录中处理，处理完成后立即删除。"])
    ws6.append(["2. 服务器不会永久存储您的银行流水或任何敏感数据。"])
    ws6.append(["3. 本工具不会将数据用于训练模型或第三方共享。"])
    ws6.append(["4. 建议处理完成后立即下载结果并关闭页面。"])
    ws6.append(["5. 本工具仅供辅助整理，结果请人工复核后再用于正式记账/报税。"])
    ws6.column_dimensions["A"].width = 80

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def process_uploads(
    file_paths: list[Path],
    original_names: list[str],
) -> tuple[bytes, dict]:
    """
    主处理入口。
    返回 (excel_bytes, result_summary)
    """
    all_dfs = []
    meta_list = []
    for path, name in zip(file_paths, original_names):
        df, meta = parse_file(path, name)
        meta_list.append(meta)
        if not df.empty:
            all_dfs.append(df)

    if not all_dfs:
        # 返回空结果 Excel
        empty = pd.DataFrame(columns=["交易日期", "摘要", "金额", "余额", "交易类型", "对方信息", "来源"])
        summary = {
            "状态": "解析失败或无有效数据",
            "提示": "请检查文件是否为银行对账单（支持文本型PDF、Excel、CSV）。扫描件PDF效果较差。",
            "文件数": len(file_paths),
        }
        excel = build_excel(empty, pd.DataFrame(), summary, meta_list=meta_list)
        return excel, {"ok": False, "summary": summary, "meta": meta_list}

    transactions = pd.concat(all_dfs, ignore_index=True)
    # 去重（完全相同行）
    transactions = transactions.drop_duplicates(
        subset=["交易日期", "摘要", "金额"], keep="first"
    ).reset_index(drop=True)

    anomalies = find_anomalies(transactions)

    # 汇总
    total_income = transactions.loc[transactions["金额"] > 0, "金额"].sum()
    total_expense = transactions.loc[transactions["金额"] < 0, "金额"].sum()
    summary = {
        "处理文件数": len(file_paths),
        "有效交易笔数": len(transactions),
        "收入合计": round(float(total_income or 0), 2),
        "支出合计": round(float(abs(total_expense or 0)), 2),
        "净变动": round(float((total_income or 0) + (total_expense or 0)), 2),
        "异常笔数": len(anomalies),
        "识别银行": ", ".join(sorted({m.get("bank", "通用") for m in meta_list})),
        "日期范围": f"{transactions['交易日期'].min()} ~ {transactions['交易日期'].max()}"
        if transactions["交易日期"].notna().any()
        else "未知",
    }

    # 如果正好两个文件，自动做对账
    reconcile_result = None
    if len(all_dfs) == 2:
        reconcile_result = reconcile_two(all_dfs[0], all_dfs[1])
        summary["对账_匹配笔数"] = reconcile_result["matched"]
        summary["对账_仅文件A"] = reconcile_result["only_in_a"]
        summary["对账_仅文件B"] = reconcile_result["only_in_b"]

    excel_bytes = build_excel(
        transactions, anomalies, summary, reconcile_result, meta_list
    )

    result = {
        "ok": True,
        "summary": summary,
        "meta": meta_list,
        "row_count": len(transactions),
        "anomaly_count": len(anomalies),
    }
    return excel_bytes, result


def safe_cleanup(paths: list[Path]):
    """强制删除临时文件，忽略错误"""
    for p in paths:
        try:
            if p.exists():
                p.unlink()
        except Exception:
            pass