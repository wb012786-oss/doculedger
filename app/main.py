"""
银行对账单标准化工具 - 主入口
隐私优先：上传文件处理后立即删除，不落库、不记录敏感内容。
"""
from __future__ import annotations

import os
import shutil
import tempfile
import uuid
from pathlib import Path
from typing import List

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

from .config import ALLOWED_EXTENSIONS, MAX_FILE_SIZE, MAX_FILES, UPLOAD_DIR
from .notify import send_notify
from .processor import process_uploads, safe_cleanup
from urllib.parse import quote

app = FastAPI(
    title="银行对账单标准化工具",
    description="PDF/Excel 银行流水一键转标准 Excel，支持差异检测与异常提示。处理完即删，不存储您的数据。",
    version="1.0.0",
    docs_url=None,  # 生产关闭 swagger，防止误用
    redoc_url=None,
)

# 静态与模板
BASE = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(BASE / "templates"))
static_dir = BASE / "static"
if static_dir.exists():
    app.mount("/static", StaticFiles(directory=str(static_dir)), name="static")


class Health(BaseModel):
    status: str
    version: str


@app.get("/health", response_model=Health)
async def health():
    return {"status": "ok", "version": "1.0.0"}


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "max_size_mb": MAX_FILE_SIZE // (1024 * 1024),
            "max_files": MAX_FILES,
        },
    )


@app.post("/api/process")
async def process(
    files: List[UploadFile] = File(...),
    webhook: str = Form(default=""),
):
    if not files:
        raise HTTPException(400, "请至少上传一个文件")

    if len(files) > MAX_FILES:
        raise HTTPException(400, f"一次最多上传 {MAX_FILES} 个文件")

    temp_paths: list[Path] = []
    original_names: list[str] = []

    try:
        for f in files:
            if not f.filename:
                continue
            suffix = Path(f.filename).suffix.lower()
            if suffix not in ALLOWED_EXTENSIONS:
                raise HTTPException(
                    400,
                    f"不支持的文件类型: {f.filename}。仅支持 PDF / Excel / CSV",
                )

            content = await f.read()
            if len(content) > MAX_FILE_SIZE:
                raise HTTPException(
                    400,
                    f"文件 {f.filename} 超过大小限制 ({MAX_FILE_SIZE // 1024 // 1024}MB)",
                )
            if len(content) == 0:
                raise HTTPException(400, f"文件 {f.filename} 为空")

            safe_name = f"{uuid.uuid4().hex}{suffix}"
            temp_path = UPLOAD_DIR / safe_name
            temp_path.write_bytes(content)
            temp_paths.append(temp_path)
            original_names.append(Path(f.filename).name)

        if not temp_paths:
            raise HTTPException(400, "没有有效文件")

        excel_bytes, result = process_uploads(temp_paths, original_names)

        # 立即清理临时文件
        safe_cleanup(temp_paths)

        # 可选：推送摘要到企业微信 / 飞书（仅摘要，不含明细）
        notify_msg = ""
        if webhook and webhook.strip():
            notify_result = send_notify(
                webhook.strip(),
                result.get("summary") or {},
                anomaly_count=result.get("anomaly_count", 0),
                filename=", ".join(original_names[:3]),
            )
            notify_msg = notify_result.get("message", "")

        filename_cn = "银行流水标准化结果.xlsx"
        filename_ascii = "bank_statement_result.xlsx"
        headers = {
            "Content-Disposition": (
				f'attachment; filename="{filename_ascii}"; '
				f"filename*=UTF-8''{quote(filename_cn)}"
            ),
            "X-Process-Ok": "1" if result.get("ok") else "0",
            "X-Row-Count": str(result.get("row_count", 0)),
            "X-Anomaly-Count": str(result.get("anomaly_count", 0)),
            "Cache-Control": "no-store, no-cache, must-revalidate",
            "Pragma": "no-cache",
        }
        if notify_msg:
            headers["X-Notify"] = "ok" if "成功" in notify_msg else "fail"

        return Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )

    except HTTPException:
        safe_cleanup(temp_paths)
        raise
    except Exception as e:
        safe_cleanup(temp_paths)
        import traceback
        traceback.print_exc()  # 完整错误打在 PowerShell 里
        raise HTTPException(500, f"处理失败: {type(e).__name__}: {e}")
    # except Exception:
        # safe_cleanup(temp_paths)
        # raise HTTPException(500, "处理失败，请检查文件格式后重试。如持续失败可联系反馈。")


@app.get("/privacy", response_class=HTMLResponse)
async def privacy(request: Request):
    return templates.TemplateResponse("privacy.html", {"request": request})

@app.get("/debug-upload-dir")
async def debug_upload_dir():
    from .config import UPLOAD_DIR
    return {
        "upload_dir": str(UPLOAD_DIR),
        "exists": UPLOAD_DIR.exists(),
        "writable": os.access(UPLOAD_DIR, os.W_OK),
    }

# 启动时清理残留临时文件
@app.on_event("startup")
async def startup_cleanup():
    try:
        for p in UPLOAD_DIR.glob("*"):
            if p.is_file():
                p.unlink(missing_ok=True)
    except Exception:
        pass


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)