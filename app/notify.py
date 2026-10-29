"""
简单推送：企业微信 / 飞书 群机器人 Webhook
不存储 webhook，仅当次请求使用。
"""
from __future__ import annotations

import json
from typing import Any
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError


def _post_json(url: str, payload: dict, timeout: int = 8) -> tuple[bool, str]:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = Request(url, data=data, headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urlopen(req, timeout=timeout) as resp:
            body = resp.read().decode("utf-8", errors="ignore")
            if resp.status >= 200 and resp.status < 300:
                return True, body[:200]
            return False, f"HTTP {resp.status}: {body[:150]}"
    except HTTPError as e:
        return False, f"HTTP {e.code}"
    except URLError as e:
        return False, str(e.reason)[:100]
    except Exception as e:
        return False, str(e)[:100]


def detect_webhook_type(url: str) -> str:
    u = (url or "").lower()
    if "qyapi.weixin.qq.com" in u or "weixin" in u:
        return "wecom"
    if "feishu.cn" in u or "larksuite.com" in u or "open.feishu.cn" in u:
        return "feishu"
    return "generic"


def build_summary_text(summary: dict, anomaly_count: int = 0, filename: str = "") -> str:
    lines = [
        "【银行流水标准化完成】",
        f"文件: {filename or '多个文件'}",
        f"有效交易: {summary.get('有效交易笔数', '-') } 笔",
        f"收入合计: {summary.get('收入合计', '-') }",
        f"支出合计: {summary.get('支出合计', '-') }",
        f"净变动: {summary.get('净变动', '-') }",
        f"异常提示: {anomaly_count} 处",
        f"日期范围: {summary.get('日期范围', '-') }",
        "",
        "结果已生成，请在浏览器下载 Excel。本消息不含流水明细（隐私保护）。",
    ]
    return "\n".join(str(x) for x in lines)


def push_wecom(webhook: str, text: str) -> tuple[bool, str]:
    payload = {
        "msgtype": "text",
        "text": {"content": text},
    }
    return _post_json(webhook, payload)


def push_feishu(webhook: str, text: str) -> tuple[bool, str]:
    # 飞书自定义机器人 text 格式
    payload = {
        "msg_type": "text",
        "content": {"text": text},
    }
    return _post_json(webhook, payload)


def push_generic(webhook: str, text: str, extra: dict | None = None) -> tuple[bool, str]:
    payload = {"text": text, "msg": text}
    if extra:
        payload.update(extra)
    return _post_json(webhook, payload)


def send_notify(
    webhook: str,
    summary: dict,
    anomaly_count: int = 0,
    filename: str = "",
) -> dict[str, Any]:
    """
    根据 webhook 地址自动识别企业微信 / 飞书并推送摘要。
    返回 {ok, channel, message}
    """
    if not webhook or not webhook.startswith("http"):
        return {"ok": False, "channel": "", "message": "无效 webhook"}

    text = build_summary_text(summary, anomaly_count, filename)
    channel = detect_webhook_type(webhook)

    if channel == "wecom":
        ok, msg = push_wecom(webhook, text)
    elif channel == "feishu":
        ok, msg = push_feishu(webhook, text)
    else:
        ok, msg = push_generic(webhook, text, {"summary": summary})

    return {
        "ok": ok,
        "channel": channel,
        "message": "推送成功" if ok else f"推送失败: {msg}",
    }
