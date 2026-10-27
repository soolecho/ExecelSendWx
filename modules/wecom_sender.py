"""企业微信群机器人 webhook 发送（v1.8.0）。

纯 HTTP 发送，不依赖微信客户端，供监控轮询任务作为独立发送渠道：
- send_text    → 文本消息（超 2000 字符截断，企微 text 上限 2048 字节）
- send_file    → 先 upload_media(type=file) 拿 media_id，再发文件消息
- send_image   → 先 upload_media(type=image) 拿 media_id，再发图片消息
- send_failure_notice → 发送失败摘要文本

判定：响应 JSON 的 errcode == 0 为成功（errcode 非 0 时 errmsg 为失败原因）。
upload_media 的 URL 由 webhook 推导：/webhook/send → /webhook/upload_media 并追加 &type=xxx。
"""
import logging
import os

logger = logging.getLogger("ExcelSendWx.wecom")

TEXT_MAX_CHARS = 2000  # 企微 text 上限 2048 字节，按字符截断留余量


def _err(detail: str) -> tuple:
    logger.warning("企微 webhook 发送失败: %s", detail)
    return False, detail


def _post_json(webhook: str, payload: dict, timeout: float):
    """POST 一次企微 webhook，返回 (ok, detail)。"""
    import requests
    try:
        resp = requests.post(webhook, json=payload, timeout=timeout)
    except Exception as exc:
        return False, f"网络异常: {exc!r}"
    try:
        body = resp.json()
    except ValueError:
        return False, f"响应非 JSON: {resp.text[:120]}"
    if body.get("errcode") != 0:
        return False, f"errcode={body.get('errcode')} errmsg={body.get('errmsg')}"
    return True, "ok"


def _upload_media(webhook: str, path: str, media_type: str, timeout: float):
    """上传临时素材（file/image），返回 (ok, media_id or detail)。"""
    import requests
    if not webhook or not os.path.isfile(path):
        return False, "webhook 为空或文件不存在"
    upload_url = webhook.replace("/webhook/send", "/webhook/upload_media")
    sep = "&" if "?" in upload_url else "?"
    upload_url = f"{upload_url}{sep}type={media_type}"
    try:
        with open(path, "rb") as fh:
            resp = requests.post(
                upload_url,
                files={"media": (os.path.basename(path), fh,
                                 "application/octet-stream")},
                timeout=timeout,
            )
    except Exception as exc:
        return False, f"上传失败（{media_type}）: {exc!r}"
    try:
        body = resp.json()
    except ValueError:
        return False, f"上传响应非 JSON: {resp.text[:120]}"
    if body.get("errcode") != 0:
        return False, f"上传 errcode={body.get('errcode')} errmsg={body.get('errmsg')}"
    media_id = body.get("media_id")
    if not media_id:
        return False, "上传成功但未返回 media_id"
    return True, media_id


def send_text(webhook: str, content: str, timeout: float = 15.0) -> tuple:
    """发送文本消息到企微群，返回 (ok, detail)。"""
    content = str(content or "")
    if len(content) > TEXT_MAX_CHARS:
        content = content[:TEXT_MAX_CHARS] + "\n…(内容过长已截断)"
    if not content.strip():
        return True, "空内容，跳过"
    return _post_json(webhook, {"msgtype": "text", "text": {"content": content}},
                      timeout)


def send_file(webhook: str, path: str, timeout: float = 30.0) -> tuple:
    """上传并发送文件消息，返回 (ok, detail)。"""
    ok, media = _upload_media(webhook, path, "file", timeout)
    if not ok:
        return _err(media)
    return _post_json(webhook, {"msgtype": "file", "file": {"media_id": media}},
                      timeout)


def send_image(webhook: str, path: str, timeout: float = 30.0) -> tuple:
    """上传并发送图片消息，返回 (ok, detail)。"""
    ok, media = _upload_media(webhook, path, "image", timeout)
    if not ok:
        return _err(media)
    return _post_json(webhook, {"msgtype": "image", "image": {"media_id": media}},
                      timeout)


def send_failure_notice(webhook: str, lines, timeout: float = 15.0) -> tuple:
    """发送失败摘要文本（多行内容），返回 (ok, detail)。"""
    content = "\n".join(str(x) for x in lines if str(x).strip())
    return send_text(webhook, content, timeout=timeout)
