# -*- coding: utf-8 -*-
"""Agent WeChat（thisnick/agent-wechat）HTTP API 推送通道。

与微信桌面端 UIA 自动化完全独立的可选发送通道：
- 全局配置 Base URL / Token / Account（见 ConfigManager 的 agent_api_*）
- 文本：POST /accounts/{acct}/sends {threadId, kind:"text", text}
- @：mentions.userIds / mentions.all（群成员 ID 由成员接口按昵称匹配）
- 图片/文件：POST /uploads 暂存 → POST /sends {kind:"image"|"file", sourceId}

AgentWeChatSender 以鸭子类型兼容 send_executor.run_send_tasks() 的调用面，
勾选后由它替代 WeChatSender，不再操作微信桌面窗口。
"""
import json
import logging
import mimetypes
import os
import random
import time
import uuid
from typing import Dict, List, Optional

import urllib.error
import urllib.parse
import urllib.request

DEFAULT_BASE_URL = "http://127.0.0.1:6174/v1"
_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp"}
_SUCCESS_STATES = {"accepted", "sent", "delivered", "succeeded", "completed", "done"}
_FAILURE_STATES = {"failed", "rejected", "error", "cancelled", "canceled"}


class AgentWeChatError(RuntimeError):
    """API 调用失败（网络/鉴权/业务错误）。"""


class AgentWeChatClient:
    """最小可用的 Agent WeChat HTTP 客户端（仅标准库）。"""

    def __init__(self, base_url: str = "", token: str = "",
                 account: str = "", timeout: float = 30.0, log=None):
        base = str(base_url or "").strip() or DEFAULT_BASE_URL
        base = base.rstrip("/")
        if not base.endswith("/v1"):
            base += "/v1"
        self.base_url = base
        self.token = str(token or "").strip()
        self.account = str(account or "").strip()
        self.timeout = max(5.0, float(timeout or 30.0))
        self._log_fn = log
        self._thread_cache: Dict[str, str] = {}
        self._member_cache: Dict[str, List[dict]] = {}

    # ------------------------------------------------------------ 基础请求
    def _log(self, message: str) -> None:
        if self._log_fn is None:
            return
        try:
            self._log_fn(str(message))
        except Exception:
            pass

    def _request(self, method: str, path: str, body=None,
                 headers: Optional[dict] = None, timeout: float = None,
                 base: str = None, auth: bool = True):
        url = (base or self.base_url) + path
        data = None
        hdrs = {"Authorization": f"Bearer {self.token}"} if (auth and self.token) else {}
        if body is not None:
            data = json.dumps(body, ensure_ascii=False).encode("utf-8")
            hdrs["Content-Type"] = "application/json"
        elif method in ("POST", "PUT", "PATCH"):
            data = b"{}"
            hdrs["Content-Type"] = "application/json"
        hdrs.update(headers or {})
        req = urllib.request.Request(url, data=data, headers=hdrs, method=method)
        try:
            with urllib.request.urlopen(req, timeout=timeout or self.timeout) as resp:
                raw = resp.read()
        except urllib.error.HTTPError as exc:
            detail = ""
            try:
                raw_err = exc.read()
                if raw_err[:2] == b"\x1f\x8b":
                    import gzip
                    raw_err = gzip.decompress(raw_err)
                detail = raw_err.decode("utf-8", "replace")[:300]
            except Exception:
                pass
            raise AgentWeChatError(
                f"HTTP {exc.code} {method} {url}: {detail or exc.reason}"
            ) from exc
        except urllib.error.URLError as exc:
            raise AgentWeChatError(
                f"无法连接 Agent WeChat（{url}）：{exc.reason}"
            ) from exc
        if not raw:
            return {}
        try:
            return json.loads(raw.decode("utf-8", "replace"))
        except ValueError:
            return {"raw": raw.decode("utf-8", "replace")}

    def _upload_request(self, path: str, file_path: str, timeout: float = None):
        """multipart/form-data 上传单个文件（字段名 file）。"""
        boundary = "----esw" + uuid.uuid4().hex
        filename = os.path.basename(file_path)
        ctype = mimetypes.guess_type(filename)[0] or "application/octet-stream"
        with open(file_path, "rb") as fh:
            content = fh.read()
        parts = [
            f"--{boundary}\r\n".encode(),
            (f'Content-Disposition: form-data; name="file"; '
             f'filename="{filename}"\r\n').encode("utf-8"),
            f"Content-Type: {ctype}\r\n\r\n".encode(),
            content,
            f"\r\n--{boundary}--\r\n".encode(),
        ]
        data = b"".join(parts)
        hdrs = {"Content-Type": f"multipart/form-data; boundary={boundary}"}
        if self.token:
            hdrs["Authorization"] = f"Bearer {self.token}"
        req = urllib.request.Request(
            self.base_url + path, data=data, headers=hdrs, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=timeout or 300) as resp:
                raw = resp.read()
        except urllib.error.HTTPError as exc:
            detail = ""
            try:
                raw_err = exc.read()
                if raw_err[:2] == b"\x1f\x8b":
                    import gzip
                    raw_err = gzip.decompress(raw_err)
                detail = raw_err.decode("utf-8", "replace")[:300]
            except Exception:
                pass
            raise AgentWeChatError(
                f"HTTP {exc.code} 上传失败 {filename}: {detail or exc.reason}"
            ) from exc
        except urllib.error.URLError as exc:
            raise AgentWeChatError(f"上传失败（{filename}）：{exc.reason}") from exc
        try:
            return json.loads(raw.decode("utf-8", "replace"))
        except ValueError:
            return {"raw": raw.decode("utf-8", "replace")}

    @staticmethod
    def _items(payload) -> List[dict]:
        if isinstance(payload, list):
            return [x for x in payload if isinstance(x, dict)]
        if isinstance(payload, dict):
            for key in ("items", "data", "results"):
                value = payload.get(key)
                if isinstance(value, list):
                    return [x for x in value if isinstance(x, dict)]
        return []

    # ------------------------------------------------------------ 账号/连通
    @property
    def root_url(self) -> str:
        """不含 /v1 的服务根地址（/health 在根路径，不在 /v1 下）。"""
        return self.base_url[:-3] if self.base_url.endswith("/v1") else self.base_url

    def health(self) -> dict:
        # 注意：/health 在根路径且不需要 token；/v1/health 会 404
        try:
            return self._request("GET", "/health", base=self.root_url, auth=False)
        except AgentWeChatError:
            # 少数部署把 health 放在 /v1 下，兜底再试一次
            return self._request("GET", "/health", auth=False)

    def list_accounts(self) -> List[dict]:
        return self._items(self._request("GET", "/accounts"))

    def ensure_account(self) -> str:
        if self.account:
            return self.account
        accounts = self.list_accounts()
        if not accounts:
            raise AgentWeChatError("Agent WeChat 未登录任何微信账号（/accounts 为空）")
        acct = accounts[0].get("id") or accounts[0].get("accountId") or ""
        if not acct:
            raise AgentWeChatError("无法从 /accounts 解析账号 ID")
        self.account = str(acct)
        return self.account

    def probe_status(self):
        """探测服务与微信在线状态，返回 (是否在线, 说明)。

        - 先 GET /health（根路径）：区分「服务器关机/不可达」
        - 再 GET /v1/accounts：读取 connection/authenticated，区分「微信掉线」
        全程只读，不发送消息，不产生聊天流量。
        """
        try:
            self.health()
        except AgentWeChatError as exc:
            return False, f"服务器不可达：{exc}"
        try:
            accounts = self._items(self._request("GET", "/accounts"))
        except AgentWeChatError as exc:
            return False, f"服务在线，但账号接口异常：{exc}"
        if not accounts:
            return False, "服务在线，但没有已登录的微信账号"
        target = None
        if self.account:
            for item in accounts:
                if str(item.get("id") or "") == self.account:
                    target = item
                    break
            if target is None:
                return False, f"服务在线，但指定账号不在线：{self.account}"
        else:
            target = accounts[0]
        acct_id = str(target.get("id") or target.get("accountId") or "")
        conn = str(target.get("connection") or "").strip().lower()
        authed = target.get("authenticated")
        if authed is False or (conn and conn != "online"):
            return False, (
                f"微信未在线（account={acct_id or '未知'}，"
                f"connection={conn or '未知'}，authenticated={authed}）"
            )
        return True, f"在线（account={acct_id or '未知'}，connection={conn or 'online'}）"

    def _api(self, suffix: str) -> str:
        return f"/accounts/{urllib.parse.quote(self.ensure_account())}{suffix}"

    # ------------------------------------------------------------ 会话/成员
    @staticmethod
    def _name_of(item: dict) -> str:
        for key in ("displayName", "nickname", "remark", "alias",
                    "name", "title", "username", "id"):
            value = item.get(key)
            if isinstance(value, str) and value.strip():
                return value.strip()
        return ""

    _NAME_KEYS = ("displayName", "nickname", "remark", "alias", "name", "title")

    @classmethod
    def _match_exact(cls, items: List[dict], name: str) -> Optional[str]:
        target = str(name or "").strip().casefold()
        if not target:
            return None
        for item in items:
            for key in cls._NAME_KEYS:
                value = item.get(key)
                if isinstance(value, str) and value.strip().casefold() == target:
                    return str(item.get("id") or item.get("username") or "") or None
        return None

    @classmethod
    def _match_partial(cls, items: List[dict], name: str) -> Optional[str]:
        target = str(name or "").strip().casefold()
        if not target:
            return None
        for item in items:
            for key in cls._NAME_KEYS:
                value = item.get(key)
                if isinstance(value, str) and target in value.casefold():
                    return str(item.get("id") or item.get("username") or "") or None
        return None

    def _match(self, items: List[dict], name: str) -> Optional[str]:
        return (self._match_exact(items, name)
                or self._match_partial(items, name))

    def find_thread(self, recipient: str) -> str:
        """把昵称/群名/联系人名解析成 threadId（直聊=username，群=xxx@chatroom）。

        先对联系人/群/会话做「全等匹配」，全部都不中才做「包含匹配」：
        否则「工作」会先命中联系人「工作」而抢走真正的群「工作报告」。
        """
        name = str(recipient or "").strip()
        if not name:
            raise AgentWeChatError("接收人为空")
        if name in self._thread_cache:
            return self._thread_cache[name]
        # 已像 ID 的直接使用
        if (name.startswith("wxid_") or "@chatroom" in name
                or name.endswith("@openim") or name == "filehelper"):
            self._thread_cache[name] = name
            return name
        query = urllib.parse.urlencode({"search": name, "limit": 50})
        sources = [
            self._items(self._request("GET", self._api(f"/contacts?{query}"))),
            self._items(self._request("GET", self._api("/groups?limit=200"))),
            self._items(self._request("GET", self._api("/threads?limit=200"))),
        ]
        for items in sources:
            found = self._match_exact(items, name)
            if found:
                self._thread_cache[name] = found
                return found
        for items in sources:
            found = self._match_partial(items, name)
            if found:
                self._thread_cache[name] = found
                return found
        raise AgentWeChatError(f"Agent WeChat 找不到联系人或群：{name}")

    def list_members(self, thread_id: str, search: str = "") -> List[dict]:
        if not search and thread_id in self._member_cache:
            return self._member_cache[thread_id]
        params = {"limit": 200}
        if search:
            params["search"] = search
        query = urllib.parse.urlencode(params)
        items = self._items(self._request(
            "GET", self._api(f"/groups/{urllib.parse.quote(thread_id)}/members?{query}")))
        if not search:
            self._member_cache[thread_id] = items
        return items

    def resolve_mentions(self, thread_id: str, names: List[str],
                         at_all: bool = False) -> Optional[dict]:
        names = [str(n).strip() for n in (names or []) if str(n).strip()]
        # @ 只在群聊生效；私聊（含 @所有人）直接忽略，避免接口 400
        if not thread_id.endswith("@chatroom"):
            if names or at_all:
                self._log("⚠ @ 仅支持群聊，当前为私聊，已忽略 @")
            return None
        if at_all:
            return {"all": True}
        if not names:
            return None
        members = self.list_members(thread_id)
        ids: List[str] = []
        missed: List[str] = []
        for name in names:
            target = name.casefold()
            hit = None
            for m in members:
                values = [str(m.get(k) or "").strip().casefold()
                          for k in ("mentionName", "displayName", "nickname",
                                    "remark", "alias", "username", "id")]
                if target in values:
                    hit = str(m.get("id") or m.get("username") or "")
                    break
            if hit:
                ids.append(hit)
            else:
                missed.append(name)
        if missed:
            self._log(f"⚠ 群成员未匹配到，跳过 @：{'、'.join(missed)}")
        return {"userIds": ids} if ids else None

    # ------------------------------------------------------------ 发送
    def _wait_send(self, send: dict, timeout: float = 120.0) -> bool:
        """轮询 GET /sends/{id}；无法确认状态时按已提交成功处理。"""
        sid = str(send.get("sendId") or send.get("id") or "").strip()
        state = str(send.get("state") or "").strip().lower()
        if not sid or state in _SUCCESS_STATES:
            return True
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                info = self._request("GET", f"/sends/{urllib.parse.quote(sid)}")
            except AgentWeChatError:
                return True
            state = str(info.get("state") or "").strip().lower()
            if state in _SUCCESS_STATES:
                return True
            if state in _FAILURE_STATES:
                self._log(f"❌ 发送未成功（state={state}）: {info}")
                return False
            time.sleep(1.0)
        self._log(f"⚠ 等待发送结果超时（send={sid}），按已提交处理")
        return True

    def send_text(self, thread_id: str, text: str,
                  mentions: Optional[dict] = None) -> bool:
        body = {"threadId": thread_id, "kind": "text", "text": text}
        if mentions:
            body["mentions"] = mentions
        try:
            send = self._request(
                "POST", self._api("/sends"), body,
                headers={"Idempotency-Key": uuid.uuid4().hex})
        except AgentWeChatError as exc:
            detail = str(exc)
            # @所有人需要群主/管理员权限；权限不足时退化为普通消息，避免整条丢失
            if mentions and ("mention_all" in detail or "HTTP 403" in detail
                             or "HTTP 409" in detail):
                self._log(f"⚠ @ 发送被拒绝（权限不足），改为普通消息重试：{detail}")
                body.pop("mentions", None)
                send = self._request(
                    "POST", self._api("/sends"), body,
                    headers={"Idempotency-Key": uuid.uuid4().hex})
            else:
                raise
        return self._wait_send(send if isinstance(send, dict) else {})

    def upload(self, file_path: str) -> str:
        payload = self._upload_request(self._api("/uploads"), file_path)
        source = str((payload or {}).get("id") or (payload or {}).get("sourceId") or "")
        if not source:
            raise AgentWeChatError(f"上传未返回 sourceId: {payload}")
        return source

    def send_file(self, thread_id: str, file_path: str) -> bool:
        ext = os.path.splitext(file_path)[1].lower()
        kind = "image" if ext in _IMAGE_EXTS else "file"
        source = self.upload(file_path)
        send = self._request(
            "POST", self._api("/sends"),
            {"threadId": thread_id, "kind": kind, "sourceId": source},
            headers={"Idempotency-Key": uuid.uuid4().hex}, timeout=600)
        if isinstance(send, dict) and send.get("jobId"):
            # 附件任务：提交成功后交给守护进程继续，不阻塞整批发送
            return True
        return self._wait_send(send if isinstance(send, dict) else {}, timeout=600)


# ---------------------------------------------------------------------------
class AgentWeChatSender:
    """把 AgentWeChatClient 适配成 send_executor 需要的发送器接口。"""

    def __init__(self, client: AgentWeChatClient, log=None):
        self.client = client
        self._file_logger = logging.getLogger("agent_wechat")
        raw_log = log or (lambda _m: None)

        def _log_both(message):
            # 两条通道都写：GUI 日志面板 + 磁盘日志（便于排查线上失败）
            try:
                self._file_logger.info(str(message))
            except Exception:
                pass
            try:
                raw_log(str(message))
            except Exception:
                pass

        self.log = _log_both
        # 客户端内部日志（@ 未匹配、等待超时等）也统一走同一出口
        try:
            self.client._log_fn = _log_both
        except Exception:
            pass
        self._thread_id = ""
        self._recipient = ""
        self._helper = None
        # Agent 专属拟人节奏（与全局 wechatauto 拟人节流彼此独立）
        self._rhythm = self._load_rhythm()
        self._batch_actions = 0
        self._last_paced_person = ""
        if self._rhythm.get("enabled"):
            ml, mh = self._rhythm["msg"]
            pl, ph = self._rhythm["person"]
            self.log(
                f"🚶 Agent 专属拟人已启用：每条消息 {ml:g}~{mh:g}s，"
                f"每个人之间 {pl:g}~{ph:g}s"
            )

    def _load_rhythm(self) -> dict:
        """读取全局设置里的 Agent 专属拟人节奏（只影响本通道）。"""
        try:
            from modules.config_manager import ConfigManager
            cfg = (ConfigManager().load_global_settings()
                   .get("agent_api_rhythm") or {})
        except Exception:
            cfg = {}

        def _pair(name, dmin, dmax):
            try:
                lo = float(cfg.get(f"{name}_min", dmin))
                hi = float(cfg.get(f"{name}_max", dmax))
            except (TypeError, ValueError):
                lo, hi = dmin, dmax
            lo = max(0.0, min(600.0, lo))
            hi = max(0.0, min(600.0, hi))
            if hi < lo:
                lo, hi = hi, lo
            return lo, hi

        return {
            "enabled": bool(cfg.get("enabled", False)),
            "msg": _pair("msg", 2.0, 5.0),
            "person": _pair("person", 3.0, 8.0),
        }

    def _pace(self, stop_event, kind: str) -> bool:
        """按配置随机等待；返回 False 表示等待期间被用户停止。"""
        if not self._rhythm.get("enabled"):
            return True
        lo, hi = self._rhythm.get(kind) or (0.0, 0.0)
        if hi <= 0:
            return True
        delay = random.uniform(lo, max(lo, hi))
        if delay <= 0:
            return True
        end = time.time() + delay
        while True:
            if stop_event is not None and stop_event.is_set():
                return False
            remain = end - time.time()
            if remain <= 0:
                return True
            time.sleep(min(0.2, remain))

    def _wait_before_action(self, stop_event) -> bool:
        """同一个人内部：第 2 条消息起等待随机间隔。"""
        if not self._rhythm.get("enabled"):
            return True
        if self._batch_actions > 0:
            return self._pace(stop_event, "msg")
        return True

    def _after_action_done(self):
        self._batch_actions += 1

    def _pace_new_person(self, recipient: str, stop_event) -> bool:
        """换下一位接收人时等待「每个人之间」随机间隔；同一人重试不等待。"""
        if not self._rhythm.get("enabled"):
            self._last_paced_person = str(recipient)
            self._batch_actions = 0
            return True
        name = str(recipient)
        if self._last_paced_person and name != self._last_paced_person:
            if not self._pace(stop_event, "person"):
                return False
        self._last_paced_person = name
        self._batch_actions = 0
        return True

    # 图片渲染沿用 WeChatSender 的表格/区域截图实现（只做本地渲染，不启动 wxauto）
    def _img_helper(self):
        if self._helper is None:
            from modules.wechat_sender import WeChatSender
            helper = WeChatSender()
            try:
                helper.log = self.log
            except Exception:
                pass
            self._helper = helper
        return self._helper

    def _create_table_images(self, table_data):
        return self._img_helper()._create_table_images(table_data)

    def _remove_temp_image(self, path):
        return self._img_helper()._remove_temp_image(path)

    @staticmethod
    def _is_shared_region_image(path):
        from modules.wechat_sender import WeChatSender
        return WeChatSender._is_shared_region_image(path)

    def cleanup_temp_images(self, include_region_cache: bool = True):
        try:
            self._img_helper().cleanup_temp_images(
                include_region_cache=include_region_cache)
        except Exception:
            pass

    def minimize_window(self):
        """API 通道不需要操作微信窗口。"""
        return None

    def initialize(self) -> bool:
        try:
            account = self.client.ensure_account()
            self.log(f"Agent WeChat API 已连接（account={account}）")
            return True
        except Exception as exc:
            self.log(f"❌ Agent WeChat 连接失败: {exc}")
            return False

    def open_chat(self, recipient, chat_delay=0, stop_event=None, force=False) -> bool:
        # Agent 专属拟人：换人时按「每个人之间」随机等待
        if not self._pace_new_person(recipient, stop_event):
            self.log("⏹ 拟人等待中被停止")
            return False
        try:
            self._thread_id = self.client.find_thread(recipient)
            self._recipient = str(recipient)
            return True
        except Exception as exc:
            self.log(f"❌ Agent WeChat 找不到接收人 {recipient}: {exc}")
            return False

    def _mentions(self, at_members=None, at_all=False):
        names = at_members if isinstance(at_members, list) else (
            [at_members] if isinstance(at_members, str) and at_members else [])
        try:
            return self.client.resolve_mentions(
                self._thread_id, names, bool(at_all))
        except Exception as exc:
            self.log(f"⚠ 解析 @ 成员失败，本条不 @: {exc}")
            return None

    def send_message(self, text=None, recipient="", chat_delay=0, fast_mode=True,
                     stop_event=None, at_members=None, at_all=False,
                     content=None, **_kw) -> bool:
        # 兼容调用方用 content= 传文本（与 WeChatSender 签名一致）
        if text is None:
            text = content or ""
        try:
            if not self._thread_id or self._recipient != str(recipient):
                if not self.open_chat(recipient, stop_event=stop_event):
                    return False
            if not self._wait_before_action(stop_event):
                self.log("⏹ 拟人等待中被停止")
                return False
            ok = self.client.send_text(
                self._thread_id, str(text), self._mentions(at_members, at_all))
            if ok:
                self._after_action_done()
            return ok
        except Exception as exc:
            self.log(f"❌ Agent WeChat 发送失败: {exc}")
            return False

    def send_multiple_messages_progress(self, messages, recipient, chat_delay=0,
                                        start_index=0, stop_event=None,
                                        fast_mode=True, at_members=None,
                                        at_all=False, **_kw):
        messages = list(messages or [])
        idx = max(0, int(start_index or 0))
        while idx < len(messages):
            if stop_event is not None and stop_event.is_set():
                return False, idx
            # @ 只在第一条带，避免每段都 @
            ok = self.send_message(
                messages[idx], recipient, stop_event=stop_event,
                at_members=at_members if idx == 0 else None,
                at_all=at_all if idx == 0 else False)
            if not ok:
                return False, idx
            idx += 1
        return True, idx

    def send_file(self, file_path, recipient, chat_delay=0, fast_mode=True,
                  stop_event=None, at_members=None, at_all=False, **_kw) -> bool:
        try:
            if not self._thread_id or self._recipient != str(recipient):
                if not self.open_chat(recipient, stop_event=stop_event):
                    return False
            if not self._wait_before_action(stop_event):
                self.log("⏹ 拟人等待中被停止")
                return False
            ok = self.client.send_file(self._thread_id, str(file_path))
            if ok:
                self._after_action_done()
            return ok
        except Exception as exc:
            self.log(f"❌ Agent WeChat 文件发送失败: {exc}")
            return False

    def send_files_batch(self, files=None, recipient="", chat_delay=0,
                         fast_mode=True, stop_event=None, at_members=None,
                         at_all=False, file_paths=None, **_kw) -> bool:
        # 兼容调用方用 file_paths= 传文件列表（与 WeChatSender 签名一致）
        files = files if files is not None else file_paths
        for path in files or []:
            if stop_event is not None and stop_event.is_set():
                return False
            if not self.send_file(path, recipient, stop_event=stop_event):
                return False
        return True

    def send_text_and_files(self, text, files=None, recipient="", chat_delay=0,
                            fast_mode=True, stop_event=None, text_first=True,
                            at_members=None, at_all=False, file_paths=None,
                            **_kw) -> bool:
        # 兼容调用方用 file_paths= 传文件列表（与 WeChatSender 签名一致）
        if files is None:
            files = file_paths
        files = list(files or [])
        if text_first:
            if text and not self.send_message(
                    text, recipient, stop_event=stop_event,
                    at_members=at_members, at_all=at_all):
                return False
            return self.send_files_batch(
                files, recipient, stop_event=stop_event)
        if not self.send_files_batch(files, recipient, stop_event=stop_event):
            return False
        return (not text) or self.send_message(
            text, recipient, stop_event=stop_event,
            at_members=at_members, at_all=at_all)

    def send_table_images_progress(self, table_data, recipient, chat_delay=0,
                                   start_index=0, stop_event=None,
                                   fast_mode=True, at_members=None,
                                   at_all=False, **_kw):
        paths = []
        try:
            paths = list(self._create_table_images(table_data) or [])
        except Exception as exc:
            self.log(f"❌ 生成表格图片失败: {exc}")
            return False, max(0, int(start_index or 0))
        idx = max(0, int(start_index or 0))
        while idx < len(paths):
            if stop_event is not None and stop_event.is_set():
                self._cleanup(paths)
                return False, idx
            if not self.send_file(paths[idx], recipient, stop_event=stop_event):
                self._cleanup(paths)
                return False, idx
            idx += 1
        self._cleanup(paths)
        return True, idx

    def _cleanup(self, paths):
        for path in paths:
            try:
                self._remove_temp_image(path)
            except Exception:
                pass
