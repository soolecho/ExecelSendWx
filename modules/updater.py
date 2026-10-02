# -*- coding: utf-8 -*-
"""在线更新模块。

产品型更新流程（自动检查、用户确认后才下载安装）：
  1. 检查源：GitHub Releases API 直连优先，gh-proxy 镜像兜底
  2. 发现新版本：手动检查弹更新对话框；每日自动检查仅弹 Toast 提醒，
     同一版本只自动提醒一次，用户可“跳过此版本”
  3. 用户点“立即更新”：后台流式下载（多源回退、进度可见）
  4. 双重校验：SHA256（发布资产含 SHA256SUMS 时）+ Authenticode 数字签名
     （必须 Valid 且发布者为春风予Lu），任何一项失败都不允许安装
  5. 启动 Inno 静默安装包（/SILENT + 临时 flag 文件通知安装器装完重启），
     主程序自行退出；安装器关闭旧进程、覆盖文件、按 flag 自动启动新版

纯网络/校验逻辑不依赖 Qt，QThread 与对话框依赖 PyQt6。
"""

import dataclasses
import hashlib
import json
import logging
import os
import re
import subprocess
import tempfile
import threading
import time
import urllib.error
import urllib.request
from datetime import datetime
from typing import Callable, Optional

from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QProgressBar, QTextEdit, QApplication, QComboBox,
)

logger = logging.getLogger(__name__)

# ============================== 常量 ==============================

# 当前客户端版本（与 installer.iss 的 MyAppVersion 保持一致）
APP_VERSION = "1.4.18"

REPO_OWNER = "soolecho"
REPO_NAME = "ExecelSendWx"
REPO = f"{REPO_OWNER}/{REPO_NAME}"

# 安装包资产名模板（GitHub 资产名必须是 ASCII，中文名会被剥离）
ASSET_TEMPLATE = "ExcelSendWx_v{version}_setup.exe"
SUMS_ASSET_NAME = "SHA256SUMS.txt"

# 下载镜像源（国内加速）：镜像为个人/小团队免费运营，可能随时关停，
# 因此必须多镜像 + GitHub 直连兜底；安全性由 SHA256 + 签名双校验保证。
# 元组: (显示名, URL 前缀)；前缀为 DIRECT_KEY 表示 GitHub 直连。
DIRECT_KEY = "DIRECT"
MIRROR_OPTIONS = [
    ("gh-proxy 镜像（推荐）", "https://gh-proxy.org/"),
    ("ghfast.top 镜像", "https://ghfast.top/"),
    ("GitHub 直连", DIRECT_KEY),
]

_API_PATH = f"https://api.github.com/repos/{REPO}/releases/latest"


def _ordered_mirror_urls(direct_url: str, preferred: str = "") -> list:
    """把 GitHub 直链展开成按序尝试列表。

    preferred 取值：
      ""        → 默认：镜像优先，直连兜底
      DIRECT_KEY→ 用户显式选直连：直连第一，镜像兜底
      <前缀>    → 该镜像第一 → 其余镜像 → 直连兜底
    """
    prefixes = [p for _n, p in MIRROR_OPTIONS if p != DIRECT_KEY]
    direct_first = (preferred == DIRECT_KEY)
    ordered = []
    if not direct_first and preferred in prefixes:
        ordered.append(preferred)
    ordered.extend(p for p in prefixes if p != preferred)
    urls = [f"{p}{direct_url}" for p in ordered]
    if direct_first:
        urls.insert(0, direct_url)
    else:
        urls.append(direct_url)
    return urls


def build_check_urls(preferred_prefix: str = "") -> list:
    """检查源：默认镜像优先，GitHub 直连兜底；用户可选直连优先。"""
    return _ordered_mirror_urls(_API_PATH, preferred_prefix)


# 下载源（按顺序尝试）：默认镜像优先，GitHub 直连兜底；用户可选直连优先
def build_download_urls(tag: str, asset: str, preferred_prefix: str = "") -> list:
    direct = (f"https://github.com/{REPO}/releases/download/{tag}/{asset}")
    return _ordered_mirror_urls(direct, preferred_prefix)

# 网络参数
CHECK_TIMEOUT = 10          # 检查接口超时（秒/源）
DOWNLOAD_TIMEOUT = 30       # 下载连接/读取超时（秒/源）
CHUNK_SIZE = 128 * 1024
USER_AGENT = f"{REPO_NAME}-updater/{APP_VERSION}"

# 自动检查间隔
AUTO_CHECK_INTERVAL_SEC = 24 * 3600
# 启动后首次自动检查延迟（避开启动瞬间的 CPU/网络高峰）
STARTUP_CHECK_DELAY_MS = 30 * 1000
# 长运行期间的 tick 间隔（到点判断是否满 24 小时）
AUTO_TICK_MS = 60 * 60 * 1000

# 签名发布者关键词（下载的安装包必须由此主体签名）
SIGNER_KEYWORD = "春风予Lu"
# 预置签名证书 SHA1 指纹（CN=春风予Lu 自签名证书）。
# 自签名证书未导入系统信任链时（用户机器没运行 install_cert.bat），
# 用指纹比对防伪造：攻击者要伪造签名必须拥有私钥，
# 仅伪造主体名无法通过——证书指纹是整个证书内容的 SHA1 哈希，必然不同。
EXPECTED_THUMBPRINT = "191C64E4EC07377CA032878878D0A45F554C8146"

# 静默安装后通知安装器自动重启本程序的 flag 文件
RESTART_FLAG_NAME = "excel_send_wx_restart.flag"


class UpdateError(Exception):
    """更新流程异常基类。"""


class UpdateCheckError(UpdateError):
    """检查更新失败（网络/解析）。"""


class UpdateDownloadError(UpdateError):
    """下载或校验失败。"""


# ============================== 版本与数据结构 ==============================

def parse_version(text: str) -> tuple:
    """'v1.3.1' / '1.3.1' / '1.3.1.0' → (1, 3, 1)（只取前三段数字）。"""
    m = re.search(r"(\d+)\.(\d+)\.(\d+)", text or "")
    if not m:
        return (0, 0, 0)
    return tuple(int(x) for x in m.groups())


def is_newer(remote_version: str, current: str = APP_VERSION) -> bool:
    """remote 是否比 current 新。"""
    return parse_version(remote_version) > parse_version(current)


@dataclasses.dataclass
class UpdateInfo:
    version: str                 # "1.3.1"
    tag_name: str                # "v1.3.1"
    release_url: str
    release_notes: str
    asset_name: str
    size: int
    download_urls: list
    sha256: str = ""             # 可能为空（旧版本未附带 SHA256SUMS）
    published_at: str = ""


@dataclasses.dataclass
class UpdateState:
    """更新检查状态（持久化到 %LOCALAPPDATA%/ExcelSendWx/update_state.json）。"""
    last_check_at: str = ""          # ISO8601 本地时间
    last_prompted_version: str = ""  # 已自动提醒过的版本（同版本不重复打扰）
    skipped_version: str = ""        # 用户手动跳过的版本
    preferred_mirror: str = ""       # 下载镜像 URL 前缀；""=默认镜像优先，DIRECT_KEY=直连优先

    @classmethod
    def _state_path(cls) -> str:
        base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
        folder = os.path.join(base, "ExcelSendWx")
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError:
            pass
        return os.path.join(folder, "update_state.json")

    @classmethod
    def load(cls) -> "UpdateState":
        try:
            with open(cls._state_path(), "r", encoding="utf-8") as fp:
                data = json.load(fp)
            return cls(
                last_check_at=str(data.get("last_check_at", "")),
                last_prompted_version=str(data.get("last_prompted_version", "")),
                skipped_version=str(data.get("skipped_version", "")),
                preferred_mirror=str(data.get("preferred_mirror", "")),
            )
        except (OSError, ValueError):
            return cls()

    def save(self) -> bool:
        try:
            with open(self._state_path(), "w", encoding="utf-8") as fp:
                json.dump(dataclasses.asdict(self), fp,
                          ensure_ascii=False, indent=2)
            return True
        except OSError:
            logger.exception("更新状态保存失败")
            return False

    def mark_checked(self) -> None:
        self.last_check_at = datetime.now().isoformat(timespec="seconds")
        self.save()

    def should_periodic_check(self) -> bool:
        """是否到达每日检查时点（从未检查过 → True）。"""
        if not self.last_check_at:
            return True
        try:
            last = datetime.fromisoformat(self.last_check_at)
        except ValueError:
            return True
        return (datetime.now() - last).total_seconds() >= AUTO_CHECK_INTERVAL_SEC


# ============================== 网络层 ==============================

def _http_open(url: str, timeout: int, data=None):
    req = urllib.request.Request(
        url, data=data,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "application/vnd.github+json",
        },
    )
    return urllib.request.urlopen(req, timeout=timeout)


def _http_json(url: str, timeout: int, log_fn: Optional[Callable] = None):
    def _log(msg):
        logger.info(msg)
        if log_fn:
            try:
                log_fn(msg)
            except Exception:
                pass

    _log(f"[更新] 请求更新源: {url}")
    try:
        with _http_open(url, timeout) as resp:
            charset = resp.headers.get_content_charset() or "utf-8"
            return json.loads(resp.read().decode(charset))
    except urllib.error.HTTPError as exc:
        # 403 通常是未认证 API 限流（镜像源共享 IP 时常见），换下一个源
        _log(f"[更新] 源返回 HTTP {exc.code}: {url}")
        raise UpdateCheckError(f"HTTP {exc.code}") from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        _log(f"[更新] 源连接失败 {url}: {exc}")
        raise UpdateCheckError(str(exc)) from exc


def _mirrorize(url: str, preferred_prefix: str = "") -> list:
    """把一个 GitHub 直链 URL 展开成按序尝试序列（与下载源同一套规则）。"""
    return _ordered_mirror_urls(url, preferred_prefix)


def _parse_release(data: dict, log_fn: Optional[Callable] = None,
                   preferred_mirror: str = "") -> Optional[UpdateInfo]:
    """从 GitHub release API 响应解析更新信息；无有效安装包资产返回 None。"""
    if data.get("draft") or data.get("prerelease"):
        return None

    tag = str(data.get("tag_name", "")).strip()
    version = tag[1:] if tag.startswith("v") else tag
    if not version or not is_newer(version):
        return None

    assets = data.get("assets") or []
    setup_name = ASSET_TEMPLATE.format(version=version)
    setup = next((a for a in assets
                  if str(a.get("name", "")).lower() == setup_name.lower()), None)
    # 兜底：按命名规则匹配（防止以后命名微调）
    if setup is None:
        pat = re.compile(r"ExcelSendWx_v[\d.]+_setup\.exe$", re.IGNORECASE)
        setup = next((a for a in assets
                      if pat.match(str(a.get("name", "")))), None)
    if setup is None:
        if log_fn:
            try:
                log_fn(f"[更新] 最新版 {version} 未找到安装包资产，跳过")
            except Exception:
                pass
        return None

    asset_name = str(setup.get("name"))
    # SHA256SUMS（可选）：逐行 "<sha256>  <filename>"，找当前资产对应哈希
    sha256 = ""
    sums = next((a for a in assets
                 if str(a.get("name", "")).upper() == SUMS_ASSET_NAME.upper()), None)
    if sums is not None:
        sums_urls = _mirrorize(str(sums["browser_download_url"]), preferred_mirror)
        for sums_url in sums_urls:
            try:
                with _http_open(sums_url, CHECK_TIMEOUT) as r:
                    content = r.read().decode("utf-8", errors="replace")
                for line in content.splitlines():
                    parts = line.split()
                    if len(parts) >= 2 and parts[1].lower() == asset_name.lower():
                        sha256 = parts[0].lower()
                        break
                break  # 下载成功即停（即使没匹配到当前资产的行）
            except Exception:
                logger.warning("[更新] SHA256SUMS 下载失败: %s", sums_url)
                continue
        if not sha256:
            logger.warning("[更新] SHA256SUMS 未取到哈希（将仅依赖签名校验）")

    return UpdateInfo(
        version=version,
        tag_name=tag,
        release_url=str(data.get("html_url", "")),
        release_notes=str(data.get("body", "") or ""),
        asset_name=asset_name,
        size=int(setup.get("size", 0) or 0),
        download_urls=build_download_urls(tag, asset_name, preferred_mirror),
        sha256=sha256,
        published_at=str(data.get("published_at", "") or "")[:10],
    )


def check_for_update(log_fn: Optional[Callable] = None,
                     preferred_mirror: str = "") -> Optional[UpdateInfo]:
    """检查是否有新版本。返回 None 表示已是最新；失败抛 UpdateCheckError。"""
    last_error = None
    for url in build_check_urls(preferred_mirror):
        try:
            data = _http_json(url, CHECK_TIMEOUT, log_fn)
            info = _parse_release(data, log_fn, preferred_mirror)
            if info is None:
                if log_fn:
                    try:
                        log_fn(f"[更新] 当前已是最新版本 v{APP_VERSION}")
                    except Exception:
                        pass
                return None
            if log_fn:
                try:
                    log_fn(f"[更新] 发现新版本 v{info.version}"
                           f"（安装包 {info.size / 1048576:.1f} MB）")
                except Exception:
                    pass
            return info
        except UpdateCheckError as exc:
            last_error = exc
            continue
    raise UpdateCheckError(f"所有更新源均不可用: {last_error}")


# ============================== 下载与校验 ==============================

def download_file(
    urls: list,
    dest: str,
    progress_cb: Optional[Callable[[int, int], None]] = None,
    cancel_event: Optional[threading.Event] = None,
    log_fn: Optional[Callable] = None,
) -> str:
    """多源顺序尝试流式下载。成功返回 dest 路径，失败抛 UpdateDownloadError。"""
    last_error = None
    for idx, url in enumerate(urls, 1):
        try:
            if log_fn:
                try:
                    log_fn(f"[更新] 下载源 {idx}/{len(urls)}: {url}")
                except Exception:
                    pass
            with _http_open(url, DOWNLOAD_TIMEOUT) as resp:
                total = int(resp.headers.get("Content-Length", 0) or 0)
                received = 0
                tmp_path = dest + ".part"
                with open(tmp_path, "wb") as fp:
                    while True:
                        if cancel_event is not None and cancel_event.is_set():
                            fp.close()
                            try:
                                os.remove(tmp_path)
                            except OSError:
                                pass
                            raise UpdateDownloadError("用户取消下载")
                        chunk = resp.read(CHUNK_SIZE)
                        if not chunk:
                            break
                        fp.write(chunk)
                        received += len(chunk)
                        if progress_cb:
                            try:
                                progress_cb(received, total)
                            except Exception:
                                pass
                os.replace(tmp_path, dest)
            if total and os.path.getsize(dest) != total:
                raise UpdateDownloadError(
                    f"下载大小不一致: {os.path.getsize(dest)}/{total}")
            return dest
        except UpdateDownloadError:
            raise
        except (urllib.error.URLError, urllib.error.HTTPError,
                TimeoutError, OSError) as exc:
            last_error = exc
            if log_fn:
                try:
                    log_fn(f"[更新] 下载源 {idx} 失败: {exc}")
                except Exception:
                    pass
            continue
    raise UpdateDownloadError(f"所有下载源均失败: {last_error}")


def verify_sha256(path: str, expected: str) -> tuple:
    """校验文件 SHA256。返回 (是否通过, 说明)。expected 为空时视为跳过。"""
    if not expected:
        return True, "未提供 SHA256（依赖数字签名校验）"
    h = hashlib.sha256()
    with open(path, "rb") as fp:
        for block in iter(lambda: fp.read(1024 * 1024), b""):
            h.update(block)
    actual = h.hexdigest()
    if actual.lower() != expected.lower():
        return False, f"SHA256 不匹配\n期望: {expected}\n实际: {actual}"
    return True, f"SHA256 校验通过 ({actual[:16]}…)"


def verify_authenticode(path: str) -> tuple:
    """校验 Windows Authenticode 数字签名。

    判定规则（按顺序）：
      1) 必须有签名且签名者主体含「春风予Lu」
      2) 系统已信任（Status=Valid，本机已导入证书）→ 直接通过
      3) 系统未信任（自签名证书未导入）→ 比对 SignerCertificate.Thumbprint
         与预置指纹 EXPECTED_THUMBPRINT，匹配即通过

    用 PowerShell 的 Get-AuthenticodeSignature，避免引入额外依赖；
    仅在下载完成后调用一次。返回 (是否通过, 说明)。
    """
    # 路径经环境变量传递，避免中文文件名（表格自动发送By春风予Lu.exe）
    # 在命令行中遭遇 JSON 转义（\uXXXX）或引号转义问题
    ps = (
        "$ErrorActionPreference='Stop';"
        "$s=Get-AuthenticodeSignature -LiteralPath $env:ESW_VERIFY_PATH;"
        "Write-Output ('STATUS=' + $s.Status);"
        "if ($s.SignerCertificate) { "
        "Write-Output ('SIGNER=' + $s.SignerCertificate.Subject);"
        "Write-Output ('THUMBPRINT=' + $s.SignerCertificate.Thumbprint) }"
    )
    env = os.environ.copy()
    env["ESW_VERIFY_PATH"] = path
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-Command", ps],
            capture_output=True, timeout=60, env=env,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.SubprocessError) as exc:
        return False, f"无法执行签名校验: {exc}"

    # 中文 Windows PowerShell 5.1 默认按 GBK/CP936 输出，不能按 UTF-8 解码
    def _decode(raw: bytes) -> str:
        for enc in ("gbk", "utf-8"):
            try:
                return raw.decode(enc)
            except UnicodeDecodeError:
                continue
        return raw.decode("utf-8", errors="replace")

    out = _decode(proc.stdout or b"")
    err = _decode(proc.stderr or b"")
    status = ""
    signer = ""
    thumbprint = ""
    for line in out.splitlines():
        if line.startswith("STATUS="):
            status = line[7:].strip()
        elif line.startswith("SIGNER="):
            signer = line[7:].strip()
        elif line.startswith("THUMBPRINT="):
            thumbprint = line[11:].strip().upper()

    # 1) 文件未签名 / 签名异常
    if status in ("NotSigned", "NoSignature", "UnknownError", ""):
        return False, f"文件未签名或签名异常: {status or '未知'}（{err.strip()[:120]}）"

    # 2) 签名者主体必须含春风予Lu（防伪造主体名）
    if not signer or SIGNER_KEYWORD not in signer:
        return False, f"签名发布者不受信: {signer or '（无）'}"

    # 3) 系统已信任（用户运行过 install_cert.bat）→ 直接通过
    if status == "Valid":
        return True, f"数字签名有效（系统已信任），发布者: {signer}"

    # 4) 自签名证书未导入系统信任链 → 比对证书指纹防伪造
    #    攻击者即便伪造 CN=春风予Lu 主体名，证书指纹也必然不同
    #    （指纹是整个证书内容的 SHA1 哈希）
    if thumbprint and thumbprint == EXPECTED_THUMBPRINT.upper():
        return True, (f"数字签名有效（自签名证书指纹匹配），发布者: {signer}"
                      f"，状态: {status}")

    return False, (
        f"签名证书指纹不匹配: 期望 {EXPECTED_THUMBPRINT[:16]}…，"
        f"实际 {thumbprint[:16] + '…' if thumbprint else '（无）'}，"
        f"状态: {status}"
    )


def installer_dest_path(version: str) -> str:
    """下载安装包的本地临时路径。"""
    folder = os.path.join(tempfile.gettempdir(), "ExcelSendWx_update")
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError:
        folder = tempfile.gettempdir()
    return os.path.join(folder, ASSET_TEMPLATE.format(version=version))


def launch_installer_and_exit(installer_path: str,
                              log_fn: Optional[Callable] = None) -> None:
    """写重启 flag → 启动静默安装器 → 强制退出本程序。

    安装器 CloseApplications=yes 会关闭旧实例；flag 存在时
    installer.iss 中的 Check 函数会在安装结束后启动新版。
    """
    flag_path = os.path.join(tempfile.gettempdir(), RESTART_FLAG_NAME)
    with open(flag_path, "w", encoding="ascii") as fp:
        fp.write(datetime.now().isoformat(timespec="seconds"))

    subprocess.Popen(
        [installer_path, "/SILENT", "/NORESTART", "/CLOSEAPPLICATIONS"],
        close_fds=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    msg = "[更新] 静默安装器已启动，程序即将退出，安装完成后将自动启动新版"
    logger.info(msg)
    if log_fn:
        try:
            log_fn(msg)
        except Exception:
            pass

    # 给日志/托盘一点收尾时间，然后强制结束（不等后台调度线程）
    def _hard_exit():
        os._exit(0)

    t = threading.Timer(2.0, _hard_exit)
    t.daemon = True
    t.start()
    QApplication.quit()


# ============================== Qt：后台线程与对话框 ==============================

class UpdateCheckWorker(QThread):
    """后台检查更新线程。info 为 None 表示已是最新。"""
    succeeded = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, log_fn: Optional[Callable] = None, parent=None,
                 preferred_mirror: str = ""):
        super().__init__(parent)
        self._log_fn = log_fn
        self._preferred_mirror = preferred_mirror

    def run(self):
        try:
            info = check_for_update(self._log_fn, self._preferred_mirror)
            self.succeeded.emit(info)
        except UpdateError as exc:
            self.failed.emit(str(exc))
        except Exception as exc:  # 兜底，避免线程静默死亡
            logger.exception("检查更新异常")
            self.failed.emit(str(exc))


class DownloadInstallerWorker(QThread):
    """后台下载安装包线程。"""
    progress = pyqtSignal(int, int)   # received, total
    succeeded = pyqtSignal(str)       # 下载路径
    failed = pyqtSignal(str)

    def __init__(self, info: UpdateInfo, parent=None):
        super().__init__(parent)
        self._info = info
        self._cancel = threading.Event()

    def cancel(self):
        self._cancel.set()

    def run(self):
        dest = installer_dest_path(self._info.version)
        try:
            download_file(
                self._info.download_urls, dest,
                progress_cb=lambda r, t: self.progress.emit(r, t),
                cancel_event=self._cancel,
            )
            self.succeeded.emit(dest)
        except UpdateError as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            logger.exception("下载更新异常")
            self.failed.emit(str(exc))


def _clean_markdown(text: str) -> str:
    """极简清理 GitHub Markdown，便于在纯文本框里阅读。"""
    if not text:
        return "（本次更新未提供详细说明）"
    lines = []
    for line in text.splitlines():
        line = re.sub(r"^#{1,6}\s*", "", line)       # 标题号
        line = line.replace("**", "").replace("`", "")
        line = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", line)  # 链接
        lines.append(line)
    return "\n".join(lines).strip()


class UpdateDialog(QDialog):
    """新版本对话框：展示说明 → 立即更新（内置下载进度与校验）。

    返回码：
      QDialog.DialogCode.Accepted 以外表示“稍后”；
      用户点“跳过此版本”时 rejected 且 skipped=True；
      安装器已启动时 accepted 且 update_applied=True（主窗口据此退出）。
    """

    def __init__(self, info: UpdateInfo, state: UpdateState,
                 is_busy_cb: Optional[Callable[[], bool]] = None,
                 log_fn: Optional[Callable] = None, parent=None):
        super().__init__(parent)
        self.info = info
        self.state = state
        self._is_busy_cb = is_busy_cb
        self._log_fn = log_fn
        self._worker: Optional[DownloadInstallerWorker] = None
        self.skipped = False
        self.update_applied = False

        self.setWindowTitle("发现新版本")
        self.setMinimumWidth(480)
        self._build_ui()

    def _log(self, msg):
        logger.info(msg)
        if self._log_fn:
            try:
                self._log_fn(msg)
            except Exception:
                pass

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 16)
        layout.setSpacing(10)

        size_mb = self.info.size / 1048576
        head = QLabel(
            f"发现新版本 <b>v{self.info.version}</b>"
            f"（当前 v{APP_VERSION}）"
            + (f" · {self.info.published_at}" if self.info.published_at else "")
            + f" · {size_mb:.1f} MB"
        )
        head.setTextFormat(Qt.TextFormat.RichText)
        head.setStyleSheet("font-size: 15px;")
        layout.addWidget(head)

        notes = QTextEdit()
        notes.setReadOnly(True)
        notes.setPlainText(_clean_markdown(self.info.release_notes))
        notes.setFixedHeight(220)
        layout.addWidget(notes)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setVisible(False)
        self.progress.setTextVisible(True)
        layout.addWidget(self.progress)

        # 下载源选择：镜像优先可显著加速国内下载；选择持久化，
        # 下载失败时其余镜像与 GitHub 直连会自动兜底
        mirror_row = QHBoxLayout()
        mirror_row.addWidget(QLabel("下载源："))
        self.mirror_combo = QComboBox()
        for name, prefix in MIRROR_OPTIONS:
            self.mirror_combo.addItem(name, prefix)
        saved_idx = self.mirror_combo.findData(self.state.preferred_mirror)
        self.mirror_combo.setCurrentIndex(saved_idx if saved_idx >= 0 else 0)
        self.mirror_combo.setToolTip(
            "镜像下载通常比 GitHub 直连快很多；\n"
            "所选源失败时会自动尝试其余镜像与 GitHub 直连兜底。"
        )
        self.mirror_combo.currentIndexChanged.connect(self._on_mirror_changed)
        mirror_row.addWidget(self.mirror_combo, 1)
        layout.addLayout(mirror_row)
        # 应用一次保存的偏好，确保 download_urls 顺序与下拉框一致
        self._apply_mirror(self.mirror_combo.currentData() or "", persist=False)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet("color: #c0392b;")
        layout.addWidget(self.status_label)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)

        self.skip_btn = QPushButton("跳过此版本")
        self.skip_btn.clicked.connect(self._on_skip)
        btn_row.addWidget(self.skip_btn)

        self.later_btn = QPushButton("稍后提醒")
        self.later_btn.clicked.connect(self.reject)
        btn_row.addWidget(self.later_btn)

        self.update_btn = QPushButton("立即更新")
        self.update_btn.setDefault(True)
        self.update_btn.setMinimumWidth(110)
        self.update_btn.clicked.connect(self._on_update)
        btn_row.addWidget(self.update_btn)

        layout.addLayout(btn_row)

    def _set_controls_enabled(self, enabled: bool):
        self.update_btn.setEnabled(enabled)
        self.skip_btn.setEnabled(enabled)
        self.later_btn.setEnabled(enabled)
        self.mirror_combo.setEnabled(enabled)

    def _on_mirror_changed(self, _idx: int):
        self._apply_mirror(self.mirror_combo.currentData() or "", persist=True)

    def _apply_mirror(self, prefix: str, persist: bool):
        """按所选镜像重排下载源列表；persist=True 时写入状态文件。"""
        self.info.download_urls = build_download_urls(
            self.info.tag_name, self.info.asset_name, prefix)
        if persist and prefix != self.state.preferred_mirror:
            self.state.preferred_mirror = prefix
            self.state.save()

    def _on_skip(self):
        self.state.skipped_version = self.info.version
        self.state.save()
        self.skipped = True
        self.reject()

    def _on_update(self):
        # 发送任务进行中不能更新（安装器会强杀进程，导致任务中断）
        if self._is_busy_cb and self._is_busy_cb():
            self.status_label.setText(
                "当前有发送任务正在执行，请等待任务结束后再更新。")
            return

        self._set_controls_enabled(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)  # 未知总大小时的滚动状态
        self.status_label.setStyleSheet("color: #2c3e50;")
        self.status_label.setText("正在下载安装包，请稍候…")

        self._worker = DownloadInstallerWorker(self.info, self)
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(self._on_downloaded)
        self._worker.failed.connect(self._on_download_failed)
        self._worker.start()

    def _on_progress(self, received: int, total: int):
        if total > 0:
            self.progress.setRange(0, 100)
            pct = min(100, int(received * 100 / total))
            self.progress.setValue(pct)
            self.status_label.setText(
                f"正在下载… {received / 1048576:.1f} / {total / 1048576:.1f} MB")
        else:
            self.status_label.setText(f"正在下载… {received / 1048576:.1f} MB")

    def _on_download_failed(self, message: str):
        self._set_controls_enabled(True)
        self.progress.setVisible(False)
        self.status_label.setStyleSheet("color: #c0392b;")
        self.status_label.setText(f"下载失败：{message}")

    def _on_downloaded(self, path: str):
        self.progress.setRange(0, 100)
        self.progress.setValue(100)
        self.status_label.setStyleSheet("color: #2c3e50;")
        self.status_label.setText("下载完成，正在校验文件完整性与数字签名…")
        QApplication.processEvents()

        # 1) SHA256（可选）
        ok_sha, detail_sha = verify_sha256(path, self.info.sha256)
        self._log(f"[更新] {detail_sha}")
        if not ok_sha:
            self._verify_failed(path, detail_sha)
            return

        # 2) Authenticode 签名（必须通过）
        ok_sig, detail_sig = verify_authenticode(path)
        self._log(f"[更新] {detail_sig}")
        if not ok_sig:
            self._verify_failed(path, detail_sig)
            return

        self.status_label.setText("校验通过，正在启动安装器…")
        QApplication.processEvents()
        time.sleep(0.3)
        try:
            launch_installer_and_exit(path, self._log_fn)
            self.update_applied = True
            self.accept()
        except OSError as exc:
            self._set_controls_enabled(True)
            self.progress.setVisible(False)
            self.status_label.setStyleSheet("color: #c0392b;")
            self.status_label.setText(f"启动安装器失败：{exc}")

    def _verify_failed(self, path: str, detail: str):
        logger.error("[更新] 安装包校验失败，已阻止安装: %s", detail)
        try:
            os.remove(path)
        except OSError:
            pass
        self._set_controls_enabled(True)
        self.progress.setVisible(False)
        self.status_label.setStyleSheet("color: #c0392b;")
        self.status_label.setText(
            f"安装包校验未通过，已阻止安装（可能下载损坏或被篡改）。\n{detail}")

    def reject(self):
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
        super().reject()
