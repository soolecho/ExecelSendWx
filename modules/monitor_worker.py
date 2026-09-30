"""监控轮询后台线程。

在 QThread 中运行，不直接操作 QWidget；日志与进度全部经 pyqtSignal 投递主线程
（遵循工程惯例：跨线程日志用 signal，不能直接写 GUI）。

两种模式（每个任务独立）：
- compare_enabled=True ：筛选后与 baseline 增量对比，只推送新增行，无新增则跳过
- compare_enabled=False：检测到"新的/指纹变化"的文件即推送其筛选结果

发送复用 wechat_sender 的 send_message / send_file（含弹窗守护、重试）。
"""
import glob
import logging
import os
import tempfile
import time
from datetime import datetime
from typing import Dict, Optional

from PyQt6.QtCore import QThread, pyqtSignal

from . import monitor_engine
from .monitor_config import MonitorManager, MonitorTask, _now_in_window
from .wechat_sender import WeChatSender

logger = logging.getLogger("ExcelSendWx.monitor")


class MonitorWorker(QThread):
    log = pyqtSignal(str)
    task_finished = pyqtSignal(str, int, int)  # task_name, success, failed
    run_finished = pyqtSignal()  # 一轮（manual/一次循环）完成

    def __init__(self, store: Optional[MonitorManager] = None, one_shot: bool = False,
                 poll_floor_sec: float = 30.0):
        super().__init__()
        self.store = store or MonitorManager()
        self.one_shot = one_shot
        self.poll_floor_sec = max(5.0, float(poll_floor_sec))
        self._stop = False
        self._last_run: Dict[str, float] = {}

    def stop(self) -> None:
        self._stop = True

    def run(self) -> None:
        try:
            if self.one_shot:
                self._pass()
            else:
                # 常驻轮询：每 poll_floor_sec 扫一次，按任务各自 poll_interval_min 节流
                self.emit(f"[监控] 轮询线程已启动，间隔 {self.poll_floor_sec:.0f}s")
                while not self._stop:
                    try:
                        self._pass()
                    except Exception as exc:
                        self.emit(f"[监控] 轮询异常: {exc}")
                    # 支持提前结束
                    start = time.time()
                    while not self._stop and time.time() - start < self.poll_floor_sec:
                        time.sleep(1.0)
            self.run_finished.emit()
        except Exception as exc:
            logger.exception("监控线程异常")
            self.emit(f"[监控] 线程异常退出: {exc}")
            self.run_finished.emit()

    # ------------------------------------------------------------ 一轮扫描
    def _pass(self) -> None:
        now = datetime.now()
        tasks = self.store.load_all()
        for task in tasks:
            if self._stop:
                break
            if not task.enabled:
                continue
            if not _now_in_window(task, now):
                continue
            if self.one_shot:
                self._process_task(task, log_prefix="[监控·手动]")
            else:
                last = self._last_run.get(task.id, 0.0)
                interval = max(1, task.poll_interval_min) * 60.0
                if time.time() - last >= interval:
                    self._last_run[task.id] = time.time()
                    self._process_task(task, log_prefix="[监控]")

    # ------------------------------------------------------------ 单个任务处理
    def _process_task(self, task: MonitorTask, log_prefix: str) -> None:
        if not task.watch_path or not os.path.isdir(task.watch_path):
            self.emit(f"{log_prefix} 任务「{task.name}」目录不存在: {task.watch_path}")
            return
        files = self._match_files(task)
        if not files:
            return
        success = 0
        failed = 0
        for fp in files:
            if self._stop:
                break
            try:
                outcome = self._process_file(task, fp, log_prefix)
                if outcome:
                    success += 1
                else:
                    failed += 1
            except Exception as exc:
                failed += 1
                logger.exception("监控处理文件失败 %s", fp)
                self.emit(f"{log_prefix} 处理失败「{os.path.basename(fp)}」: {exc}")
        self.task_finished.emit(task.name, success, failed)

    def _match_files(self, task: MonitorTask):
        patterns = [p.strip() for p in str(task.file_pattern or "*.*").split(";") if p.strip()]
        if not patterns:
            patterns = ["*.xlsx"]
        seen = set()
        out = []
        for pat in patterns:
            for p in glob.glob(os.path.join(task.watch_path, pat)):
                if os.path.isfile(p) and p not in seen:
                    seen.add(p)
                    out.append(p)
        out.sort(key=lambda p: os.path.getmtime(p))
        return out

    def _process_file(self, task: MonitorTask, fp: str, log_prefix: str) -> bool:
        name = os.path.basename(fp)
        # 读取并筛选
        headers, all_rows = monitor_engine.read_sheet_rows(
            fp, sheet_name=task.sheet_name,
            log_fn=lambda m: logger.info("monitor: %s", m),
        )
        rows = monitor_engine.filter_rows(
            headers, all_rows, task.filter_column, task.filter_values
        )
        key_columns = list(task.compare_columns) or ([task.filter_column] if task.filter_column else None)

        new_rows: list = rows
        if task.compare_enabled:
            baseline_keys = task.baseline.get(name, [])
            new_rows, new_keys = monitor_engine.diff_incremental(
                headers, rows, baseline_keys, key_columns
            )
            if not new_rows:
                return False  # 无新增，跳过
        else:
            # 关闭对比：文件未见过/指纹变 → 视为新文件
            if not monitor_engine.is_new_file(task.seen_files, fp):
                return False
            new_keys = monitor_engine.build_baseline(headers, rows, key_columns)

        # 有新内容 → 清洗/校验 → 按提取列裁剪 → 生成并推送
        clean_rows, dropped = monitor_engine.apply_clean_rules(
            headers, new_rows, task.clean_rules
        )
        for d in dropped:
            self.emit(f"{log_prefix} 任务「{task.name}」行{d['row_idx'] + 1} 清洗剔除（{d['column']}）: {d['reason']}")

        # 推送/生成基于清洗后的行 + 指定的提取列；文字说明基于裁剪后的结果
        view_headers, view_rows = monitor_engine.extract_columns_by(
            headers, clean_rows, task.extract_columns
        )
        texts = []
        png_path: Optional[str] = None
        out_file: Optional[str] = None
        if task.send_text:
            texts.append(monitor_engine.render_rows_text(view_headers, view_rows))
        if task.send_image:
            tmp = os.path.join(tempfile.gettempdir(),
                               f"monitor_{datetime.now().strftime('%H%M%S%f')}.png")
            png_path = monitor_engine.render_rows_image(view_headers, view_rows, tmp,
                                                        log_fn=lambda m: logger.info("monitor img: %s", m))
        if task.send_file:
            tmpx = os.path.join(tempfile.gettempdir(),
                                f"monitor_{datetime.now().strftime('%H%M%S%f')}.xlsx")
            out_file = monitor_engine.build_out_file(view_headers, view_rows, tmpx)

        self.emit(f"{log_prefix} 任务「{task.name}」文件「{name}」检出新增 {len(new_rows)} 条，清洗后保留 {len(clean_rows)} 条，开始推送…")

        delivered = 0
        for person in task.recipients:
            if self._stop or not person or not clean_rows:
                continue
            try:
                if texts:
                    self._send_text(person, "\n".join(texts), task, log_prefix)
                if png_path and os.path.exists(png_path):
                    self._send_file(person, png_path, task, log_prefix)
                if out_file and os.path.exists(out_file):
                    self._send_file(person, out_file, task, log_prefix)
                delivered += 1
            except Exception as exc:
                logger.exception("推送失败 %s", person)
                self.emit(f"{log_prefix} → 「{person}」推送失败: {exc}")

        # 更新基线 / 已见文件（无论推送成败，标记已处理避免无限重推）
        if task.compare_enabled:
            merged = set(task.baseline.get(name, []))
            merged.update(new_keys)
            task.baseline[name] = sorted(merged)
        else:
            task.seen_files[name] = monitor_engine.file_fingerprint(fp)
        self.store.save(task)

        # 清理临时图/表
        for tmp in (png_path, out_file):
            if tmp and os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass
        if delivered:
            self.emit(f"{log_prefix} 任务「{task.name}」推送完成，接收人 {delivered} 个")
        else:
            self.emit(f"{log_prefix} 任务「{task.name}」无可用接收人，跳过发送；已记录处理进度")
        return True

    def _send_text(self, person: str, text: str, task: MonitorTask, prefix: str) -> None:
        sender = WeChatSender.shared_instance()
        ok = sender.send_message(
            content=text,
            recipient=person,
            first_send=False,
            chat_delay=task.chat_delay,
            fast_mode=True,
        )
        if not ok:
            raise RuntimeError("微信文本发送返回失败")
        self.emit(f"{prefix} → 「{person}」已发送文字")

    def _send_file(self, person: str, path: str, task: MonitorTask, prefix: str) -> None:
        sender = WeChatSender.shared_instance()
        ok = sender.send_file(
            file_path=path,
            recipient=person,
            chat_delay=task.chat_delay,
            fast_mode=True,
        )
        if not ok:
            raise RuntimeError("微信文件发送返回失败")
        self.emit(f"{prefix} → 「{person}」已发送文件 {os.path.basename(path)}")

    def emit(self, msg: str) -> None:
        self.log.emit(msg)
        logger.info(msg)