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
import threading
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

    # 任务级互斥：同一任务可能同时被「常驻轮询」与「手动执行」两个 worker 处理，
    # 用 busy 标志保证同一时刻只有一个 worker 在处理同一任务，避免重复推送
    _TASK_BUSY: Dict[str, bool] = {}
    _TASK_BUSY_LOCK = threading.Lock()

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
        # 任务级互斥：该任务正在被其他 worker（常驻轮询/手动执行）处理则本轮跳过
        with self._TASK_BUSY_LOCK:
            if self._TASK_BUSY.get(task.id, False):
                self.emit(f"{log_prefix} 任务「{task.name}」正在执行中（手动/轮询），本轮跳过")
                return
            self._TASK_BUSY[task.id] = True
        try:
            self._process_task_inner(task, log_prefix)
        finally:
            with self._TASK_BUSY_LOCK:
                self._TASK_BUSY.pop(task.id, None)

    def _process_task_inner(self, task: MonitorTask, log_prefix: str) -> None:
        if not task.watch_path or not os.path.isdir(task.watch_path):
            self.emit(f"{log_prefix} 任务「{task.name}」目录不存在: {task.watch_path}")
            return
        files = self._match_files(task)
        if not files:
            return
        # 监控目录可能堆积大量历史文件：每个轮询周期只处理「时间最新」的那一个。
        # 开启对比时，基于已记录的 baseline 做行增量（等同与上一个时间点的内容对比），
        # 避免首次/每次轮询把目录里的历史表全部读一遍并重复推送。
        latest = files[-1]  # _match_files 已按 mtime 升序，最后一个是最新
        success = 0
        failed = 0
        try:
            outcome = self._process_file(task, latest, log_prefix)
            if outcome is None:
                # 无新增/文件未变：_process_file 已打出原因日志，不发完成信号
                # （避免每轮「成功 0 / 失败 0」噪音，无新增不涉及发送结果）
                return
            # outcome = (成功接收人数, 失败接收人数)
            success += outcome[0]
            failed += outcome[1]
        except Exception as exc:
            failed += 1
            logger.exception("监控处理文件失败 %s", latest)
            self.emit(f"{log_prefix} 处理失败「{os.path.basename(latest)}」: {exc}")
        self.task_finished.emit(task.name, success, failed)

    def _match_files(self, task: MonitorTask):
        patterns = [p.strip() for p in str(task.file_pattern or "*.*").split(";") if p.strip()]
        if not patterns:
            patterns = ["*.xlsx"]
        seen = set()
        out = []
        recursive = bool(getattr(task, "include_subdir", False))
        for pat in patterns:
            search = pat
            if recursive and not search.startswith("**/"):
                search = "**/" + pat
            for p in glob.glob(os.path.join(task.watch_path, search),
                               recursive=recursive):
                if not os.path.isfile(p) or p in seen:
                    continue
                base = os.path.basename(p)
                # 跳过 Excel 打开编辑时的临时锁文件（~$xxx.xlsx）与临时备份（~xxx.tmp）
                if base.startswith("~$") or (base.startswith("~") and base.lower().endswith(".tmp")):
                    continue
                # 文件名前缀过滤：非空时只处理以该前缀开头的文件（如"主干及分支"）
                if getattr(task, "file_prefix", "") and not base.startswith(
                        str(task.file_prefix).strip()):
                    continue
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
            # 语义：只跟「上一个时间点的文件」对比（滚动快照），找出相对上一份新增的行。
            # 因此基线只用最近一份文件的行键，而非累积所有历史文件。
            prev_keys = (task.baseline or {}).get("__prev__", [])
            new_rows, new_keys = monitor_engine.diff_incremental(
                headers, rows, prev_keys, key_columns
            )
            if not new_rows:
                # 相对上一份无新增：打日志说明原因，避免"无新增"被误认为发送失败
                cols = "、".join(key_columns) if key_columns else "全部列"
                self.emit(f"{log_prefix} 任务「{task.name}」文件「{name}」无新增数据：与上一份文件对比（对比列: {cols}）无新行，跳过")
                return None  # 无新增，跳过（不算成功也不算失败）
        else:
            # 关闭对比：文件未见过/指纹变 → 视为新文件
            if not monitor_engine.is_new_file(task.seen_files, fp):
                # 文件未变化：打日志说明原因，避免被误认为发送失败
                self.emit(f"{log_prefix} 任务「{task.name}」文件「{name}」文件未变化（大小/修改时间相同），跳过")
                return None  # 文件未变化，无新内容，跳过（不算失败）
            new_keys = monitor_engine.build_baseline(headers, rows, key_columns)

        # 有新内容 → 清洗/校验 → 按提取列裁剪 → 生成并推送
        clean_rows, dropped = monitor_engine.apply_clean_rules(
            headers, new_rows, task.clean_rules
        )
        for d in dropped:
            self.emit(f"{log_prefix} 任务「{task.name}」行{d['row_idx'] + 1} 清洗剔除（{d['column']}）: {d['reason']}")

        # 清洗后条数判断拦截：若 len(clean_rows) OP limit_count 命中，则本次不发送。
        # 基线照常推进（等同"已处理"），避免同批数据被反复拦截刷日志。
        if task.limit_enabled and task.limit_op and self._limit_hit(task, len(clean_rows)):
            self.emit(f"{log_prefix} 任务「{task.name}」文件「{name}」清洗后 {len(clean_rows)} 条命中条数判断（{task.limit_op} {task.limit_count}），本次不发送")
            self._advance_baseline(task, fp, headers, rows, key_columns)
            return None

        # 推送/生成基于清洗后的行 + 指定的提取列；文字说明基于裁剪后的结果
        view_headers, view_rows = monitor_engine.extract_columns_by(
            headers, clean_rows, task.extract_columns
        )
        # 自定义备注列：非空时在提取列每行末尾追加一列固定文案（所有行相同，紧随提取列之后），
        # 作用于文字/图片/文件/智能表格四个共享 view_rows 的输出。
        remark = getattr(task, "remark_text", "") or ""
        remark = remark.strip()
        if remark:
            view_headers = [*view_headers, "备注"]
            view_rows = [list(r) + [remark] for r in view_rows]

        # 指定区域截图（仅「关闭对比」模式生效，与筛选/清洗/提取列并存、独立勾选）：
        # 勾选后图片内容改为截取 sheet 原始区域（不参与筛选/清洗），留空范围 = 整表已用区域。
        # 开启对比时忽略本功能并给出提示，图片仍按筛选数据渲染。
        snapshot_headers: Optional[List[str]] = None
        snapshot_rows: Optional[List[List[str]]] = None
        if getattr(task, "snapshot_enabled", False):
            if not task.compare_enabled:
                grid = [list(headers)] + [list(r) for r in all_rows]  # 整表已用区域（含表头行）
                region = monitor_engine.parse_region(getattr(task, "snapshot_range", ""))
                if region:
                    region_rows = monitor_engine.slice_region(grid, *region)
                    self.emit(f"{log_prefix} 任务「{task.name}」指定区域截图: {task.snapshot_range}")
                else:
                    region_rows = grid
                    self.emit(f"{log_prefix} 任务「{task.name}」指定区域截图: 整表已用区域")
                if region_rows:
                    snapshot_headers, snapshot_rows = region_rows[0], region_rows[1:]
            else:
                self.emit(f"{log_prefix} 任务「{task.name}」已勾选指定区域截图，但当前为开启对比模式，本轮忽略，使用筛选图片")

        texts = []
        png_path: Optional[str] = None
        out_file: Optional[str] = None
        if task.send_text:
            texts.append(monitor_engine.render_rows_text(
                view_headers, view_rows,
                title=task.text_title,
                show_detail=task.text_detail,
            ))
        if task.send_image:
            tmp = os.path.join(tempfile.gettempdir(),
                               f"monitor_{datetime.now().strftime('%H%M%S%f')}.png")
            if snapshot_headers is not None:
                # 指定区域截图优先：图片 = 用户勾选的区域原样截取
                png_path = monitor_engine.render_rows_image(
                    snapshot_headers, snapshot_rows, tmp,
                    log_fn=lambda m: logger.info("monitor img: %s", m))
            else:
                png_path = monitor_engine.render_rows_image(view_headers, view_rows, tmp,
                                                            log_fn=lambda m: logger.info("monitor img: %s", m))
        if task.send_file:
            tmpx = os.path.join(tempfile.gettempdir(),
                                f"monitor_{datetime.now().strftime('%H%M%S%f')}.xlsx")
            out_file = monitor_engine.build_out_file(view_headers, view_rows, tmpx)

        # 可选出口：写入智能表格（AirScript webhook）。纯 HTTP，不占用微信发送批次锁。
        # 将清洗/提取后的行（去掉首行标题）追加写入用户自己的金山智能表格指定 sheet 指定列。
        airsync_failed = False
        if task.airsync_enabled and clean_rows:
            # 同一批重试去重：上次 airsync 已成功写入的行键与本次 new_keys 一致时，
            # 本轮是「微信发送失败保留基线」后的重试 → 跳过重复写入，仅重发微信。
            written_keys = list(task.airsync_written_keys or [])
            if written_keys and set(written_keys) == set(new_keys):
                self.emit(f"{log_prefix} 任务「{task.name}」本轮新增已写入过智能表格，跳过重复写入（仅重试微信）")
            else:
                try:
                    ok, detail = monitor_engine.airsync_append(
                        task.airsync_webhook, task.airsync_token,
                        task.airsync_sheet, task.airsync_start_col,
                        task.airsync_col_count, view_rows,
                    )
                    if ok:
                        # 记录已写标记并立即落盘：失败重试时据此跳过 airsync；成功后随基线清空
                        task.airsync_written_keys = list(new_keys)
                        self.store.save(task)
                        self.emit(f"{log_prefix} 任务「{task.name}」已写入智能表格（{detail}）")
                    else:
                        airsync_failed = True
                        self.emit(f"{log_prefix} 任务「{task.name}」写入智能表格失败: {detail}")
                except Exception as exc:
                    airsync_failed = True
                    logger.exception("写入智能表格失败 %s", task.name)
                    self.emit(f"{log_prefix} 任务「{task.name}」写入智能表格失败: {exc}")

        any_failed = False  # 微信推送 或 智能表格写入 任一失败即视为本次未完成 → 保留基线重试

        self.emit(f"{log_prefix} 任务「{task.name}」文件「{name}」检出新增 {len(new_rows)} 条，清洗后保留 {len(clean_rows)} 条，开始推送…")

        # 全局发送批次锁：与其他发送任务（手动/定时/链路）排队串行，
        # 避免多个 worker 同时驱动同一份微信句柄导致窗口争抢发错对象。
        # 排队无时间上限；队列期间被停止则返回 None（跳过本轮）。
        acquired = WeChatSender.acquire_batch(
            log_fn=lambda m: self.emit(m),
            should_stop=lambda: self._stop,
        )
        if not acquired:
            self.emit(f"{log_prefix} 任务「{task.name}」已停止，取消排队")
            return None
        msg_delivered = 0
        msg_failed = 0
        # v1.5.0 批处理合并：文字/图片/附件合并为一次粘贴 + 一次 Enter 发送（text+files 混合
        # 走两步粘贴、一次 Enter），绝不逐条发送；与本轮内容同批推送，保持消息完整。
        # 粘贴顺序按任务配置 send_order（文字/图片/文件）联动：文字在前 → 先粘文字再粘文件。
        text = "\n".join(texts) if texts else None
        files, text_first = self._resolve_send_order(task, png_path, out_file)
        try:
            for person in task.recipients:
                if self._stop or not person or not clean_rows:
                    continue
                try:
                    self._send_merged(person, text, files, task, log_prefix, text_first)
                    msg_delivered += 1
                except Exception as exc:
                    msg_failed += 1
                    logger.exception("推送失败 %s", person)
                    self.emit(f"{log_prefix} → 「{person}」推送失败: {exc}")
        finally:
            # 释放全局发送批次锁（推送结束/异常都必须放锁，避免后续任务死等）
            try:
                WeChatSender.release_batch()
            except Exception:
                pass
            # 发送完成后按配置最小化微信（隐私保护），避免窗口长时间停留前台/全屏
            if task.minimize_after and (msg_delivered or msg_failed):
                try:
                    WeChatSender.shared_instance().minimize_window()
                except Exception:
                    pass
        # 微信推送 或 智能表格写入 任一失败即视为本次未完成 → 保留基线重试
        any_failed = bool(msg_failed or airsync_failed)

        # 微信全部发送成功时，清空 airsync 已写标记（该批已完全处理，后续新增行重新正常写入）
        if not any_failed and task.airsync_written_keys:
            task.airsync_written_keys = []

        # 更新基线 / 已见文件：只有发送全部成功才推进，否则保留上一基线，
        # 让下次轮询重新检出同一批新增并重试（避免"有新增但发送失败"被当作已处理而永久丢失）
        if task.compare_enabled:
            if any_failed:
                # 存在发送/写入失败：不更新 __prev__ 基线 → 下次轮询重新检出新增行重试
                self.emit(f"{log_prefix} 任务「{task.name}」存在失败（微信/写入智能表格），保留对比基线，下次轮询将重试推送")
            else:
                # 只保留「当前这份文件的全部行键」作为下一轮对比基线（滚动快照，只对比上一个时间点）。
                # 注意必须存当前文件所有行的键，而非仅 new_keys（新增行），否则下轮会重复推上轮已存在的行。
                task.baseline = {"__prev__": monitor_engine.build_baseline(headers, rows, key_columns)}
        else:
            if not any_failed:
                task.seen_files[name] = monitor_engine.file_fingerprint(fp)
            else:
                self.emit(f"{log_prefix} 任务「{task.name}」存在失败（微信/写入智能表格），不标记已见文件，下次轮询将重试推送")
        self.store.save(task)

        # 清理临时图/表：延迟 15s 再删。wxauto 的 SendFiles 返回时微信只是开始异步上传，
        # 立即删除会导致上传读到一半文件消失而发送失败。
        # 注意：threading.Timer 不接受 daemon 直接关键字（daemon 只支持经 kwargs 传入，
        # 且仅 Python 3.10+），直接传 daemon=True 会抛 TypeError 导致任务被误判失败，
        # 因此这里不传 daemon（Python 3.10+ Timer 默认即 daemon）。
        def _delayed_remove(path: str) -> None:
            def _rm() -> None:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                        logger.info("监控临时文件已删除: %s", path)
                        # 已删除提示（GUI 面板日志；Timer 线程 emit 线程安全）
                        try:
                            self.emit(f"[监控] 临时文件已删除: {os.path.basename(path)}")
                        except Exception:
                            pass
                except OSError as exc:
                    logger.warning("删除监控临时文件失败 %s: %s", path, exc)
            threading.Timer(15.0, _rm).start()

        # 清理环节异常（如定时器创建失败）只记录日志，不得中断/误判本轮发送结果
        try:
            for tmp in (png_path, out_file):
                if tmp and os.path.exists(tmp):
                    _delayed_remove(tmp)
        except Exception as exc:
            logger.warning("清理监控临时文件失败（不影响发送结果）: %s", exc)
        if msg_delivered:
            self.emit(f"{log_prefix} 任务「{task.name}」推送完成，接收人 {msg_delivered} 个")
        elif msg_failed:
            self.emit(f"{log_prefix} 任务「{task.name}」所有接收人推送均失败")
        else:
            self.emit(f"{log_prefix} 任务「{task.name}」无可用接收人，跳过发送；已记录处理进度")
        return (msg_delivered, msg_failed)

    @staticmethod
    def _limit_hit(task: MonitorTask, n: int) -> bool:
        """清洗后条数判断：len(clean_rows) OP limit_count 为真即拦截（本次不发送）。"""
        op = task.limit_op
        c = task.limit_count
        if op == ">":
            return n > c
        if op == "<":
            return n < c
        if op == "==":
            return n == c
        return False

    def _advance_baseline(self, task: MonitorTask, fp: str,
                          headers, rows, key_columns) -> None:
        """把当前文件推进为"已处理"基线（与发送成功语义一致），供条数判断拦截后调用。"""
        if task.compare_enabled:
            task.baseline = {"__prev__":
                             monitor_engine.build_baseline(headers, rows, key_columns)}
        else:
            task.seen_files[os.path.basename(fp)] = monitor_engine.file_fingerprint(fp)
        self.store.save(task)

    @staticmethod
    def _resolve_send_order(task, png_path, out_file):
        """按任务配置的发送顺序（text/image/attachment）解析本轮粘贴方案。

        返回 (files, text_first)：
        - files：图片/附件按顺序排成的文件列表（一次 CF_HDROP 粘贴，卡片顺序=列表顺序）
        - text_first：文字类排最前 → 先粘文字再粘文件（文字在前）；否则文件在前
        send_order 缺失或非法时回退默认（文字在前，然后图片、文件），兼容旧配置。
        """
        order = list(getattr(task, "send_order", None)
                     or ["text", "image", "attachment"])
        files_map = {}
        if png_path and os.path.exists(png_path):
            files_map["image"] = png_path
        if out_file and os.path.exists(out_file):
            files_map["attachment"] = out_file
        kinds = [k for k in order if k in files_map]
        for k in ("image", "attachment"):
            if k in files_map and k not in kinds:
                kinds.append(k)
        files = [files_map[k] for k in kinds]
        text_first = bool(order) and order[0] == "text"
        return files, text_first

    def _send_merged(self, person: str, text, files, task: MonitorTask,
                     prefix: str, text_first: bool = False) -> None:
        """合并推送：文字/图片/附件一次粘贴 + 一次 Enter（v1.5.0 批处理合并约束）。

        - text+files → send_text_and_files（两步粘贴：文字/文件先后由 text_first 决定）
        - 仅 files → send_files_batch（一次粘贴 + Enter）
        - 仅 text → send_message（原逻辑）
        发送返回 False 视为失败（抛异常由调用方计入 msg_failed）。
        """
        sender = WeChatSender.shared_instance()
        if files and text:
            ok = sender.send_text_and_files(
                text=text, file_paths=files, recipient=person,
                chat_delay=task.chat_delay, fast_mode=True,
                text_first=text_first,
            )
        elif files:
            ok = sender.send_files_batch(
                file_paths=files, recipient=person,
                chat_delay=task.chat_delay, fast_mode=True,
            )
        else:
            ok = sender.send_message(
                content=text or "", recipient=person, first_send=False,
                chat_delay=task.chat_delay, fast_mode=True,
            )
        if not ok:
            raise RuntimeError("微信合并发送返回失败")
        names = [os.path.basename(f) for f in files]
        if text and names:
            self.emit(f"{prefix} → 「{person}」已发送文字+{len(names)}个文件")
        elif names:
            self.emit(f"{prefix} → 「{person}」已发送{len(names)}个文件")
        else:
            self.emit(f"{prefix} → 「{person}」已发送文字")

    def emit(self, msg: str) -> None:
        self.log.emit(msg)
        logger.info(msg)