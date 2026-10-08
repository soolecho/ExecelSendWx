# -*- coding: utf-8 -*-
"""无 Qt 依赖的发送执行循环。

数据发送 Tab 的 SendWorker、定时配置链 ProfileChainWorker 共用同一套
"打开聊天 → 按顺序发送 文字/图片/自定义消息/附件" 逻辑，避免两处实现漂移。

调用方负责：
- 构造 tasks：[{name, person_data, table_data, recipient, custom_msg, ...}]
- 初始化/传入 WeChatSender
- 通过回调接收日志/进度，并在结束后最小化微信窗口
"""

import os

VALID_KINDS = ("text", "image", "custom", "attachment")


def _create_steps(send_order, custom_msg, attachment):
    """按配置的发送顺序生成步骤；内容为空的步骤跳过。"""
    steps = []
    for kind in send_order or []:
        if kind == "image":
            steps.append({"type": "image", "index": 0})
        elif kind == "text":
            steps.append({"type": "text", "index": 0})
        elif kind == "custom" and custom_msg:
            steps.append({"type": "custom", "index": 0})
        elif kind == "attachment" and attachment:
            steps.append({"type": "attachment", "index": 0})
    return steps


def run_send_tasks(
    tasks,
    sender,
    *,
    send_order,
    attachment,
    send_interval,
    chat_delay,
    log,
    stop_event,
    pause_event=None,
    progress=None,
    label="",
):
    """执行发送任务列表。

    返回 (success_count, failed_tasks)。failed_tasks 为未完成的 task dict。
    线程安全前提：本函数在 worker 线程内执行，log/progress 回调必须自身
    线程安全（Qt 信号或调度器转发），不能直接操作 GUI。
    """
    def _log(msg):
        try:
            log(msg)
        except Exception:
            pass

    def _progress(current, total):
        if progress is None:
            return
        try:
            progress(current, total)
        except Exception:
            pass

    # 去重保序，只保留合法类型
    order = []
    for kind in send_order or []:
        if kind in VALID_KINDS and kind not in order:
            order.append(kind)

    attach_path = str(attachment or "").strip()
    if attach_path and not os.path.exists(attach_path):
        _log(f"⚠ 附加文件不存在，本次发送不带附件: {attach_path}")
        attach_path = ""

    normalized = []
    for raw in tasks:
        if isinstance(raw, dict):
            # 浅拷贝：失败重试时不污染调用方的 dict（table_data 等大对象共享）
            task = dict(raw)
        else:
            name, person_data, table_data, recipient, custom_msg = raw
            task = {
                "name": name,
                "person_data": person_data,
                "table_data": table_data,
                "recipient": recipient,
                "custom_msg": custom_msg,
            }
        # 已带 pending_steps 说明是失败后续发：保留剩余步骤及断点 index
        if not task.get("pending_steps"):
            task["pending_steps"] = _create_steps(
                order, task.get("custom_msg", ""), attach_path
            )
        # 每次进入新的发送循环都要重新打开聊天窗口
        task["_chat_opened"] = False
        normalized.append(task)
    tasks = normalized

    success_count = 0
    failed_tasks = []
    total_count = len(tasks)

    for i, task in enumerate(tasks):
        if stop_event is not None and stop_event.is_set():
            _log("⏹ 发送已停止")
            failed_tasks.extend(tasks[i:])
            break

        if pause_event is not None:
            import time as _time
            while pause_event.is_set() and not (
                stop_event is not None and stop_event.is_set()
            ):
                if stop_event is not None:
                    stop_event.wait(0.1)
                else:
                    _time.sleep(0.1)
        if stop_event is not None and stop_event.is_set():
            _log("⏹ 发送已停止")
            failed_tasks.extend(tasks[i:])
            break

        name = task["name"]
        person_data = task["person_data"]
        table_data = task["table_data"]
        recipient = task["recipient"]
        prefix = f"[{label}] " if label else ""

        try:
            if i == 0:
                _log("等待微信就绪...")
                if stop_event is not None and stop_event.wait(0.2):
                    failed_tasks.extend(tasks[i:])
                    break

            _log(f"{prefix}[{i+1}/{total_count}] 正在发送给 {recipient} ({name})...")
            _progress(i + 1, total_count)

            messages = []
            max_message_length = 2000
            for j in range(0, len(person_data), max_message_length):
                messages.append(person_data[j:j + max_message_length])

            while task["pending_steps"] and not (
                stop_event is not None and stop_event.is_set()
            ):
                # 每个收件人只在第一个 step 前切换一次聊天窗口
                if not task.get("_chat_opened"):
                    if not sender.open_chat(
                        recipient,
                        chat_delay=chat_delay,
                        stop_event=stop_event,
                    ):
                        _log(
                            f"{prefix}[{i+1}/{total_count}] ❌ 打开聊天窗口失败: {recipient}"
                        )
                        break
                    task["_chat_opened"] = True

                steps = task["pending_steps"]
                step_types = [s["type"] for s in steps]
                has_text = "text" in step_types
                has_custom = "custom" in step_types
                has_files = ("image" in step_types) or (
                    "attachment" in step_types and bool(attach_path)
                )

                if has_files:
                    # ---------- v1.5.0 批处理合并：文字/自定义文字 + 图片 + 附件 ----------
                    # 合并为一次 send_text_and_files（一次粘贴 + 一次 Enter），绝不逐条发；
                    # 仅当同一内容同时含 text 与 custom 两种文字时才拆两次：
                    # text+图片+附件一批，custom 单独一批。
                    images = []
                    try:
                        if "image" in step_types:
                            images = list(
                                sender._create_table_images(task["table_data"]) or []
                            )
                            for _p in images:
                                _log(
                                    f"{prefix}[{i+1}/{total_count}] 已生成表格图片: {os.path.basename(_p)}"
                                )
                    except Exception as exc:
                        _log(
                            f"{prefix}[{i+1}/{total_count}] ❌ 生成表格图片失败: {exc}"
                        )
                        for _p in images:
                            try:
                                sender._remove_temp_image(_p)
                            except Exception:
                                pass
                        break
                    file_paths = images + (
                        [attach_path]
                        if "attachment" in step_types and attach_path
                        else []
                    )
                    # 发送顺序联动粘贴顺序：按 send_order 中「文字类（text/custom）」
                    # 与「文件类（image/attachment）」的相对先后，决定两步粘贴谁先谁后
                    # （文字在前 → 先粘文字再粘文件；文件在前 → 先粘文件再追加文字），
                    # 以及图片/附件卡片之间谁在前（一次 CF_HDROP 粘贴，卡片顺序=列表顺序）。
                    order_kinds = [
                        k for k in (send_order or [])
                        if k in ("text", "custom", "image", "attachment")
                    ]
                    text_first = bool(order_kinds) and order_kinds[0] in ("text", "custom")
                    if ("attachment" in order_kinds and "image" in order_kinds
                            and order_kinds.index("attachment") < order_kinds.index("image")):
                        file_paths = file_paths[len(images):] + file_paths[:len(images)]
                    # 合并文字：text 形式优先（与 custom 并存时先发 text 批次）；
                    # 仅 custom 时用自定义文字；都没有则纯文件批。
                    if has_text:
                        merge_text = "".join(messages)
                    elif has_custom:
                        merge_text = task["custom_msg"]
                    else:
                        merge_text = None

                    if merge_text is not None:
                        _log(
                            f"{prefix}[{i+1}/{total_count}] 合并发送文字+{len(file_paths)}个文件给 {recipient}"
                        )
                        success = sender.send_text_and_files(
                            merge_text,
                            file_paths,
                            recipient,
                            chat_delay=chat_delay,
                            fast_mode=True,
                            stop_event=stop_event,
                            text_first=text_first,
                        )
                    else:
                        _log(
                            f"{prefix}[{i+1}/{total_count}] 合并发送{len(file_paths)}个文件给 {recipient}"
                        )
                        success = sender.send_files_batch(
                            file_paths,
                            recipient,
                            chat_delay=chat_delay,
                            fast_mode=True,
                            stop_event=stop_event,
                        )
                    # 清理表格图片临时文件（附件是用户文件，不删）
                    for _p in images:
                        try:
                            sender._remove_temp_image(_p)
                        except Exception:
                            pass
                    if not success:
                        _log(
                            f"{prefix}[{i+1}/{total_count}] ❌ 合并发送失败，整批保留待重发"
                        )
                        break
                    # 消费整批：仅 text 与 custom 并存时留下 custom 单独发
                    if has_text and has_custom:
                        task["pending_steps"] = [
                            s for s in steps if s["type"] == "custom"
                        ]
                    else:
                        task["pending_steps"] = []
                else:
                    # ---------- 无图片/附件：逐项发送（原逻辑） ----------
                    step = steps[0]
                    step_type = step["type"]
                    start_index = step.get("index", 0)
                    success = False

                    if step_type == "image":
                        success, next_index = sender.send_table_images_progress(
                            table_data,
                            recipient,
                            chat_delay=chat_delay,
                            start_index=start_index,
                            stop_event=stop_event,
                            fast_mode=True,
                        )
                        step["index"] = next_index
                    elif step_type == "text":
                        success, next_index = sender.send_multiple_messages_progress(
                            messages,
                            recipient,
                            chat_delay=chat_delay,
                            start_index=start_index,
                            stop_event=stop_event,
                            fast_mode=True,
                        )
                        step["index"] = next_index
                    elif step_type == "attachment":
                        _log(f"{prefix}[{i+1}/{total_count}] 发送附件给 {recipient}")
                        success = sender.send_file(
                            attach_path,
                            recipient,
                            chat_delay=chat_delay,
                            fast_mode=True,
                            stop_event=stop_event,
                        )
                        if success:
                            _log(f"{prefix}[{i+1}/{total_count}] 已发送附加文件")
                    else:  # custom
                        _log(f"{prefix}[{i+1}/{total_count}] 发送自定义消息给 {recipient}")
                        success = sender.send_message(
                            task["custom_msg"],
                            recipient,
                            chat_delay=chat_delay,
                            fast_mode=True,
                            stop_event=stop_event,
                        )
                        if success:
                            _log(f"{prefix}[{i+1}/{total_count}] 已发送自定义消息")

                    if not success:
                        break

                    task["pending_steps"].pop(0)

                if task["pending_steps"] and stop_event is not None:
                    if stop_event.wait(0.1):
                        break

            if not task["pending_steps"]:
                _log(f"{prefix}[{i+1}/{total_count}] ✅ 成功发送给 {recipient}")
                success_count += 1
            else:
                if stop_event is not None and stop_event.is_set():
                    _log(
                        f"{prefix}[{i+1}/{total_count}] ⏹ 已停止，保留未完成任务: {recipient}"
                    )
                else:
                    _log(f"{prefix}[{i+1}/{total_count}] ❌ 发送失败: {recipient}")
                failed_tasks.append(task)

        except Exception as exc:
            _log(f"{prefix}[{i+1}/{total_count}] ❌ 发送异常: {name} - {exc}")
            failed_tasks.append(task)

        if stop_event is not None and stop_event.is_set():
            failed_tasks.extend(tasks[i + 1:])
            _log("⏹ 发送已停止，剩余任务已保留")
            break

        if i < total_count - 1 and send_interval > 0 and stop_event is not None:
            if stop_event.wait(send_interval):
                failed_tasks.extend(tasks[i + 1:])
                _log("⏹ 发送已停止，剩余任务已保留")
                break

    return success_count, failed_tasks
