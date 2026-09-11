---
name: "pyqt-collapsible-chip-panels"
description: "PyQt6/PyQt5 chip-button collapsible panels with jelly animation, elastic auto-resizing window, accordion/multi-open modes, and auto-shrunk rail columns. Invoke when building PyQt desktop UIs needing collapsible sections, compact dashboards, or show/hide panel interactions with smooth animation."
---

# PyQt 芯片折叠面板 + 果冻动画 + 弹性窗口（工作区组件模式）

适用于 PyQt6（PyQt5 仅需改 enum 命名空间）。已在 wxauto 项目实战验证，解决四类问题：
分组占空间、折叠/展开生硬、窗口尺寸不跟随内容、全收起后大块空背板。

## 核心组件（模板见本文件末尾）

1. **`ChipSection(QWidget)`**：无标题面板，内容动画 `QPropertyAnimation(content, b"maximumHeight")`。
2. **`ChipBar(QWidget)` + `FlowLayout(QLayout)`**：一排胶囊芯片按钮；支持
   - `exclusive=False` 多开模式；
   - `exclusive=True` 手风琴模式（展开一个自动收起其他）。
3. **MainWindow `fit_to_content(expanding, animate_window)`**：窗口几何跟随内容弹性缩放。
4. **空栏导轨**：某栏（QSplitter 子页）所有面板收起时，用 `QVariantAnimation`
   把该栏宽度平滑收到 ~126px，腾出的宽度按记忆比例分给展开栏。

## 必须遵守的 7 条工程规则（踩坑总结）

1. **动画对象持久复用，绝不每次新建**。`__init__` 建一个 `QPropertyAnimation`，
   切换时 `stop()` → 重设 start/end/duration/curve → `start()`。
   反例：同步连点时 `new + deleteLater` 堆积会触发 Qt 原生崩溃（0xC0000409）。
2. **即时路径与动画路径并发**：程序调用 `set_collapsed()`（精简模式/自动展开）
   前必须先 `stop()` 旧动画，否则旧动画 finished 回调把内容可见状态写反。
   `stop()` 不触发 finished，可安全阻断回调。
3. **反转续动起点**：展开从 0 起；折叠从当前实际高度起（`maximumHeight>=16777215`
   表示未限高，取 `height() or sizeHint().height()`）。
4. **窗口屏幕约束用框架几何，不是客户区几何**：标题栏高度 =
   `geo.top() - frameGeometry().top()`（注意是正 31，符号别写反），
   约束时扣掉四边框架开销，否则内容近满屏时标题栏被顶出屏幕（T=-31），
   表现为"最小化/最大化/关闭按钮消失"。
5. **切 tab / 首次显示按 sizeHint（理想尺寸）适配**，不要按 minimumSizeHint，
   否则高内容页（如多行编辑框）文字被压扁。
6. **程序化填充表单必须加守卫**（如 `self._loading_form`），避免加载数据时
   触发脏检测/自动展开；首次数据加载完成前 `_form_ready=False`。
7. **手风琴排中程序自动展开要带 `accordion_close=False`**，否则用户在 A 面板
   打字触发脏检测弹操作区时，A 面板会被手风琴挤掉。

## 动画曲线（柔和"果冻/允吸"手感）

- 展开：`OutBack` + `setOvershoot(1.05)`（默认 1.70 太猛），420ms；
- 折叠：`InQuart`（吸入）260ms，或 `InOutCubic` 280ms；
- 窗口/栏宽跟随：`OutQuart` 300ms。

## 组件代码模板

```python
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLayout
from PyQt6.QtCore import (Qt, pyqtSignal, QTimer, QEvent, QSize, QRect, QPoint,
                          QPropertyAnimation, QEasingCurve, QAbstractAnimation,
                          QVariantAnimation)


class ChipSection(QWidget):
    collapsedChanged = pyqtSignal(bool)

    def __init__(self, parent=None, collapsed=False):
        super().__init__(parent)
        self._collapsed = False
        self._chip = None
        self._no_accordion_close = False
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        self._content = QWidget()
        lay = QVBoxLayout(self._content)
        lay.setContentsMargins(8, 6, 8, 8)
        lay.setSpacing(5)
        outer.addWidget(self._content)
        # 规则1：单一持久动画
        self._anim = QPropertyAnimation(self._content, b"maximumHeight", self)
        self._anim.finished.connect(self._on_anim_finished)
        if collapsed:
            self.set_collapsed(True)

    def contentLayout(self):
        return self._content.layout()

    def is_collapsed(self):
        return self._collapsed

    def _kill_anim(self):  # 规则2
        self._anim.stop()
        self._content.setMaximumHeight(16777215)

    def set_collapsed(self, collapsed, animate=False, accordion_close=True):
        collapsed = bool(collapsed)
        if collapsed == self._collapsed:
            return
        self._no_accordion_close = not accordion_close and not collapsed
        if animate:
            self._animate(collapsed)
            return
        self._kill_anim()
        self._collapsed = collapsed
        self._content.setVisible(not collapsed)
        self.collapsedChanged.emit(collapsed)
        self._notify_window(not collapsed)

    def _animate(self, collapsed):
        content, anim = self._content, self._anim
        if anim.state() == QAbstractAnimation.State.Running:
            anim.stop()
        cur = max(0, content.maximumHeight())
        cur = (content.height() or content.sizeHint().height()) if cur >= 16777215 \
            else cur
        if not collapsed and content.maximumHeight() >= 16777215:
            cur = 0  # 规则3
        self._collapsed = collapsed
        self.collapsedChanged.emit(collapsed)
        content.setVisible(True)
        content.setMaximumHeight(cur)
        anim.setStartValue(cur)
        if collapsed:
            anim.setEndValue(0); anim.setDuration(260)
            anim.setEasingCurve(QEasingCurve.Type.InQuart)
        else:
            target = content.sizeHint().height()
            if target <= 0:  # 尺寸未就绪：降级即时展开
                content.setMaximumHeight(16777215); content.setVisible(True)
                self._notify_window(True); return
            anim.setEndValue(target); anim.setDuration(420)
            curve = QEasingCurve(QEasingCurve.Type.OutBack); curve.setOvershoot(1.05)
            anim.setEasingCurve(curve)
            self._notify_window(True, animate_window=True)
        anim.start()

    def _on_anim_finished(self):
        if self._collapsed:
            self._content.setVisible(False)
        self._content.setMaximumHeight(16777215)
        self._notify_window(not self._collapsed, animate_window=True)

    def toggle(self):
        self.set_collapsed(not self._collapsed, animate=True)

    def _notify_window(self, expanding, animate_window=False):
        win = self.window()
        if win is not None and win.isVisible() and hasattr(win, "fit_to_content"):
            if animate_window:
                win.fit_to_content(expanding, animate_window=True)
            else:
                QTimer.singleShot(0, lambda: win.fit_to_content(expanding))
```

`ChipBar` 关键点：芯片 `setCheckable(True)`；`section.collapsedChanged` 里同步
`chip.setChecked(not collapsed)`；手风琴分支读一次性标志
`section._no_accordion_close`（读后即复位）跳过自动收起。
`FlowLayout` 用 Qt 官方 flow layout 示例（heightForWidth 驱动换行）。

## 窗口弹性适配骨架（规则4/5）

```python
def fit_to_content(self, expanding=None, animate_window=False):
    if not self.isVisible() or self.isMinimized() or self.isMaximized():
        return
    QApplication.instance().sendPostedEvents(None, QEvent.Type.LayoutRequest)
    page = self.tab_widget.currentWidget()
    hint = page.sizeHint() if expanding else page.minimumSizeHint()
    frame, geo = self.frameGeometry(), self.geometry()
    ft, fl = max(0, geo.top()-frame.top()), max(0, geo.left()-frame.left())
    fr, fb = max(0, frame.right()-geo.right()), max(0, frame.bottom()-geo.bottom())
    avail = self.screen().availableGeometry()
    w = max(520, min(hint.width(),  avail.width()  - fl - fr))
    h = max(300, min(hint.height(), avail.height() - ft - fb))
    x = max(avail.left()+fl, min(self.x(), avail.left()+avail.width()-fr-w))
    y = max(avail.top()+ft,  min(self.y(), avail.top()+avail.height()-fb-h))
    # animate_window=True 时用持久 QPropertyAnimation(b"geometry")，
    # OutQuart 300ms；目标与当前相同不重启动画
```

## 离屏测试清单（QT_QPA_PLATFORM=offscreen，必须全过）

1. 首次显示后 `showEvent` 触发 `fit_to_content(True)`，窗口框架完整在
   availableGeometry 内（Win32 真机探针验证标题栏 T>=0）。
2. chip 选中态与面板 collapsed 永远互斥一致（含动画中途、精简模式后）。
3. 手风琴：展开 B 自动收 A；程序 `accordion_close=False` 展开不收 A。
4. 脏检测：打字→保存面板弹出（程序 `_loading_form` 填充不弹）；保存/加载→收起。
5. 导轨：栏全收起→宽度~126px；重开→按记忆比例恢复；三栏总宽守恒。
6. 压测：120ms 间隔连点 40 轮 + 后台线程 125 信号/秒；50ms 心跳看门狗
   无 >300ms 卡顿；worker 全部退出；芯片状态一致。
7. **测试线程泄漏不要用 `QApplication.threads()`**——可能返回已析构 QThread 的
   悬垂包装，触碰即 0xC0000409 崩溃；改为显式持有并检查应用自己的 worker 引用。

## 线程规则（配合本 UI）

- 所有 worker（QThread）由页面实例属性持有（`self.worker=...`），
  finished 后置 None；禁止局部变量持有运行中线程。
- worker→UI 只走 queued 信号；网络请求 8s 超时 + 工作线程，禁止主线程同步等待。
