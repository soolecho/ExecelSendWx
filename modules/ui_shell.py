# -*- coding: utf-8 -*-
"""现代玻璃拟态 UI 组件：顶部标题栏 / 左侧抽屉导航 / 常驻底部日志坞。

设计说明：
- Qt 没有原生的“背景模糊”控件，这里用「半透明底色 + 顶部高光渐变 + 细描边」
  模拟磨砂玻璃（paint_glass），明暗主题色值统一取自 ui_theme；
- 所有动画对象都在 __init__ 建一次后复用（stop → 重设参数 → start），
  避免反复 new/deleteLater 造成的原生崩溃；
- 左侧导航为抽屉式：宽度在宽栏与图标栏之间平滑动画，并且只放主标签
  （子菜单不再挂在导航里，避免与右侧内容重复）；
- 右侧编辑区顶部是 SectionTabs 小节切换条：左侧分组名 + 右侧等宽小节按钮，
  对齐网格排布，点击后渐变换色并驱动页面内对应的折叠面板展开。
"""

import datetime
import html
import importlib.util
import math
import os
import re
import tempfile
import threading
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from PyQt6.QtCore import (
    QAbstractAnimation,
    QEasingCurve,
    QPoint,
    QPointF,
    QPropertyAnimation,
    QRect,
    QRectF,
    QSize,
    Qt,
    QTimer,
    QUrl,
    QVariantAnimation,
    pyqtProperty,
    pyqtSignal,
)
from PyQt6.QtGui import (
    QAction,
    QActionGroup,
    QBrush,
    QColor,
    QDesktopServices,
    QFont,
    QFontMetrics,
    QGuiApplication,
    QLinearGradient,
    QPainter,
    QPainterPath,
    QPen,
)
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QGraphicsOpacityEffect,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMenu,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from modules import news_fetcher
from modules import ui_theme
from modules import weather_fetcher

# 日志着色正则（与 gui._colorize_log 语义保持一致）
_LOG_ERR_RE = re.compile(r"失败(?!\s*0\b)|出错|错误|异常|❌|Error|Traceback")
_LOG_WARN_RE = re.compile(r"警告|⚠|跳过|超时|WARN")
_LOG_OK_RE = re.compile(r"成功|完成|已启动|运行中|✅|恢复")

# 相同日志行在该窗口内判为重复，直接丢弃（避免多路由重复写入同一面板）
_LOG_DEDUPE_SECONDS = 1.5
_LOG_MAX_BLOCKS = 1500


def _mode() -> str:
    """当前主题模式（优先读取 app 上由 ui_theme.apply_theme 写入的属性）。"""
    app = QApplication.instance()
    if app is not None:
        mode = app.property("ui_theme")
        if mode in (ui_theme.LIGHT, ui_theme.DARK):
            return mode
    return ui_theme.LIGHT


def _colors() -> Dict[str, str]:
    return ui_theme.colors_for(_mode())


def colorize_log(text: str, colors: Optional[Dict[str, str]] = None) -> str:
    """按关键词给单行日志着色（红=失败、橙=警告、绿=成功）。"""
    c = colors or _colors()
    esc = html.escape(str(text))
    if _LOG_ERR_RE.search(str(text)):
        return f'<span style="color:{c["danger"]};">{esc}</span>'
    if _LOG_WARN_RE.search(str(text)):
        return f'<span style="color:{c["warn"]};">{esc}</span>'
    if _LOG_OK_RE.search(str(text)):
        return f'<span style="color:{c["ok"]};">{esc}</span>'
    return esc


def paint_glass(painter: QPainter,
                rect: QRectF,
                radius: float = 14.0,
                strong: bool = False) -> None:
    """在给定矩形内绘制玻璃面板（半透明底 + 顶部高光 + 细描边）。"""
    c = _colors()
    path = QPainterPath()
    path.addRoundedRect(rect, radius, radius)
    painter.fillPath(path, ui_theme.qcolor(c["glass_strong" if strong else "glass"]))

    # 顶部到中部的白色高光，模拟磨砂玻璃受光面
    grad = QLinearGradient(rect.topLeft(), rect.bottomLeft())
    top = ui_theme.qcolor(c["glass_sheen"])
    mid = QColor(top)
    mid.setAlpha(0)
    bottom = QColor(255, 255, 255, 0)
    grad.setColorAt(0.0, top)
    grad.setColorAt(0.45, mid)
    grad.setColorAt(1.0, bottom)
    painter.fillPath(path, QBrush(grad))

    pen = QPen(ui_theme.qcolor(c["glass_border"]))
    pen.setWidthF(1.0)
    painter.setPen(pen)
    painter.drawPath(path)


# ---------------------------------------------------------------------------
# 日志视图
# ---------------------------------------------------------------------------
class GlobalLogView(QTextEdit):
    """全局日志框：着色 + 自动滚底 + 短窗口去重（同一行不重复刷两遍）。"""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("GlobalLog")
        self.setReadOnly(True)
        self.setAcceptRichText(True)
        self.setUndoRedoEnabled(False)
        self.setFont(QFont("Cascadia Mono", 9))
        self.document().setMaximumBlockCount(_LOG_MAX_BLOCKS)
        self._last_line = ""
        self._last_at = 0.0
        self._auto_follow = True

    def set_auto_follow(self, follow: bool) -> None:
        self._auto_follow = bool(follow)

    def append_log(self, message: str, ts: str = "") -> None:
        """追加一行日志（ts 为空则不带时间前缀）。"""
        line = f"[{ts}] {message}" if ts else str(message)
        now = time.monotonic()
        # 去重：同一行在 1.5 秒内只保留一条（多条日志路由写同一面板时兜底）
        if line == self._last_line and now - self._last_at < _LOG_DEDUPE_SECONDS:
            self._last_at = now
            return
        self._last_line, self._last_at = line, now
        self.append(colorize_log(line))
        if self._auto_follow:
            sb = self.verticalScrollBar()
            sb.setValue(sb.maximum())

    def clear_log(self) -> None:
        self.clear()
        self._last_line = ""


# ---------------------------------------------------------------------------
# 进度条
# ---------------------------------------------------------------------------
class GlassProgressBar(QProgressBar):
    """玻璃质感进度条：圆角轨道 + 渐变进度 + 流动高光。"""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setRange(0, 100)
        self.setValue(0)
        self.setTextVisible(True)
        self.setFixedHeight(18)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._phase = 0.0
        self._shine = QVariantAnimation(self)
        self._shine.setStartValue(0.0)
        self._shine.setEndValue(1.0)
        self._shine.setDuration(2400)
        self._shine.setLoopCount(-1)
        self._shine.setEasingCurve(QEasingCurve.Type.Linear)
        self._shine.valueChanged.connect(self._on_phase)

    def _on_phase(self, value) -> None:
        self._phase = float(value)
        self.update()

    def showEvent(self, event):
        super().showEvent(event)
        self._shine.start()

    def hideEvent(self, event):
        self._shine.stop()
        super().hideEvent(event)

    def paintEvent(self, event):
        c = _colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        radius = max(1.0, r.height() / 2.0)

        track = QPainterPath()
        track.addRoundedRect(r, radius, radius)
        p.fillPath(track, ui_theme.qcolor(c["glass_track"]))
        p.setPen(QPen(ui_theme.qcolor(c["glass_border"])))
        p.drawPath(track)

        span = self.maximum() - self.minimum()
        frac = 0.0 if span <= 0 else (self.value() - self.minimum()) / float(span)
        frac = max(0.0, min(1.0, frac))
        if frac > 0.002:
            width = max(r.height(), r.width() * frac)
            fill_rect = QRectF(r.left(), r.top(), width, r.height())
            fill = QPainterPath()
            fill.addRoundedRect(fill_rect, radius, radius)
            grad = QLinearGradient(fill_rect.topLeft(), fill_rect.topRight())
            grad.setColorAt(0.0, ui_theme.qcolor(c["accent"]))
            grad.setColorAt(0.65, ui_theme.qcolor(c["accent"]))
            grad.setColorAt(1.0, ui_theme.qcolor(c["ok"]))
            p.fillPath(fill, QBrush(grad))

            # 流动高光：一条柔和的白色宽带从左到右循环扫过
            shine = QLinearGradient(fill_rect.topLeft(), fill_rect.topRight())
            clear = QColor(255, 255, 255, 0)
            glow = QColor(255, 255, 255, 150 if _mode() == ui_theme.LIGHT else 80)
            pos = self._phase
            shine.setColorAt(max(0.0, min(0.999, pos - 0.18)), clear)
            shine.setColorAt(max(0.001, min(0.999, pos)), glow)
            shine.setColorAt(min(1.0, pos + 0.18), clear)
            p.fillPath(fill, QBrush(shine))

        if self.isTextVisible():
            text = self.text()
            if text:
                p.setPen(ui_theme.qcolor(c["text"]))
                font = QFont(self.font())
                font.setPointSizeF(max(7.5, font.pointSizeF() * 0.92))
                p.setFont(font)
                p.drawText(QRectF(self.rect()), Qt.AlignmentFlag.AlignCenter, text)


# ---------------------------------------------------------------------------
# 右侧顶部：小节切换条（分组小按钮 + 动画）
# ---------------------------------------------------------------------------
def _mix(a: QColor, b: QColor, t: float) -> QColor:
    """按比例混合两个颜色（含 alpha），用于选中态/悬停动画插值。"""
    t = max(0.0, min(1.0, float(t)))
    return QColor(
        int(round(a.red() + (b.red() - a.red()) * t)),
        int(round(a.green() + (b.green() - a.green()) * t)),
        int(round(a.blue() + (b.blue() - a.blue()) * t)),
        int(round(a.alpha() + (b.alpha() - a.alpha()) * t)),
    )


class SectionChip(QPushButton):
    """小节按钮（小号胶囊）：未选中为幽灵态，悬停浅色，选中渐变换色 + 按下回弹。"""

    def __init__(self, key: str, label: str, width: int,
                 parent: Optional[QWidget] = None):
        super().__init__(label, parent)
        self.key = key
        self.setCheckable(True)
        self.setFixedSize(int(width), 26)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setAttribute(Qt.WidgetAttribute.WA_Hover, True)
        self.setToolTip(f"{label}（再次点击可收起）")
        self._t = 0.0       # 选中过渡 0→1
        self._hover = 0.0
        self._press = 0.0

        self._sel_anim = self._make_anim(220, QEasingCurve.Type.OutCubic,
                                         self._set_t)
        self._hover_anim = self._make_anim(150, QEasingCurve.Type.OutCubic,
                                           self._set_hover)
        self._press_anim = self._make_anim(150, QEasingCurve.Type.OutBack,
                                           self._set_press)

    def _make_anim(self, duration: int, curve, slot) -> QVariantAnimation:
        anim = QVariantAnimation(self)
        anim.setDuration(duration)
        anim.setEasingCurve(curve)
        anim.valueChanged.connect(slot)
        return anim

    def _set_t(self, value) -> None:
        self._t = float(value)
        self.update()

    def _set_hover(self, value) -> None:
        self._hover = float(value)
        self.update()

    def _set_press(self, value) -> None:
        self._press = float(value)
        self.update()

    def _run(self, anim: QVariantAnimation, start: float, end: float) -> None:
        anim.stop()
        anim.setStartValue(float(start))
        anim.setEndValue(float(end))
        anim.start()

    def _chip_font(self) -> QFont:
        font = QFont(self.font())
        font.setPixelSize(11)
        return font

    # ---- 状态 ----
    def setChecked(self, checked: bool) -> None:
        """选中态变化走动画（而不是瞬间跳色）。"""
        super().setChecked(checked)
        self._run(self._sel_anim, self._t, 1.0 if checked else 0.0)

    def enterEvent(self, event):
        self._run(self._hover_anim, self._hover, 1.0)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._run(self._hover_anim, self._hover, 0.0)
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._run(self._press_anim, self._press, 1.0)
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._run(self._press_anim, self._press, 0.0)
        super().mouseReleaseEvent(event)

    # ---- 绘制 ----
    def paintEvent(self, event):
        c = _colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        radius = rect.height() / 2.0

        scale = 1.0 - 0.04 * self._press
        if scale != 1.0:
            center = rect.center()
            p.translate(center)
            p.scale(scale, scale)
            p.translate(-center)

        t = self._t
        hover = self._hover
        accent = ui_theme.qcolor(c["accent"])
        soft = ui_theme.qcolor(c["accent_soft"])
        text = ui_theme.qcolor(c["text"])

        path = QPainterPath()
        path.addRoundedRect(rect, radius, radius)

        # 幽灵按钮：平时无底色，悬停浅色描边，选中整块强调色填充（渐变过渡）
        if t > 0.01:
            p.fillPath(path, _mix(soft, accent, t))
            pen = QPen(accent)
            pen.setWidthF(1.0)
            p.setPen(pen)
            p.drawPath(path)
            sheen = QLinearGradient(rect.topLeft(), rect.bottomLeft())
            sheen.setColorAt(0.0, QColor(255, 255, 255, int(52 * t)))
            sheen.setColorAt(1.0, QColor(255, 255, 255, 0))
            p.fillPath(path, QBrush(sheen))
        elif hover > 0.01:
            bg = QColor(soft)
            bg.setAlpha(int(round(255 * 0.75 * hover)))
            p.fillPath(path, bg)
            edge = QColor(accent)
            edge.setAlpha(int(round(105 * hover)))
            pen = QPen(edge)
            pen.setWidthF(1.0)
            p.setPen(pen)
            p.drawPath(path)

        font = self._chip_font()
        if t > 0.35:
            font.setBold(True)
        p.setFont(font)
        base_text = _mix(text, accent, 0.5 * hover) if t < 0.3 else text
        p.setPen(QPen(_mix(base_text, QColor("#ffffff"), t)))
        p.drawText(rect, Qt.AlignmentFlag.AlignCenter, self.text())


class GroupChip(SectionChip):
    """分组胶囊：一个胶囊代表一个分组。

    - 该组含当前展开的小节时染强调色（一眼看出内容属于哪一组）；
    - 右侧小三角随展开状态旋转（收起 ▸ / 展开 ▾）；
    - 点击展开该组的小节行，再次点击收起（只收起小节行，不影响内容）。
    """

    def __init__(self, key: str, title: str, width: int,
                 parent: Optional[QWidget] = None):
        super().__init__(key, title, width, parent)
        self._open_t = 0.0
        self._open_anim = self._make_anim(200, QEasingCurve.Type.OutCubic,
                                          self._set_open_t)
        self.setToolTip(f"{title}（点击展开 / 收起该组小节）")

    def _set_open_t(self, value) -> None:
        self._open_t = float(value)
        self.update()

    def set_open(self, opened: bool) -> None:
        self._run(self._open_anim, self._open_t, 1.0 if opened else 0.0)

    def paintEvent(self, event):
        c = _colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        radius = rect.height() / 2.0

        t = self._t          # 当前分组是否含激活小节
        hover = self._hover
        accent = ui_theme.qcolor(c["accent"])
        soft = ui_theme.qcolor(c["accent_soft"])
        border = ui_theme.qcolor(c["glass_border"])
        muted = ui_theme.qcolor(c["muted"])
        strong = ui_theme.qcolor(c["glass_strong"])

        fill = _mix(strong, soft, t)
        if hover > 0.01:
            fill = _mix(fill, soft, 0.45 * hover)
        edge = _mix(border, accent, max(t, 0.55 * hover))

        path = QPainterPath()
        path.addRoundedRect(rect, radius, radius)
        p.fillPath(path, fill)
        pen = QPen(edge)
        pen.setWidthF(1.0)
        p.setPen(pen)
        p.drawPath(path)

        font = self._chip_font()
        font.setBold(True)
        p.setFont(font)
        p.setPen(QPen(_mix(muted, accent, t)))
        p.drawText(rect.adjusted(11, 0, -18, 0),
                   Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                   self.text())

        # 右侧小三角：收起时朝右，展开时（动画）旋转朝下
        center_x = rect.right() - 10.0
        center_y = rect.center().y()
        tri = QPainterPath()
        tri.moveTo(-2.6, -4.6)
        tri.lineTo(3.4, 0.0)
        tri.lineTo(-2.6, 4.6)
        tri.closeSubpath()
        p.save()
        p.translate(center_x, center_y)
        p.rotate(90.0 * self._open_t)
        p.fillPath(tri, _mix(muted, accent, max(t, 0.4 * hover)))
        p.restore()


class SectionTabs(QWidget):
    """右侧编辑区顶部的小节切换条（两级折叠）。

    第一级：分组胶囊（一排，窄窗口也放得下）；
    第二级：当前展开分组的若干小节按钮（只在需要时展开成第二行）。

    这样常规状态只占一行高度，又保留「看得到全部小节名」的能力：
    点分组 → 展开该组小节（并自动展开组内当前/第一个小节），再点收起。
    """

    section_selected = pyqtSignal(str)

    _CHIP_H = 26
    _GAP = 6
    _PAD_X = 12
    _PAD_Y = 9

    def __init__(self, spec: Sequence[Tuple[str, Sequence[Tuple[str, str]]]],
                 parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("SectionTabs")
        self._chips: Dict[str, SectionChip] = {}
        self._group_chips: Dict[str, GroupChip] = {}
        self._groups: Dict[str, List[Any]] = {}
        self._group_of_section: Dict[str, str] = {}
        self._sections_of_group: Dict[str, List[str]] = {}
        self._open_group = ""
        self._active = ""

        # 等宽小节按钮：按最长标签算一次，保证第二行按钮整齐一致
        chip_font = QFont(self.font())
        chip_font.setPixelSize(11)
        chip_metrics = QFontMetrics(chip_font)
        texts = [str(text) for _group, items in spec for _key, text in items]
        widest = max(chip_metrics.horizontalAdvance(t) for t in texts) if texts else 0
        self._chip_w = max(88, widest + 34)

        group_font = QFont(self.font())
        group_font.setPixelSize(11)
        group_font.setBold(True)
        group_metrics = QFontMetrics(group_font)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(self._PAD_X, self._PAD_Y, self._PAD_X, self._PAD_Y)
        layout.setSpacing(self._GAP)

        # ---- 第一级：分组胶囊 ----
        group_row = QHBoxLayout()
        group_row.setContentsMargins(0, 0, 0, 0)
        group_row.setSpacing(8)
        layout.addLayout(group_row)

        # ---- 第二级：当前分组的小节按钮 ----
        self._chip_row = QWidget(self)
        self._chip_row.setMaximumHeight(0)
        self._chip_row.setVisible(False)
        chip_row_layout = QHBoxLayout(self._chip_row)
        chip_row_layout.setContentsMargins(0, 0, 0, 0)
        chip_row_layout.setSpacing(self._GAP)
        layout.addWidget(self._chip_row)

        self._chip_row_expanded = False
        self._chip_anim = QVariantAnimation(self)
        self._chip_anim.valueChanged.connect(self._on_chip_row_anim)
        self._chip_anim.finished.connect(self._on_chip_row_anim_finished)

        widest_chip_row = 0
        group_total = 0
        for index, (group_label, items) in enumerate(spec):
            group_key = f"group{index}"
            title = str(group_label or "分组")
            group_w = max(88, group_metrics.horizontalAdvance(title) + 40)
            group_chip = GroupChip(group_key, title, group_w, self)
            group_chip.clicked.connect(
                lambda _checked=False, k=group_key: self._on_group(k))
            group_row.addWidget(group_chip)
            self._group_chips[group_key] = group_chip
            self._sections_of_group[group_key] = []
            group_total += group_w
            for key, text in items:
                chip = SectionChip(key, str(text), self._chip_w, self._chip_row)
                chip.setVisible(False)
                chip.clicked.connect(
                    lambda _checked=False, k=key: self._on_click(k))
                chip_row_layout.addWidget(chip)
                self._chips[key] = chip
                self._group_of_section[key] = group_key
                self._sections_of_group[group_key].append(key)
            count = len(items)
            row_w = count * self._chip_w + max(0, count - 1) * self._GAP
            widest_chip_row = max(widest_chip_row, row_w)

        group_row.addStretch(1)

        # 面板宽度固定为「最宽的一行」，切换分组时不会忽宽忽窄
        row1_w = group_total + 8 * max(0, len(self._group_chips) - 1)
        panel_w = max(widest_chip_row, row1_w) + self._PAD_X * 2
        self.setFixedWidth(int(panel_w))
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

    def sizeHint(self):
        return self.layout().sizeHint()

    def minimumSizeHint(self):
        return self.layout().sizeHint()

    # ---- 分组展开 / 收起 ----
    def _animate_chip_row(self, target_h: int, animated: bool) -> None:
        anim = self._chip_anim
        anim.stop()
        if not animated:
            self._chip_row.setMaximumHeight(int(target_h))
            self._chip_row.setVisible(target_h > 0)
            self.updateGeometry()
            return
        if target_h > 0:
            self._chip_row.setVisible(True)
        anim.setDuration(160 if target_h <= 0 else 200)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.setStartValue(float(self._chip_row.maximumHeight()))
        anim.setEndValue(float(target_h))
        anim.start()

    def _on_chip_row_anim(self, value) -> None:
        self._chip_row.setMaximumHeight(int(round(float(value))))
        self.updateGeometry()

    def _on_chip_row_anim_finished(self) -> None:
        if not self._chip_row_expanded:
            self._chip_row.setVisible(False)

    def open_group(self, group_key: str, animated: bool = True) -> None:
        if group_key not in self._group_chips:
            return
        self._open_group = group_key
        for key, chip in self._group_chips.items():
            chip.set_open(key == group_key)
            chip.setChecked(key == self._group_of_section.get(self._active))
        for key, chip in self._chips.items():
            chip.setVisible(self._group_of_section.get(key) == group_key)
        self._chip_row_expanded = True
        self._animate_chip_row(self._CHIP_H, animated)

    def close_group(self, animated: bool = True) -> None:
        self._open_group = ""
        for chip in self._group_chips.values():
            chip.set_open(False)
        self._chip_row_expanded = False
        self._animate_chip_row(0, animated)

    def _on_group(self, group_key: str) -> None:
        if self._open_group == group_key:
            self.close_group()
            return
        self.open_group(group_key)
        target = self._active if self._group_of_section.get(self._active) == group_key else ""
        if not target:
            sections = self._sections_of_group.get(group_key) or []
            target = sections[0] if sections else ""
        if not target:
            return
        self.set_active(target)
        items = self._groups.get(target) or []
        # 组内目标小节的面板还折叠着 → 通知页面展开（否则只展开小节行，不切内容）
        if items and not any(not g.is_collapsed() for g in items):
            self.section_selected.emit(target)

    # ---- 状态同步 ----
    def bind_group(self, section_key: str, groups: Sequence[Any]) -> None:
        """把页面内折叠面板的展开/收起状态同步到芯片选中态。"""
        items = [g for g in groups if g is not None]
        self._groups[section_key] = items
        for g in items:
            g.collapsedChanged.connect(
                lambda _collapsed, k=section_key: self._sync(k))

    def _sync(self, section_key: str) -> None:
        chip = self._chips.get(section_key)
        items = self._groups.get(section_key) or []
        if chip is None or not items:
            return
        expanded = any(not g.is_collapsed() for g in items)
        if chip.isChecked() != expanded:
            chip.setChecked(expanded)
        if not expanded:
            if self._active == section_key:
                self._active = ""
            return
        self._active = section_key
        group_key = self._group_of_section.get(section_key, "")
        for key, group_chip in self._group_chips.items():
            group_chip.setChecked(key == group_key)
        if group_key and self._open_group != group_key:
            self.open_group(group_key, animated=True)

    def set_active(self, section_key: str) -> None:
        self._active = section_key
        group_key = self._group_of_section.get(section_key, "")
        for key, chip in self._chips.items():
            chip.setChecked(key == section_key)
        for key, group_chip in self._group_chips.items():
            group_chip.setChecked(key == group_key)
        if group_key and self._open_group != group_key:
            self.open_group(group_key, animated=False)

    def active(self) -> str:
        return self._active

    def _on_click(self, key: str) -> None:
        items = self._groups.get(key) or []
        # 已选中的小节再次点击：收起它的面板（背景色跟随动画淡出）
        if self._active == key and any(not g.is_collapsed() for g in items):
            self._active = ""
            for chip in self._chips.values():
                chip.setChecked(False)
            for group_chip in self._group_chips.values():
                group_chip.setChecked(False)
            for g in items:
                if not g.is_collapsed():
                    g.set_collapsed(True, animate=True)
            return
        self.set_active(key)
        self.section_selected.emit(key)

    # ---- 主题 / 绘制 ----
    def refresh_theme(self) -> None:
        self.update()
        for chip in self._group_chips.values():
            chip.update()
        for chip in self._chips.values():
            chip.update()

    def paintEvent(self, event):
        # 整条只画一块玻璃底：分组胶囊与小节按钮同属一个切换条
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        paint_glass(painter, rect, 12.0, strong=True)

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == event.Type.PaletteChange:
            self.refresh_theme()


class NavButton(QWidget):
    """自绘导航主按钮：hover 渐变、按下缩放回弹、选中填色带动画。"""

    clicked = pyqtSignal()

    def __init__(self, key: str, icon: str, title: str,
                 parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.key = key
        self._icon = icon
        self._title = title
        self._active = False
        self._active_t = 0.0
        self._hover = 0.0
        self._press = 0.0
        self._compact = False
        self.setFixedHeight(42)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setAttribute(Qt.WidgetAttribute.WA_Hover, True)

        self._active_anim = QVariantAnimation(self)
        self._active_anim.setDuration(240)
        self._active_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._active_anim.valueChanged.connect(self._on_active)

        self._hover_anim = QVariantAnimation(self)
        self._hover_anim.setDuration(170)
        self._hover_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._hover_anim.valueChanged.connect(self._on_hover)

        self._press_anim = QVariantAnimation(self)
        self._press_anim.setDuration(150)
        self._press_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._press_anim.valueChanged.connect(self._on_press)

    # ---- 动画 ----
    def _animate(self, anim: QVariantAnimation, current: float, target: float,
                 duration: int, curve: QEasingCurve.Type) -> None:
        anim.stop()
        anim.setStartValue(float(current))
        anim.setEndValue(float(target))
        anim.setDuration(duration)
        anim.setEasingCurve(curve)
        anim.start()

    def _on_active(self, value) -> None:
        self._active_t = float(value)
        self.update()

    def _on_hover(self, value) -> None:
        self._hover = float(value)
        self.update()

    def _on_press(self, value) -> None:
        self._press = float(value)
        self.update()

    # ---- 事件 ----
    def enterEvent(self, event):
        self._animate(self._hover_anim, self._hover, 1.0, 170,
                      QEasingCurve.Type.OutCubic)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._animate(self._hover_anim, self._hover, 0.0, 220,
                      QEasingCurve.Type.OutCubic)
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._animate(self._press_anim, self._press, 1.0, 90,
                          QEasingCurve.Type.OutCubic)
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._animate(self._press_anim, self._press, 0.0, 220,
                          QEasingCurve.Type.OutBack)
            if self.rect().contains(event.position().toPoint()):
                self.clicked.emit()
        super().mouseReleaseEvent(event)

    # ---- 状态 ----
    def set_active(self, active: bool) -> None:
        active = bool(active)
        if self._active == active:
            return
        self._active = active
        self._animate(self._active_anim, self._active_t,
                      1.0 if active else 0.0, 240,
                      QEasingCurve.Type.OutCubic)

    def is_active(self) -> bool:
        return self._active

    def set_compact(self, compact: bool) -> None:
        self._compact = bool(compact)
        self.update()

    def refresh_theme(self) -> None:
        self.update()

    def sizeHint(self):
        return QSize(72, 42)

    # ---- 绘制 ----
    def paintEvent(self, event):
        c = _colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)

        rect = QRectF(self.rect()).adjusted(2.0, 2.0, -2.0, -2.0)
        scale = 1.0 - 0.035 * self._press
        if scale != 1.0:
            center = rect.center()
            p.translate(center)
            p.scale(scale, scale)
            p.translate(-center)

        path = QPainterPath()
        path.addRoundedRect(rect, 11.0, 11.0)

        t = self._active_t
        bg = _mix(ui_theme.qcolor(c["glass"]),
                  ui_theme.qcolor(c["accent_soft"]), t)
        if self._hover > 0.0:
            bg = _mix(bg, ui_theme.qcolor(c["glass_strong"]), 0.5 * self._hover)
        p.fillPath(path, bg)

        pen = QPen(_mix(ui_theme.qcolor(c["glass_border"]),
                        ui_theme.qcolor(c["accent"]), 0.55 * t))
        pen.setWidthF(1.0)
        p.setPen(pen)
        p.drawPath(path)

        # 选中态左侧光条（高度随动画一起长出来）
        if t > 0.01:
            shade = QColor(ui_theme.qcolor(c["accent"]))
            shade.setAlpha(int(255 * t))
            bar = QRectF(rect.left() + 3.0, rect.top() + 9.0, 3.0,
                         (rect.height() - 18.0) * t)
            bar_path = QPainterPath()
            bar_path.addRoundedRect(bar, 1.5, 1.5)
            p.fillPath(bar_path, shade)

        # 未选中用正文色（灰字在浅色玻璃上几乎看不清），选中渐变为强调色
        text_color = _mix(ui_theme.qcolor(c["text"]),
                          ui_theme.qcolor(c["accent"]), t)
        p.setPen(text_color)

        font = QFont(self.font())
        font.setPointSizeF(max(9.0, font.pointSizeF()))
        if t > 0.35:
            font.setBold(True)
        p.setFont(font)

        if self._compact:
            p.drawText(rect, Qt.AlignmentFlag.AlignCenter, self._icon)
            return

        icon_rect = QRectF(rect.left() + 14.0, rect.top(), 24.0, rect.height())
        p.drawText(icon_rect, Qt.AlignmentFlag.AlignVCenter
                   | Qt.AlignmentFlag.AlignLeft, self._icon)
        text_rect = QRectF(rect.left() + 44.0, rect.top(),
                           max(10.0, rect.width() - 54.0), rect.height())
        metrics = QFontMetrics(font)
        elided = metrics.elidedText(self._title, Qt.TextElideMode.ElideRight,
                                    int(text_rect.width()))
        p.drawText(text_rect, Qt.AlignmentFlag.AlignVCenter
                   | Qt.AlignmentFlag.AlignLeft, elided)


class SideNavPanel(QFrame):
    """左侧竖排导航：只放主标签（无子菜单），选中项填色 + 过渡动画。"""

    page_selected = pyqtSignal(str)

    WIDE = 200
    RAIL = 72

    def __init__(self, spec: Sequence[Dict], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("NavPanel")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, False)
        self.setMinimumWidth(self.WIDE)
        self.setMaximumWidth(self.WIDE)

        self._spec = list(spec)
        self._nav_buttons: Dict[str, NavButton] = {}
        self._current_page = ""
        self._rail = False

        self._width_anim = QVariantAnimation(self)
        self._width_anim.setDuration(300)
        self._width_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._width_anim.valueChanged.connect(self._on_width_value)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # ---- 品牌区 ----
        self._brand = QWidget(self)
        brand_row = QHBoxLayout(self._brand)
        brand_row.setContentsMargins(4, 2, 4, 2)
        brand_row.setSpacing(8)
        self.brand_icon = QLabel("📊", self._brand)
        self.brand_icon.setStyleSheet("font-size:20px; background:transparent;")
        brand_row.addWidget(self.brand_icon)
        self._brand_text = QWidget(self._brand)
        brand_text_col = QVBoxLayout(self._brand_text)
        brand_text_col.setContentsMargins(0, 0, 0, 0)
        brand_text_col.setSpacing(0)
        self.brand_title = QLabel("表格自动发送", self._brand_text)
        self.brand_title.setObjectName("NavBrand")
        self.brand_sub = QLabel("By 春风予Lu", self._brand_text)
        self.brand_sub.setObjectName("NavBrandSub")
        brand_text_col.addWidget(self.brand_title)
        brand_text_col.addWidget(self.brand_sub)
        brand_row.addWidget(self._brand_text, 1)
        layout.addWidget(self._brand)

        # ---- 导航项（滚动区域，窗口很矮时也不会溢出）----
        host = QWidget(self)
        host_layout = QVBoxLayout(host)
        host_layout.setContentsMargins(0, 0, 0, 0)
        host_layout.setSpacing(6)
        for page in self._spec:
            key = str(page.get("key", ""))
            btn = NavButton(key, page.get("icon", "•"), page.get("title", key), host)
            btn.clicked.connect(lambda k=key: self._on_nav_click(k))
            host_layout.addWidget(btn)
            self._nav_buttons[key] = btn
        host_layout.addStretch(1)

        self._scroll = QScrollArea(self)
        self._scroll.setObjectName("NavScroll")
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.Shape.NoFrame)
        self._scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._scroll.viewport().setAutoFillBackground(False)
        self._scroll.setWidget(host)
        layout.addWidget(self._scroll, 1)

        # ---- 底部：抽屉收起/展开 ----
        self.rail_btn = QPushButton("⏴ 收起导航", self)
        self.rail_btn.setObjectName("NavRailBtn")
        self.rail_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.rail_btn.setFixedHeight(28)
        self.rail_btn.clicked.connect(lambda: self.set_rail(not self._rail))
        layout.addWidget(self.rail_btn)

    # ---- 自适应：矮窗口下压缩导航项，保证四个主模块一眼看全 ----
    def showEvent(self, event):
        super().showEvent(event)
        self._fit_nav()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._fit_nav()

    def _fit_nav(self) -> None:
        """按可用高度压缩导航项与品牌区。

        正常编辑模式下导航必须完整显示四个主模块（不能出现滚动条等用户拖），
        因此窗口变矮时依次：收副标题 → 收品牌块 → 收紧按钮高度与间距。
        """
        count = len(self._nav_buttons)
        host = self._scroll.widget() if getattr(self, "_scroll", None) else None
        host_layout = host.layout() if host is not None else None
        if count <= 0 or host_layout is None:
            return
        lay = self.layout()
        margins = lay.contentsMargins()
        spacing = lay.spacing()
        rail_h = self.rail_btn.height() or 30
        avail = max(0, self.height() - margins.top() - margins.bottom()
                    - rail_h - spacing * 2)
        # 预算：四个按钮最少 20px + 2px 间隔；品牌区按剩余空间分级显示
        need_buttons = count * 20 + (count - 1) * 2
        show_sub = (not self._rail) and avail >= need_buttons + 33 + spacing + 10
        show_brand = avail >= need_buttons + (33 if self._rail else 23) + spacing + 8
        if self._rail:
            self._brand.setVisible(show_brand)
            self._brand_text.setVisible(False)
        else:
            self._brand.setVisible(show_brand)
            self._brand_text.setVisible(show_brand)
            self.brand_sub.setVisible(show_sub)
        brand_h = (self._brand.sizeHint().height() if show_brand else 0)
        area = max(0, avail - (brand_h + spacing if show_brand else 0))

        gap = 6 if area >= count * 40 else (4 if area >= count * 30 else 2)
        btn_h = int((area - gap * (count - 1)) / count)
        btn_h = max(20, min(42, btn_h))
        if host_layout.spacing() != gap:
            host_layout.setSpacing(gap)
        for btn in self._nav_buttons.values():
            if btn.height() != btn_h:
                btn.setFixedHeight(btn_h)
        need_h = count * btn_h + gap * (count - 1)
        if host.minimumHeight() != need_h:
            host.setMinimumHeight(need_h)

    # ---- 绘制 ----
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        paint_glass(painter, rect, 16.0, strong=True)

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == event.Type.PaletteChange:
            self.refresh_theme()

    # ---- 主题 ----
    def refresh_theme(self) -> None:
        for btn in self._nav_buttons.values():
            btn.refresh_theme()
        self.update()

    # ---- 抽屉 ----
    def set_rail(self, rail: bool) -> None:
        rail = bool(rail)
        if rail == self._rail:
            return
        self._rail = rail
        for btn in self._nav_buttons.values():
            btn.set_compact(rail)
        self._brand_text.setVisible(not rail)
        self.rail_btn.setText("⏵" if rail else "⏴ 收起导航")
        self.rail_btn.setToolTip("展开导航" if rail else "收起为图标导航")
        self._width_anim.stop()
        self._width_anim.setStartValue(self.width())
        self._width_anim.setEndValue(self.RAIL if rail else self.WIDE)
        self._width_anim.start()
        self._fit_nav()

    def _on_width_value(self, value) -> None:
        width = int(value)
        self.setMinimumWidth(width)
        self.setMaximumWidth(width)

    def is_rail(self) -> bool:
        return self._rail

    # ---- 导航交互 ----
    def _set_current(self, key: str) -> None:
        self._current_page = key
        for k, btn in self._nav_buttons.items():
            btn.set_active(k == key)

    def _on_nav_click(self, key: str) -> None:
        # 收起成图标栏后点击图标只做切页，保持图标栏状态（不再自动展开）
        self._set_current(key)
        self.page_selected.emit(key)

    # ---- 供 MainWindow 调用 ----
    def select_page(self, key: str) -> None:
        """程序切换页面：只高亮导航项（不重复发 page_selected）。"""
        if key not in self._nav_buttons:
            return
        self._set_current(key)

    def active_page(self) -> str:
        return self._current_page


# ---------------------------------------------------------------------------
# 内置简易浏览器（可选：需要安装 PyQt6-WebEngine）
# ---------------------------------------------------------------------------
_WEBENGINE_READY: Optional[bool] = None
_WEBENGINE_INSTALLED: Optional[bool] = None


def webengine_installed() -> bool:
    """轻量探测：只查 PyQt6-WebEngine 是否装了，不加载 Chromium 内核。

    用于界面提示文案：启动时若真去 import QtWebEngineWidgets，会白白加载
    几十 MB 的 Chromium 动态库，拖慢冷启动；这里只查包目录是否存在。
    """
    global _WEBENGINE_INSTALLED
    if _WEBENGINE_INSTALLED is None:
        try:
            _WEBENGINE_INSTALLED = (
                importlib.util.find_spec("PyQt6.QtWebEngineWidgets") is not None)
        except Exception:
            _WEBENGINE_INSTALLED = False
    return _WEBENGINE_INSTALLED


def webengine_available() -> bool:
    """是否可以使用 Qt WebEngine 内核（PyQt6-WebEngine 已安装）。

    结果缓存一次；offscreen 测试环境直接视为不可用，避免加载浏览器内核。
    注意：调用方须在创建 QApplication 前设置 AA_ShareOpenGLContexts，
    否则 QtWebEngineWidgets 的延迟导入会抛 ImportError。
    """
    global _WEBENGINE_READY
    if _WEBENGINE_READY is not None:
        return _WEBENGINE_READY
    if os.environ.get("QT_QPA_PLATFORM", "").lower() == "offscreen":
        _WEBENGINE_READY = False
        return False
    try:
        from PyQt6.QtWebEngineWidgets import QWebEngineView  # noqa: F401
        from PyQt6.QtWebEngineCore import QWebEngineProfile  # noqa: F401
        _WEBENGINE_READY = True
    except Exception:
        _WEBENGINE_READY = False
    return _WEBENGINE_READY


_BROWSER_PROFILE: Optional[Any] = None


def browser_profile():
    """内置浏览器的持久化配置（Cookie / 缓存写入 LOCALAPPDATA）。

    Qt 的默认配置是「无痕」的：Cookie 不落盘，微博、知乎这类站点会把内嵌
    窗口当成访客或直接防盗链，表现为 "Sina Visitor System" 或加载失败。
    这里改用持久化配置 + 常规 Chrome UA，登录/验证一次后可长期复用。
    """
    global _BROWSER_PROFILE
    if _BROWSER_PROFILE is not None:
        return _BROWSER_PROFILE
    try:
        from PyQt6.QtWebEngineCore import QWebEngineProfile
    except Exception:
        return None
    local = os.environ.get("LOCALAPPDATA")
    base = (Path(local) / "ExcelSendWx" if local
            else Path(tempfile.gettempdir()) / "ExcelSendWx")
    try:
        base.mkdir(parents=True, exist_ok=True)
    except OSError:
        return None
    profile = QWebEngineProfile("wxauto-browser", QApplication.instance())
    try:
        profile.setPersistentStoragePath(str(base / "browser"))
        profile.setCachePath(str(base / "browser" / "cache"))
        profile.setPersistentCookiesPolicy(
            QWebEngineProfile.PersistentCookiesPolicy.ForcePersistentCookies)
        profile.setHttpUserAgent(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36")
    except Exception:
        pass  # 配置失败时退回 Qt 默认行为，不影响浏览器可用性
    _BROWSER_PROFILE = profile
    return profile


class StandbyBrowser(QWidget):
    """待机画面里的内置浏览器（嵌在待机右侧，不弹新窗口）。

    由 StandbyView 覆盖在滚动区上：点新闻标题时从右侧滑入，
    点「← 待机」再滑出，回到时钟 / 天气视图；主窗口随之放大 / 收回。
    """

    back_requested = pyqtSignal()

    def __init__(self, parent: QWidget):
        super().__init__(parent)
        self.setObjectName("StandbyBrowser")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.view = None
        self._home_url = ""
        self._last_url = ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        bar = QHBoxLayout()
        bar.setSpacing(6)
        self.back_btn = self._nav_btn("← 待机", "返回待机画面（Esc）")
        self.back_btn.setObjectName("BrowserBackBtn")
        self.hist_back_btn = self._nav_btn("◀", "后退")
        self.hist_forward_btn = self._nav_btn("▶", "前进")
        self.reload_btn = self._nav_btn("⟳", "刷新当前页面")
        self.home_btn = self._nav_btn("🏠", "回到新闻首页")
        self.address = QLineEdit(self)
        self.address.setObjectName("BrowserAddress")
        self.address.setReadOnly(True)
        self.edge_btn = self._nav_btn("在 Edge 打开", "用系统默认浏览器打开当前页")
        for btn in (self.back_btn, self.hist_back_btn, self.hist_forward_btn,
                    self.reload_btn, self.home_btn):
            bar.addWidget(btn)
        bar.addWidget(self.address, 1)
        bar.addWidget(self.edge_btn)
        layout.addLayout(bar)

        self._view_holder = QWidget(self)
        holder_layout = QVBoxLayout(self._view_holder)
        holder_layout.setContentsMargins(0, 0, 0, 0)
        holder_layout.setSpacing(0)
        layout.addWidget(self._view_holder, 1)

        self.back_btn.clicked.connect(self.back_requested)
        self.hist_back_btn.clicked.connect(self._go_back)
        self.hist_forward_btn.clicked.connect(self._go_forward)
        self.reload_btn.clicked.connect(self._reload)
        self.home_btn.clicked.connect(self._go_home)
        self.edge_btn.clicked.connect(self._open_external)
        self._update_nav()

    def _nav_btn(self, text: str, tip: str) -> QPushButton:
        btn = QPushButton(text, self)
        btn.setObjectName("BrowserBtn")
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setToolTip(tip)
        btn.setMinimumHeight(26)
        return btn

    # ---- 视图（懒创建，之后复用；返回待机也不销毁，再点开是同一个页面）----
    def ensure_view(self) -> bool:
        """创建 Chromium 视图；内核不可用时返回 False（调用方回退 Edge）。"""
        if self.view is not None:
            return True
        try:
            from PyQt6.QtWebEngineWidgets import QWebEngineView
            from PyQt6.QtWebEngineCore import QWebEnginePage
        except Exception:
            return False

        class _BrowserView(QWebEngineView):
            def createWindow(self, _window_type):
                # 新闻站站内链接常要求“新窗口”，这里复用当前视图，避免点了没反应
                return self

        view = _BrowserView(self._view_holder)
        profile = browser_profile()
        if profile is not None:
            try:
                view.setPage(QWebEnginePage(profile, view))
            except Exception:
                pass
        view.urlChanged.connect(self._on_url_changed)
        view.loadFinished.connect(self._on_load_finished)
        self._view_holder.layout().addWidget(view)
        self.view = view
        self._update_nav()
        return True

    def open_url(self, url: str = "", home_url: str = "") -> bool:
        """加载指定页面（url 为空时退回新闻首页 / 上次页面）。"""
        if home_url:
            self._home_url = home_url
        if not self.ensure_view():
            return False
        target = str(url or "").strip()
        if target:
            self.view.setUrl(QUrl(target))
        elif self._home_url:
            self.view.setUrl(QUrl(self._home_url))
        elif self._last_url:
            self.view.setUrl(QUrl(self._last_url))
        self._update_nav()
        return True

    def stop_page(self) -> None:
        """滑出后停止加载，避免在后台白白占用网络 / CPU。"""
        if self.view is not None:
            try:
                self.view.stop()
            except Exception:
                pass

    def _go_back(self) -> None:
        if self.view is not None:
            self.view.back()

    def _go_forward(self) -> None:
        if self.view is not None:
            self.view.forward()

    def _reload(self) -> None:
        if self.view is not None:
            self.view.reload()

    def _go_home(self) -> None:
        if self.view is not None and self._home_url:
            self.view.setUrl(QUrl(self._home_url))
        self._update_nav()

    def _open_external(self) -> None:
        url = self.view.url().toString() if self.view is not None else ""
        if url:
            QDesktopServices.openUrl(QUrl(url))

    def _on_url_changed(self, url: QUrl) -> None:
        text = url.toString()
        self._last_url = text
        self.address.setText(text)
        self.address.setCursorPosition(0)
        self._update_nav()

    def _update_nav(self) -> None:
        history = self.view.history() if self.view is not None else None
        can_back = history.canGoBack() if history is not None else False
        can_forward = history.canGoForward() if history is not None else False
        self.hist_back_btn.setEnabled(can_back)
        self.hist_forward_btn.setEnabled(can_forward)
        self.reload_btn.setEnabled(self.view is not None)
        self.home_btn.setEnabled(self.view is not None and bool(self._home_url))
        self.edge_btn.setEnabled(self.view is not None)

    def _on_load_finished(self, ok: bool) -> None:
        """加载失败时给出退路提示（很多新闻站会拦内嵌浏览器）。"""
        if ok:
            self.address.setToolTip("")
            return
        self.address.setToolTip(
            "该页面可能限制了内嵌浏览：可点「⟳」重试，"
            "或点「在 Edge 打开」用系统浏览器查看")

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.back_requested.emit()
            event.accept()
            return
        super().keyPressEvent(event)


# ---------------------------------------------------------------------------
# 待机画面（取代原来的“精简模式”：收起导航 + 右侧只留一块优雅的待机区）
# ---------------------------------------------------------------------------
class StandbyView(QWidget):
    """待机画面：大时钟 + 多城市天气预报轮询 + 新闻热点。

    天气直接复用定时任务里的和风天气模块（api_key / base_url 来自 weather.json），
    城市列表单独存放（standby_cities），可手动切换 / 轮询 / 刷新；
    热点复用 news_fetcher，支持知乎 / 头条 / B站 / 微博四源切换。

    条目点击默认在「内置浏览器」里打开（需安装 PyQt6-WebEngine）；
    如果没装 WebEngine，则自动回退到系统默认浏览器（Edge），
    保证不安装额外内核也照样能用。
    """

    weather_ready = pyqtSignal(dict)
    weather_error = pyqtSignal(str)
    # 一批城市全部拉取完成（含失败），用于解除 busy 锁
    weather_batch_done = pyqtSignal()
    # 平台 key + 条目列表：异步返回时校验平台，避免切换来源后旧结果覆盖新内容
    news_ready = pyqtSignal(str, list)
    # 待机内容高度变化（新闻加载完成）：主窗口据此再自适应一次窗口大小
    content_updated = pyqtSignal()
    # 待机页内置浏览器打开（True）/ 返回待机（False）：主窗口据此放大 / 收回窗口
    browser_open_changed = pyqtSignal(bool)

    _TIPS = (
        "后台任务与监控持续运行中，随时点左侧图标返回编辑区",
        "发送进度与运行日志一直在下方实时刷新",
        "数据表更新后可直接回到「数据发送」重新执行",
        "定时任务与轮询监控在后台默默值守",
    )
    _WEATHER_ICONS = (
        ("雷", "⛈"), ("雪", "❄"), ("雾", "🌫"), ("霾", "🌫"),
        ("雨", "🌧"), ("阴", "☁"), ("多云", "⛅"), ("晴", "☀"),
    )
    # 多城市轮询间隔：切换过快要一直盯着看，放宽到 20 秒
    _CITY_ROTATE_MS = 20000
    # 各新闻源的首页（内置浏览器「🏠」按钮与「🌐 打开新闻页」都指向这里）
    _NEWS_HOME = {
        "zhihu": "https://www.zhihu.com/hot",
        "toutiao": "https://www.toutiao.com/",
        "bilibili": "https://www.bilibili.com/v/popular/all",
        "weibo": "https://s.weibo.com/top/summary",
    }

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("StandbyView")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, False)
        self._phase = 0.0
        self._cities: List[str] = []
        self._city_index = 0
        self._weather: Dict = {}
        self._weather_cache: Dict[str, Dict] = {}
        self._news_items: List[Dict] = []
        self._news_platform = news_fetcher.load_standby_platform()
        self._refresh_minutes = weather_fetcher.load_standby_refresh_minutes()
        self._weather_busy = False
        self._news_busy = False
        # 上次真实取数时间：进待机时只在数据过期时才后台刷新，
        # 避免每次切到待机画面都重新请求天气/热点
        self._weather_updated_at = 0.0
        self._news_updated_at = 0.0
        # 内置浏览器内嵌面板（懒创建、复用同一个实例；返回待机也不销毁）
        self._inline_browser: Optional[StandbyBrowser] = None
        self._browser_anim: Optional[QPropertyAnimation] = None
        self._browser_open = False
        # 滑入动画结束后再加载的页面（首次创建 Chromium 视图较慢，放到动画之后）
        self._pending_browser_load: Optional[Tuple[str, str]] = None
        # 手动刷新 / 后台定时刷新时是否忽略天气模块 10 分钟内存缓存
        self._force_refresh = False
        # 天气文字淡入淡出：切换城市 / 数据到达时不再“硬跳”
        self._weather_opacity = 1.0
        self._weather_fade: Optional[QVariantAnimation] = None

        self._phase_anim = QVariantAnimation(self)
        self._phase_anim.setStartValue(0.0)
        self._phase_anim.setEndValue(1.0)
        self._phase_anim.setDuration(18000)
        self._phase_anim.setLoopCount(-1)
        self._phase_anim.setEasingCurve(QEasingCurve.Type.Linear)
        self._phase_anim.valueChanged.connect(self._on_phase)

        self._tick = QTimer(self)
        self._tick.setInterval(1000)
        self._tick.timeout.connect(self._update_text)

        # 多城市轮询：每 20 秒切下一个城市（只切显示，不重复请求）；
        # 后台天气真实刷新走独立定时器（默认 2 小时，可在待机页修改）
        self._rotate_timer = QTimer(self)
        self._rotate_timer.setInterval(self._CITY_ROTATE_MS)
        self._rotate_timer.timeout.connect(self._on_rotate_tick)
        self._refresh_timer = QTimer(self)
        self._refresh_timer.setInterval(self._refresh_minutes * 60 * 1000)
        self._refresh_timer.timeout.connect(self._on_refresh_tick)

        self.weather_ready.connect(self._apply_weather)
        self.weather_error.connect(self._apply_weather_error)
        self.weather_batch_done.connect(self._on_weather_batch_done)
        self.news_ready.connect(self._apply_news)

        # 待机内容放进滚动区：窗口被缩得很小时也能滚动查看，而不是被裁掉
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea(self)
        scroll.setObjectName("StandbyScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.viewport().setAutoFillBackground(False)
        outer.addWidget(scroll)
        self._scroll_widget = scroll

        content = QWidget()
        content.setObjectName("StandbyContent")
        scroll.setWidget(content)
        self._content = content
        layout = QVBoxLayout(content)
        layout.setContentsMargins(20, 14, 20, 14)
        layout.setSpacing(4)
        layout.addStretch(1)

        self.icon_label = QLabel("🌤", content)
        self.icon_label.setObjectName("StandbyIcon")
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_label)

        self.clock_label = QLabel("--:--", content)
        self.clock_label.setObjectName("StandbyClock")
        self.clock_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.clock_label)

        self.date_label = QLabel("", content)
        self.date_label.setObjectName("StandbyDate")
        self.date_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.date_label)

        self.weather_main_label = QLabel("🌤 天气加载中…", content)
        self.weather_main_label.setObjectName("StandbyWeather")
        self.weather_main_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.weather_main_label.setWordWrap(True)
        layout.addWidget(self.weather_main_label)

        self.weather_detail_label = QLabel("", content)
        self.weather_detail_label.setObjectName("StandbyWeatherSub")
        self.weather_detail_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.weather_detail_label.setWordWrap(True)
        layout.addWidget(self.weather_detail_label)

        # 淡入淡出效果：切换城市时先淡出旧城市，再淡入新城市，避免文字硬跳
        self._weather_effect_main = QGraphicsOpacityEffect(self.weather_main_label)
        self._weather_effect_main.setOpacity(1.0)
        self.weather_main_label.setGraphicsEffect(self._weather_effect_main)
        self._weather_effect_detail = QGraphicsOpacityEffect(self.weather_detail_label)
        self._weather_effect_detail.setOpacity(1.0)
        self.weather_detail_label.setGraphicsEffect(self._weather_effect_detail)
        self._weather_fade = QVariantAnimation(self)
        self._weather_fade.setDuration(130)
        self._weather_fade.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._weather_fade.valueChanged.connect(self._on_weather_opacity)

        self.news_title_label = QLabel(
            f"📰 今日{news_fetcher.platform_label(self._news_platform)}"
            f" · {self._open_hint()}", content)
        self.news_title_label.setObjectName("StandbySection")
        self.news_title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.news_title_label)

        self.news_label = QLabel("热点加载中…", content)
        self.news_label.setObjectName("StandbyNews")
        self.news_label.setTextFormat(Qt.TextFormat.RichText)
        self.news_label.setOpenExternalLinks(False)  # 交给内置浏览器 / Edge 处理
        self.news_label.linkActivated.connect(self._open_link)
        self.news_label.setWordWrap(True)
        self.news_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.news_label)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(6)
        btn_row.addStretch(1)
        self.next_city_btn = self._action_btn("⏭ 下一城市", "切换到下一个城市")
        self.refresh_btn = self._action_btn("↻ 刷新", "立即刷新天气与热点")
        self.city_btn = self._action_btn("📍 城市", "设置待机轮询显示的城市（1~8 个）")
        self.news_btn = self._action_btn("📰 来源", "切换新闻来源 / 立即刷新热点")
        self.interval_btn = self._action_btn(
            self._interval_text(self._refresh_minutes),
            "设置后台自动更新天气的间隔（默认 2 小时）")
        for btn in (self.next_city_btn, self.refresh_btn,
                    self.city_btn, self.interval_btn, self.news_btn):
            btn_row.addWidget(btn)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)
        self.next_city_btn.clicked.connect(self._rotate_to_next)
        self.refresh_btn.clicked.connect(self._refresh_all)
        self.city_btn.clicked.connect(self._edit_cities)
        self._build_interval_menu()
        self._build_news_menu()

        self.tip_label = QLabel(self._TIPS[0], content)
        self.tip_label.setObjectName("StandbyTip")
        self.tip_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tip_label.setWordWrap(True)
        layout.addWidget(self.tip_label)

        layout.addStretch(1)
        self._update_text()
        self._sync_city_ui()

    def _action_btn(self, text: str, tip: str) -> QPushButton:
        btn = QPushButton(text)
        btn.setObjectName("StandbyAction")
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setToolTip(tip)
        btn.setMinimumHeight(24)
        return btn

    # ---- 尺寸提示（主窗口进入待机时按内容自适应窗口大小）----
    def content_widget(self) -> QWidget:
        """滚动区里真正承载待机内容的控件。"""
        return self._content

    def scroll_area(self) -> QScrollArea:
        return self._scroll_widget

    def preferred_height(self, width: int) -> int:
        """在给定内容宽度下，待机画面不被裁切所需的高度（含内边距）。"""
        w = max(320, int(width))
        layout = self._content.layout()
        if layout is not None and layout.hasHeightForWidth():
            h = int(layout.heightForWidth(w))
            if h > 0:
                return h
        return int(self._content.sizeHint().height())

    # ---- 待机页设置菜单：新闻来源 / 后台更新间隔 ----
    @staticmethod
    def _interval_text(minutes: int) -> str:
        if minutes >= 60 and minutes % 60 == 0:
            return f"⏱ {minutes // 60}小时"
        return f"⏱ {minutes}分钟"

    @staticmethod
    def _interval_plain(minutes: int) -> str:
        if minutes >= 60 and minutes % 60 == 0:
            return f"{minutes // 60} 小时"
        return f"{minutes} 分钟"

    def _build_interval_menu(self) -> None:
        """后台天气刷新间隔菜单：30 分钟 / 1 小时 / 2 小时 / 3 小时 / 6 小时。"""
        menu = QMenu(self.interval_btn)
        group = QActionGroup(menu)
        group.setExclusive(True)
        for minutes in weather_fetcher.STANDBY_REFRESH_OPTIONS:
            act = menu.addAction(self._interval_text(minutes).replace("⏱ ", ""))
            act.setCheckable(True)
            act.setChecked(minutes == self._refresh_minutes)
            act.triggered.connect(
                lambda _checked=False, m=minutes: self._set_refresh_minutes(m))
            group.addAction(act)
        self.interval_btn.setMenu(menu)

    def _set_refresh_minutes(self, minutes: int) -> None:
        try:
            self._refresh_minutes = weather_fetcher.save_standby_refresh_minutes(minutes)
        except Exception:
            self._refresh_minutes = minutes
        self._refresh_timer.setInterval(self._refresh_minutes * 60 * 1000)
        if self.isVisible():
            self._refresh_timer.start()
        self.interval_btn.setText(self._interval_text(self._refresh_minutes))
        self.interval_btn.setToolTip(
            f"后台每 {self._refresh_minutes} 分钟自动刷新天气；切换城市不会触发请求")
        self._render_weather(animated=False)

    def _build_news_menu(self) -> None:
        """新闻来源菜单：切换平台 + 立即刷新热点。"""
        menu = QMenu(self.news_btn)
        group = QActionGroup(menu)
        group.setExclusive(True)
        self._news_actions: Dict[str, QAction] = {}
        for key in ("zhihu", "toutiao", "bilibili", "weibo"):
            act = menu.addAction(news_fetcher.platform_label(key))
            act.setCheckable(True)
            act.setChecked(key == self._news_platform)
            act.triggered.connect(
                lambda _checked=False, k=key: self._set_news_platform(k))
            group.addAction(act)
            self._news_actions[key] = act
        menu.addSeparator()
        browse_act = menu.addAction("🌐 打开新闻页")
        browse_act.triggered.connect(self._open_news_page)
        refresh_act = menu.addAction("↻ 立即刷新热点")
        refresh_act.triggered.connect(self._refresh_news)
        self.news_btn.setMenu(menu)
        self.news_btn.setToolTip(
            f"当前新闻来源：{news_fetcher.platform_label(self._news_platform)}"
            "（点击可切换知乎 / 头条 / B站 / 微博）")

    def _set_news_platform(self, platform: str) -> None:
        platform = news_fetcher.save_standby_platform(platform)
        self._news_platform = platform
        if getattr(self, "_news_actions", None):
            act = self._news_actions.get(platform)
            if act is not None:
                act.setChecked(True)
        label = news_fetcher.platform_label(platform)
        self.news_title_label.setText(f"📰 今日{label} · {self._open_hint()}")
        self.news_btn.setToolTip(
            f"当前新闻来源：{label}（点击可切换知乎 / 头条 / B站 / 微博）")
        self.news_label.setText(f"{label}加载中…")
        self._fetch_news_async(force=True)

    # ---- 生命周期 ----
    def showEvent(self, event):
        super().showEvent(event)
        self._update_text()
        self._tick.start()
        self._phase_anim.start()
        self._rotate_timer.start()
        self._refresh_timer.start()
        self.refresh_now()

    def hideEvent(self, event):
        self._tick.stop()
        self._phase_anim.stop()
        self._rotate_timer.stop()
        self._refresh_timer.stop()
        if self._weather_fade is not None:
            self._weather_fade.stop()
        self._apply_weather_opacity(1.0)
        super().hideEvent(event)

    def _on_phase(self, value) -> None:
        self._phase = float(value)
        self.update()

    def _update_text(self) -> None:
        now = datetime.datetime.now()
        weekdays = "一二三四五六日"
        self.clock_label.setText(now.strftime("%H:%M:%S"))
        self.date_label.setText(
            f"{now.year} 年 {now.month} 月 {now.day} 日　星期{weekdays[now.weekday()]}")
        self.tip_label.setText(self._greeting(now.hour) + "。" + self._TIPS[now.minute % len(self._TIPS)])

    @staticmethod
    def _greeting(hour: int) -> str:
        if hour < 5:
            return "夜深了，程序还在替你值守"
        if hour < 9:
            return "早上好，新的一天从整理数据开始"
        if hour < 12:
            return "上午好，发送任务已经就绪"
        if hour < 14:
            return "中午好，记得稍作休息"
        if hour < 18:
            return "下午好，表格与消息都在有序流转"
        if hour < 23:
            return "晚上好，今天的任务正在收尾"
        return "夜深了，程序还在替你值守"

    # ---- 城市列表 / 手动操作 ----
    def _current_city(self) -> str:
        if not self._cities:
            return ""
        index = max(0, min(self._city_index, len(self._cities) - 1))
        return self._cities[index]

    def _sync_city_ui(self) -> None:
        city = self._current_city()
        if len(self._cities) > 1:
            self.city_btn.setText(
                f"📍 {city} {self._city_index + 1}/{len(self._cities)}")
        elif city:
            self.city_btn.setText(f"📍 {city}")
        else:
            self.city_btn.setText("📍 城市")
        self.city_btn.setToolTip(
            "当前轮询城市：" + ("、".join(self._cities) if self._cities else "未设置")
            + "\n点击可修改（逗号或换行分隔，最多 8 个）")
        self.next_city_btn.setEnabled(len(self._cities) > 1)

    def _reload_cities(self) -> None:
        self._cities = weather_fetcher.load_standby_cities()
        if self._cities:
            self._city_index %= len(self._cities)
        else:
            self._city_index = 0
        self._sync_city_ui()

    def _show_city(self, index: int) -> None:
        """切换到指定城市：优先用本地缓存，无缓存才补拉这一个城市。"""
        if not self._cities:
            return
        self._city_index = index % len(self._cities)
        self._sync_city_ui()
        city = self._current_city()
        cached = self._weather_cache.get(city)
        if cached:
            self._weather = cached
            self._render_weather(animated=True)
            return
        self._fetch_weather_async(city=city)

    def _rotate_to_next(self) -> None:
        if not self._cities:
            self._edit_cities()
            return
        self._show_city(self._city_index + 1)

    def _on_rotate_tick(self) -> None:
        # 只切换显示的城市，不发起网络请求（请求由后台刷新定时器负责）
        if len(self._cities) > 1:
            self._show_city(self._city_index + 1)

    def _edit_cities(self) -> None:
        current = "，".join(self._cities)
        text, ok = QInputDialog.getText(
            self, "待机多城市天气",
            "输入要轮询显示的城市（逗号或换行分隔，最多 8 个）：",
            text=current)
        if not ok:
            return
        parts = [p.strip() for p in re.split(r"[，,、;；\s]+", text or "") if p.strip()]
        try:
            self._cities = weather_fetcher.save_standby_cities(parts)
        except Exception:
            self._cities = weather_fetcher.load_standby_cities()
        self._city_index = 0
        self._weather_cache.clear()
        self._sync_city_ui()
        self._fetch_weather_async(prefetch_all=True)

    def _refresh_all(self, force: bool = True) -> None:
        self._reload_cities()
        self._fetch_weather_async(prefetch_all=True, force=force)
        self._fetch_news_async(force=True)

    def _on_refresh_tick(self) -> None:
        """后台定时刷新：一次性更新所有城市缓存与新闻，不打扰当前画面。"""
        self._refresh_all(force=True)

    def _refresh_news(self) -> None:
        self._fetch_news_async(force=True)

    # ---- 后台拉取（不阻塞 UI）----
    @staticmethod
    def _net_allowed() -> bool:
        if os.environ.get("QT_QPA_PLATFORM", "").lower() == "offscreen":
            return False
        return os.environ.get("WXUI_NO_NETWORK", "") != "1"

    def _fetch_weather_async(self, city: str = "", force: bool = False,
                             prefetch_all: bool = False) -> None:
        """后台拉取天气；prefetch_all 时把所有轮询城市一次性拉完。

        说明：城市轮换只切显示，不请求网络；只有首次无缓存、手动刷新
        和后台定时刷新（默认 2 小时）才会真正发起请求。
        """
        city = city or self._current_city()
        if not city:
            self._animate_weather_text(
                "🌤 未设置待机城市",
                "点下方「📍 城市」添加 1~8 个城市，即可自动轮询显示天气预报")
            return
        cfg = weather_fetcher.load_weather_config()
        api_key = str(cfg.get("api_key") or "").strip()
        base_url = str(cfg.get("base_url") or "")
        if not api_key:
            self._animate_weather_text(
                f"🌤 {city}　天气未启用",
                "和风天气 API Key 未配置：在「定时发送 → 天气设置」里填写后即可显示")
            return
        if not self._net_allowed():
            self._animate_weather_text(f"🌤 {city}", "离线环境：天气请求已跳过")
            return
        if self._weather_busy:
            return
        self._weather_busy = True
        cities = list(self._cities) if (prefetch_all and self._cities) else [city]
        if not prefetch_all:
            self._animate_weather_text(
                f"🌤 {city}　天气加载中…", "正在获取实时天气与今日预报…")
        threading.Thread(target=self._weather_worker,
                         args=(cities, api_key, base_url, force),
                         daemon=True).start()

    def _weather_worker(self, cities: List[str], api_key: str,
                        base_url: str, force: bool) -> None:
        try:
            for city in cities:
                try:
                    data = weather_fetcher.fetch_weather(
                        city, base_url, api_key, force=force)
                except Exception:
                    data = None
                if data:
                    self.weather_ready.emit(data)
                else:
                    self.weather_error.emit(city)
        finally:
            self.weather_batch_done.emit()

    def _on_weather_batch_done(self) -> None:
        self._weather_busy = False

    def _apply_weather(self, data) -> None:
        data = dict(data or {})
        city = str(data.get("city") or "")
        if city:
            self._weather_cache[city] = data
        self._weather_updated_at = time.time()
        if city != self._current_city():
            return  # 后台预取的其它城市：只入缓存，等轮换到它再显示
        self._weather = data
        self._render_weather(animated=True)

    def _apply_weather_error(self, city: str) -> None:
        if city != self._current_city():
            return
        self._animate_weather_text(
            f"🌤 {city}　天气获取失败",
            "请检查网络或和风天气 Key；点「↻ 刷新」可重试")

    # ---- 天气文字：淡入淡出 + 内容拼装 ----
    def _on_weather_opacity(self, value) -> None:
        self._apply_weather_opacity(float(value))

    def _apply_weather_opacity(self, opacity: float) -> None:
        self._weather_opacity = max(0.0, min(1.0, float(opacity)))
        self._weather_effect_main.setOpacity(self._weather_opacity)
        self._weather_effect_detail.setOpacity(self._weather_opacity)

    def _set_weather_texts(self, head: str, detail: str) -> None:
        self.weather_main_label.setText(head)
        self.weather_detail_label.setText(detail)

    def _fade_weather_to(self, target: float, duration: int,
                         on_finished=None) -> None:
        anim = self._weather_fade
        if anim is None:
            self._apply_weather_opacity(target)
            if on_finished is not None:
                on_finished()
            return
        anim.stop()
        try:
            anim.finished.disconnect()
        except (TypeError, RuntimeError):
            pass
        anim.setDuration(int(duration))
        anim.setStartValue(self._weather_opacity)
        anim.setEndValue(float(target))
        if on_finished is not None:
            anim.finished.connect(on_finished)
        anim.start()

    def _animate_weather_text(self, head: str, detail: str) -> None:
        """旧文字淡出 → 替换文字 → 新文字淡入（不可见时直接替换）。"""
        if not self.isVisible():
            self._set_weather_texts(head, detail)
            self._apply_weather_opacity(1.0)
            return

        def _swap() -> None:
            self._set_weather_texts(head, detail)
            self._fade_weather_to(1.0, 240)

        self._fade_weather_to(0.0, 140, _swap)

    def _render_weather(self, animated: bool = True) -> None:
        data = self._weather
        if not data:
            return
        head, detail = self._format_weather(data)
        if animated:
            self._animate_weather_text(head, detail)
        else:
            self._set_weather_texts(head, detail)
            self._apply_weather_opacity(1.0)

    def _format_weather(self, data: Dict) -> Tuple[str, str]:
        """把天气数据拼成「头部一行 + 详细多行」。"""
        city = str(data.get("city") or self._current_city())
        text = str(data.get("text") or "")
        temp = str(data.get("temp") or "")
        feels = str(data.get("feels_like") or "")
        icon = "🌤"
        for key, emoji in self._WEATHER_ICONS:
            if key in text:
                icon = emoji
                break
        head = f"{icon} {city}　{text}".strip()
        if temp:
            head += f"　{temp}°C"
        if feels:
            head += f"（体感 {feels}°C）"

        lines: List[str] = []
        max_t = str(data.get("max_temp") or "")
        min_t = str(data.get("min_temp") or "")
        day_t = str(data.get("day_text") or "")
        night_t = str(data.get("night_text") or "")
        if max_t or min_t:
            rng = f"今日 {min_t}~{max_t}°C"
            if day_t or night_t:
                rng += f" · 白天{day_t or '—'} / 夜间{night_t or '—'}"
            lines.append(rng)
        extras = []
        if data.get("wind"):
            extras.append(str(data.get("wind")))
        if data.get("humidity"):
            extras.append(f"湿度 {data.get('humidity')}%")
        if data.get("uv_index"):
            extras.append(f"紫外线 {data.get('uv_index')}")
        if extras:
            lines.append(" · ".join(extras))
        sun = []
        if data.get("sunrise"):
            sun.append(f"日出 {self._hm(data.get('sunrise'))}")
        if data.get("sunset"):
            sun.append(f"日落 {self._hm(data.get('sunset'))}")
        if sun:
            lines.append(" · ".join(sun))
        advice = str(data.get("advice") or "").strip()
        if advice:
            first = next((ln.strip("· ").strip() for ln in advice.splitlines()
                          if ln.strip()), "")
            if first:
                if len(first) > 34:
                    first = first[:34] + "…"
                lines.append(f"💡 {first}")
        updated = self._hm(data.get("update_time"))
        tail = f"更新 {updated}" if updated else ""
        meta = []
        if len(self._cities) > 1:
            meta.append(f"城市 {self._city_index + 1}/{len(self._cities)}")
        meta.append(f"后台每 {self._interval_plain(self._refresh_minutes)}刷新")
        if meta:
            tail += ("　·　" if tail else "") + "　·　".join(meta)
        if tail:
            lines.append(tail)
        return head, "\n".join(lines)

    @staticmethod
    def _hm(value) -> str:
        s = str(value or "")
        if "T" in s and len(s) >= 16:
            return s[11:16]
        return s

    def _fetch_news_async(self, force: bool = False) -> None:
        if not self._net_allowed() or self._news_busy:
            return
        self._news_busy = True
        # 已有内容时后台刷新不显示“加载中”，避免每次进待机都闪一下
        if not self._news_items:
            self.news_label.setText("热点加载中…")
        threading.Thread(target=self._news_worker,
                         args=(self._news_platform,), daemon=True).start()

    def _news_worker(self, platform: str) -> None:
        try:
            items = news_fetcher.fetch_top(platform, 6)
        except Exception:
            items = []
        self.news_ready.emit(platform, items or [])

    def _apply_news(self, platform: str, items) -> None:
        self._news_busy = False
        if platform != self._news_platform:
            return  # 切换来源后旧平台迟到的结果，丢弃
        self._news_items = list(items or [])[:6]
        self._news_updated_at = time.time()
        self._render_news()
        self.content_updated.emit()

    def _render_news(self) -> None:
        c = _colors()
        if not self._news_items:
            self.news_label.setText(
                f'<span style="color:{c["muted"]};">热点暂不可用，'
                '点「📰 来源」→「立即刷新热点」重试</span>')
            return
        rows: List[str] = []
        for item in self._news_items:
            title = html.escape(str(item.get("title") or ""))
            url = str(item.get("url") or "").strip()
            rank = item.get("rank") or 0
            if url:
                rows.append(
                    f'<div>{rank}. <a href="{html.escape(url, quote=True)}" '
                    f'style="color:{c["accent"]};text-decoration:none;">{title}</a></div>')
            else:
                rows.append(f"<div>{rank}. {title}</div>")
        self.news_label.setText("".join(rows))

    def _on_news_link(self, url: str) -> None:
        """兼容旧调用点：等价于 _open_link。"""
        self._open_link(url)

    # ---- 内置浏览器 ----
    @staticmethod
    def _open_hint() -> str:
        """标题提示：装了 WebEngine 就提示在待机页内打开，否则提示用 Edge。"""
        return ("点击标题在右侧打开" if webengine_installed()
                else "点击标题用 Edge 打开")

    def _news_home_url(self) -> str:
        return self._NEWS_HOME.get(self._news_platform, "")

    # ---- 内置浏览器（嵌在待机画面里滑入 / 滑出，不弹独立窗口）----
    def is_browser_open(self) -> bool:
        return bool(getattr(self, "_browser_open", False))

    def _browser_rect(self) -> QRect:
        """浏览器面板占用的区域：与待机滚动区完全重合。"""
        rect = self._scroll_widget.geometry()
        if rect.width() <= 0 or rect.height() <= 0:
            rect = self.rect()
        return rect

    def _ensure_inline_browser(self) -> StandbyBrowser:
        if self._inline_browser is None:
            self._inline_browser = StandbyBrowser(self)
            self._inline_browser.back_requested.connect(self.close_browser)
            self._inline_browser.hide()
        return self._inline_browser

    def _show_browser(self, url: str = "", home_url: str = "") -> None:
        """在待机画面右侧打开新闻页；没有 WebEngine 时回退到系统 Edge。"""
        home_url = home_url or self._news_home_url()
        if not webengine_installed():
            QDesktopServices.openUrl(QUrl(url or home_url))
            return
        panel = self._ensure_inline_browser()
        if self._browser_open:
            # 浏览器已经打开：只切换页面，不重复播放滑入动画
            panel.raise_()
            panel.setFocus()
            self._load_in_browser(panel, url, home_url)
            return
        # 先滑入再创建 Chromium 视图：首次打开要加载内核，放在动画之后不会卡住过渡
        self._pending_browser_load = (url, home_url)
        self._slide_browser(show=True)

    def _load_in_browser(self, panel: "StandbyBrowser",
                         url: str, home_url: str) -> None:
        """加载页面；内核不可用 / 加载异常时收起面板并回退到系统 Edge。"""
        if not self._browser_open:
            return
        try:
            if panel.open_url(url=url, home_url=home_url):
                return
        except Exception:
            pass
        self._pending_browser_load = None
        self.close_browser(animate=False)
        QDesktopServices.openUrl(QUrl(url or home_url))

    def close_browser(self, animate: bool = True) -> None:
        """收起内置浏览器，回到待机画面（Esc / 「← 待机」都会走到这里）。"""
        panel = self._inline_browser
        if panel is None or not self._browser_open:
            return
        self._pending_browser_load = None
        self._stop_browser_anim()
        if not animate:
            self._browser_open = False
            panel.hide()
            panel.stop_page()
            self.browser_open_changed.emit(False)
            return
        self._slide_browser(show=False)

    def _slide_browser(self, show: bool) -> None:
        """滑入 / 滑出过渡：面板从右侧边界滑进来，返回时再滑出去。"""
        panel = self._inline_browser
        if panel is None:
            return
        rect = self._browser_rect()
        panel.setGeometry(rect)
        if show:
            self._browser_open = True
            start = QPoint(rect.x() + rect.width(), rect.y())
            end = QPoint(rect.x(), rect.y())
        else:
            self._browser_open = False
            start = panel.pos()
            end = QPoint(rect.x() + rect.width(), rect.y())
        anim = QPropertyAnimation(panel, b"pos", panel)
        anim.setDuration(240)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.setStartValue(start)
        anim.setEndValue(end)
        anim.finished.connect(self._on_browser_slide_finished)
        self._browser_anim = anim
        panel.move(start)
        panel.show()
        panel.raise_()
        anim.start()

    def _stop_browser_anim(self) -> None:
        anim = self._browser_anim
        if anim is None:
            return
        self._browser_anim = None
        try:
            anim.stop()
        except Exception:
            pass
        anim.deleteLater()

    def _on_browser_slide_finished(self) -> None:
        panel = self._inline_browser
        self._browser_anim = None
        if panel is None:
            return
        if self._browser_open:
            # 滑入结束：贴合容器，再加载页面（主窗口随后放大，下一次 resize 再同步）
            panel.setGeometry(self._browser_rect())
            panel.setFocus()
            pending = self._pending_browser_load
            self._pending_browser_load = None
            if pending is not None:
                self._load_in_browser(panel, pending[0], pending[1])
                if not self._browser_open:
                    return
            self.browser_open_changed.emit(True)
            return
        panel.hide()
        panel.stop_page()
        self.browser_open_changed.emit(False)

    def _sync_browser_geometry(self) -> None:
        """窗口尺寸变化时让面板贴合容器（滑动动画进行中先不打断）。"""
        panel = self._inline_browser
        if panel is None or not self._browser_open:
            return
        anim = self._browser_anim
        if anim is not None and anim.state() == QAbstractAnimation.State.Running:
            return
        if panel.isVisible():
            panel.setGeometry(self._browser_rect())

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._sync_browser_geometry()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape and self._browser_open:
            self.close_browser()
            event.accept()
            return
        super().keyPressEvent(event)

    def _open_link(self, url: str) -> None:
        """点击新闻条目的链接。"""
        url = str(url or "").strip()
        if not url:
            return
        self._show_browser(url=url)

    def _open_news_page(self) -> None:
        """「🌐 打开新闻页」：直接打开当前来源的热榜首页。"""
        self._show_browser(home_url=self._news_home_url())

    # ---- 绘制 ----
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        paint_glass(painter, rect, 16.0, strong=True)

        # 缓慢漂移的柔光球：待机时画面不呆板，也不抢时钟的视觉重心
        c = _colors()
        accent = ui_theme.qcolor(c["accent"])
        diameter = min(rect.width(), rect.height())
        painter.setPen(Qt.PenStyle.NoPen)
        for alpha, scale, phase in ((26, 0.34, 0.0), (18, 0.26, 0.34),
                                    (14, 0.18, 0.67)):
            t = (self._phase + phase) % 1.0
            angle = 2.0 * math.pi * t
            cx = rect.center().x() + rect.width() * 0.30 * math.cos(angle)
            cy = rect.center().y() + rect.height() * 0.26 * math.sin(angle)
            glow = QColor(accent)
            glow.setAlpha(alpha)
            painter.setBrush(glow)
            radius = diameter * scale * 0.5
            painter.drawEllipse(QPointF(cx, cy), radius, radius)

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == event.Type.PaletteChange:
            self.update()
            self._render_news()
            self._render_weather(animated=False)

    def refresh_now(self) -> None:
        """进入待机：先立刻显示缓存，只有数据过期时才后台更新。

        以前每次进入待机都会重新请求天气与热点，反复进出会看到画面
        跟着刷新；现在按「后台更新间隔」判断数据年龄，间隔内直接复用
        缓存，手动「↻ 刷新」与后台定时器仍会强制更新。
        """
        self._update_text()
        self._refresh_minutes = weather_fetcher.load_standby_refresh_minutes()
        self._refresh_timer.setInterval(self._refresh_minutes * 60 * 1000)
        self.interval_btn.setText(self._interval_text(self._refresh_minutes))
        self._reload_cities()
        now = time.time()
        stale_after = max(600.0, self._refresh_minutes * 60.0)
        cached = self._weather_cache.get(self._current_city())
        if cached:
            self._weather = cached
            self._render_weather(animated=False)
        if not cached or (now - self._weather_updated_at) >= stale_after:
            self._fetch_weather_async(prefetch_all=True)
        if not self._news_items or (now - self._news_updated_at) >= stale_after:
            self._fetch_news_async()
        self.update()


# ---------------------------------------------------------------------------
# 顶部玻璃标题栏
# ---------------------------------------------------------------------------
class GlassHeaderBar(QFrame):
    """应用内顶部标题栏（玻璃质感）：品牌、版本、当前页、快捷操作、窗口按钮。

    主窗口改为无边框后，最小化 / 最大化 / 关闭都收在这条自绘标题栏里；
    按住空白处可直接拖动窗口，双击空白处可最大化 / 还原。
    """

    theme_toggled = pyqtSignal()
    density_toggled = pyqtSignal()
    hide_requested = pyqtSignal()
    minimize_requested = pyqtSignal()
    maximize_requested = pyqtSignal()
    close_requested = pyqtSignal()

    def __init__(self, app_title: str, version: str,
                 parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("GlassHeader")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, False)
        self.setFixedHeight(52)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        # 拖动窗口（无边框窗口没有系统标题栏，拖动由这里接管）
        self._drag_active = False
        self._drag_maximized = False
        self._drag_ratio = 0.0
        self._drag_rel_y = 0
        self._drag_origin = QPointF()

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 6, 6, 6)
        layout.setSpacing(10)

        self.icon_label = QLabel("📊", self)
        self.icon_label.setObjectName("HeaderIcon")
        layout.addWidget(self.icon_label)

        self.title_label = QLabel(app_title, self)
        self.title_label.setObjectName("HeaderTitle")
        layout.addWidget(self.title_label)

        self.version_label = QLabel(f"v{version}", self)
        self.version_label.setObjectName("HeaderPill")
        layout.addWidget(self.version_label)

        self.page_label = QLabel("数据发送", self)
        self.page_label.setObjectName("HeaderStatus")
        layout.addWidget(self.page_label)

        layout.addStretch(1)

        self.density_btn = QPushButton("🌙 待机模式", self)
        self.density_btn.setObjectName("HeaderBtn")
        self.density_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.density_btn.setToolTip(
            "进入待机：收起导航，右侧只保留时钟画面（后台任务不受影响）")
        self.density_btn.clicked.connect(self.density_toggled.emit)
        layout.addWidget(self.density_btn)

        self.theme_btn = QPushButton("🌙 夜间模式", self)
        self.theme_btn.setObjectName("HeaderBtn")
        self.theme_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.theme_btn.setToolTip("切换日间 / 夜间主题（会记住选择）")
        self.theme_btn.clicked.connect(self.theme_toggled.emit)
        layout.addWidget(self.theme_btn)

        # 说明：原来这里还有一个「🗕 隐藏到托盘」按钮，现已并入最右侧的 ✕
        # （本程序关闭窗口本身就是隐藏到托盘、继续后台运行），避免标题栏放不下。

        # ---- 窗口控制按钮（无边框窗口自带的最小化/最大化/关闭）----
        layout.addSpacing(4)
        self.min_btn = self._win_btn("—", "最小化到任务栏")
        self.max_btn = self._win_btn("□", "最大化 / 还原窗口")
        self.close_btn = self._win_btn(
            "✕", "关闭窗口（有托盘时改为隐藏到托盘继续运行）", close=True)
        layout.addWidget(self.min_btn)
        layout.addWidget(self.max_btn)
        layout.addWidget(self.close_btn)
        self.min_btn.clicked.connect(self.minimize_requested.emit)
        self.max_btn.clicked.connect(self.maximize_requested.emit)
        self.close_btn.clicked.connect(self.close_requested.emit)

    def _win_btn(self, text: str, tip: str, close: bool = False) -> QPushButton:
        """标题栏窗口按钮（无边框、悬停变色；关闭按钮悬停为红）。"""
        btn = QPushButton(text, self)
        btn.setObjectName("WinBtnClose" if close else "WinBtn")
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setToolTip(tip)
        btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        btn.setFixedSize(34, 24)
        return btn

    # ---- 拖动窗口 / 双击最大化 ----
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            win = self.window()
            self._drag_active = True
            self._drag_maximized = win.isMaximized()
            self._drag_origin = event.globalPosition()
            self._drag_ratio = self._drag_origin.x() / max(1, win.width())
            self._drag_rel_y = int(self._drag_origin.y()) - win.frameGeometry().top()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if not self._drag_active or not (
                event.buttons() & Qt.MouseButton.LeftButton):
            super().mouseMoveEvent(event)
            return
        win = self.window()
        pos = event.globalPosition()
        if self._drag_maximized:
            # 最大化时拖动：先还原，再让光标保持在新窗口标题栏的相对位置
            win.showNormal()
            self._drag_maximized = False
            new_x = int(pos.x()) - int(win.width() * self._drag_ratio)
            new_y = int(pos.y()) - self._drag_rel_y
            win.move(new_x, new_y)
            self._drag_origin = QPointF(pos)
        else:
            delta = pos - self._drag_origin
            win.move(int(win.x() + delta.x()), int(win.y() + delta.y()))
            self._drag_origin = QPointF(pos)
        event.accept()

    def mouseReleaseEvent(self, event):
        self._drag_active = False
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_active = False
            self.maximize_requested.emit()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)

    def set_maximized(self, maximized: bool) -> None:
        """最大化按钮回显：最大化时显示“还原”图标。"""
        self.max_btn.setText("❐" if maximized else "□")
        self.max_btn.setToolTip("还原窗口" if maximized else "最大化窗口")

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        paint_glass(painter, rect, 14.0)

    def set_page_name(self, name: str) -> None:
        self.page_label.setText(name)

    def set_theme_mode(self, mode: str) -> None:
        """标题栏主题按钮回显：显示当前模式，提示下一次点击的切换结果。"""
        if mode == ui_theme.AUTO:
            self.theme_btn.setText("🌓 跟随系统")
            self.theme_btn.setToolTip("当前跟随系统深浅色（点击切到日间模式）")
        elif mode == ui_theme.DARK:
            self.theme_btn.setText("🌙 夜间模式")
            self.theme_btn.setToolTip("当前夜间模式（点击改为跟随系统）")
        else:
            self.theme_btn.setText("☀ 日间模式")
            self.theme_btn.setToolTip("当前日间模式（点击切到夜间模式）")

    def set_density(self, standby: bool) -> None:
        """待机按钮回显：standby=True 表示当前处于待机画面。"""
        if standby:
            self.density_btn.setText("📋 返回编辑")
            self.density_btn.setToolTip("退出待机，返回当前编辑页")
        else:
            self.density_btn.setText("🌙 待机模式")
            self.density_btn.setToolTip(
                "进入待机：收起导航，右侧只保留时钟画面（后台任务不受影响）")


# ---------------------------------------------------------------------------
# 无边框窗口的边缘缩放热区
# ---------------------------------------------------------------------------
class WindowEdgeHandle(QWidget):
    """无边框主窗口的边缘拖拽缩放条（透明，仅 6px，盖在外层留白上）。

    主窗口去掉系统标题栏后，Windows 不再提供边缘缩放，这里在窗口四周铺
    8 条不可见的“把手”：压住它拖动即可像原生窗口一样调整大小。
    放在主窗口外层 10px 留白内，不遮挡任何业务控件。
    """

    def __init__(self, parent: QWidget, edges: Qt.Edge,
                 shape: Qt.CursorShape):
        super().__init__(parent)
        self.setObjectName("WindowEdgeHandle")
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setCursor(shape)
        self._edges = edges
        self._active = False
        self._geo0 = QRect()
        self._pos0 = QPoint()

    def mousePressEvent(self, event):
        win = self.window()
        if (event.button() != Qt.MouseButton.LeftButton
                or win.isMaximized() or win.isFullScreen()):
            event.ignore()
            return
        self._active = True
        self._geo0 = win.geometry()
        self._pos0 = event.globalPosition().toPoint()
        event.accept()

    def mouseMoveEvent(self, event):
        if not self._active:
            event.ignore()
            return
        delta = event.globalPosition().toPoint() - self._pos0
        geo = QRect(self._geo0)
        min_w = max(320, self.window().minimumWidth())
        min_h = max(240, self.window().minimumHeight())
        if self._edges & Qt.Edge.LeftEdge:
            geo.setLeft(min(geo.right() - min_w + 1, geo.left() + delta.x()))
        if self._edges & Qt.Edge.RightEdge:
            geo.setRight(max(geo.left() + min_w - 1, geo.right() + delta.x()))
        if self._edges & Qt.Edge.TopEdge:
            geo.setTop(min(geo.bottom() - min_h + 1, geo.top() + delta.y()))
        if self._edges & Qt.Edge.BottomEdge:
            geo.setBottom(max(geo.top() + min_h - 1, geo.bottom() + delta.y()))
        self.window().setGeometry(geo)
        event.accept()

    def mouseReleaseEvent(self, event):
        self._active = False
        event.accept()


# ---------------------------------------------------------------------------
# 底部常驻日志坞
# ---------------------------------------------------------------------------
class BottomDock(QFrame):
    """软件底部常驻区域：发送进度 + 全局日志（不随页面切换隐藏）。"""

    _GLOW_KEYS = ((0.0, 0.0), (0.16, 1.0), (0.42, 0.3), (0.68, 0.9), (1.0, 0.0))

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("BottomDock")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, False)
        self.setMinimumHeight(176)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self._glow = 0.0
        self._glow_anim = QVariantAnimation(self)
        self._glow_anim.setDuration(1500)
        self._glow_anim.setEasingCurve(QEasingCurve.Type.InOutSine)
        for at, value in self._GLOW_KEYS:
            self._glow_anim.setKeyValueAt(at, value)
        self._glow_anim.valueChanged.connect(self._on_glow)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 8, 14, 12)
        layout.setSpacing(6)

        # ---- 进度行 ----
        progress_row = QHBoxLayout()
        progress_row.setSpacing(8)
        self.progress_title = QLabel("📈 发送进度", self)
        self.progress_title.setObjectName("DockTitle")
        progress_row.addWidget(self.progress_title)

        self.progress_bar = GlassProgressBar(self)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_row.addWidget(self.progress_bar, 1)

        self.progress_label = QLabel("等待发送…", self)
        self.progress_label.setObjectName("DockStatus")
        # 不预留固定宽度：标签按文字自适应，进度条才能一直伸展到文字开头
        self.progress_label.setSizePolicy(
            QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Preferred)
        self.progress_label.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        progress_row.addWidget(self.progress_label)
        layout.addLayout(progress_row)

        # ---- 日志标题行 ----
        log_row = QHBoxLayout()
        log_row.setSpacing(8)
        self.log_title = QLabel("📝 运行日志（手动 / 定时 / 轮询）", self)
        self.log_title.setObjectName("DockTitle")
        log_row.addWidget(self.log_title)
        log_row.addStretch(1)

        self.follow_check = QCheckBox("自动跟随", self)
        self.follow_check.setChecked(True)
        self.follow_check.setToolTip(
            "勾选后新日志自动滚动到底部；取消勾选可停在当前位置回看旧日志")
        log_row.addWidget(self.follow_check)

        self.clear_btn = QPushButton("🧹 清空日志", self)
        self.clear_btn.setObjectName("DockBtn")
        self.clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_btn.clicked.connect(lambda: self.log_text.clear_log())
        log_row.addWidget(self.clear_btn)
        layout.addLayout(log_row)

        # ---- 日志本体 ----
        self.log_text = GlobalLogView(self)
        layout.addWidget(self.log_text, 1)

    # ---- 绘制 ----
    def _get_glow(self) -> float:
        return self._glow

    def _set_glow(self, value: float) -> None:
        self._glow = float(value)
        self.update()

    glow = pyqtProperty(float, fget=_get_glow, fset=_set_glow)

    def _on_glow(self, value) -> None:
        self._glow = float(value)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        paint_glass(painter, rect, 14.0, strong=True)
        if self._glow > 0.02:
            c = _colors()
            glow_color = QColor(ui_theme.qcolor(c["accent"]))
            glow_color.setAlpha(int(min(1.0, self._glow) * 220))
            pen = QPen(glow_color)
            pen.setWidthF(2.2)
            painter.setPen(pen)
            painter.drawRoundedRect(rect.adjusted(0.6, 0.6, -0.6, -0.6),
                                    14.0, 14.0)

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == event.Type.PaletteChange:
            self.update()

    def flash_attention(self) -> None:
        """有失败时让底部日志坞“呼吸”高亮两下，提示用户看日志。"""
        self._glow_anim.stop()
        self._glow_anim.start()

    def refresh_theme(self) -> None:
        self.update()
