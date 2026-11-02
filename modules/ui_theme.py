# -*- coding: utf-8 -*-
"""日夜主题与现代化 QSS。

- 主题选择持久化到 %LOCALAPPDATA%\\ExcelSendWx\\ui_state.json；
- 「玻璃质感」三档（经典磨砂 / 通透玻璃 / 实心高对比）与明暗主题独立组合，
  只缩放玻璃类色值的 alpha，同样持久化在 ui_state.json；
- 支持「跟随系统」：读取 Windows 深浅色设置，系统切换时自动重刷主题；
- apply_theme 同时切换 QPalette 与全局 QSS，既有页面里写死的彩色按钮样式
  仍优先生效（setStyleSheet 的局部规则覆盖全局），不会破坏原有功能配色；
- 颜色集合同时供左侧导航自绘使用（colors_for），避免导航硬编码色值。
"""
import json
import os
import re
from typing import Dict, Tuple

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QFont, QPalette
from PyQt6.QtWidgets import QApplication

LIGHT = "light"
DARK = "dark"
AUTO = "auto"
_VALID = (LIGHT, DARK, AUTO)

# 玻璃质感：同一套明暗配色下，只调整玻璃底的透明度 / 高光 / 描边强度
GLASS_STANDARD = "standard"
GLASS_CLEAR = "clear"
GLASS_SOLID = "solid"
GLASS_STYLES = (GLASS_STANDARD, GLASS_CLEAR, GLASS_SOLID)


def _state_path() -> str:
    local = os.environ.get("LOCALAPPDATA")
    if not local:
        local = os.path.join(os.environ.get("USERPROFILE", os.getcwd()), "AppData", "Local")
    return os.path.join(local, "ExcelSendWx", "ui_state.json")


def _load_state() -> Dict:
    try:
        with open(_state_path(), "r", encoding="utf-8") as fp:
            data = json.load(fp)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_state(**updates) -> None:
    """把若干键合并写回 ui_state.json（保留其它已有键）。"""
    state = _load_state()
    state.update(updates)
    path = _state_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as fp:
            json.dump(state, fp, ensure_ascii=False, indent=2)
    except OSError:
        pass


def load_theme() -> str:
    """读取已保存的主题；未设置或非法值回退「跟随系统」。"""
    mode = str(_load_state().get("theme", AUTO)).lower()
    return mode if mode in _VALID else AUTO


def current_mode() -> str:
    """当前主题选择（light / dark / auto），未应用过主题时读存档。"""
    app = QApplication.instance()
    mode = app.property("ui_theme_mode") if app is not None else None
    return mode if mode in _VALID else load_theme()


def system_theme() -> str:
    """读取当前系统深浅色：优先 Qt 的 colorScheme，回退 Windows 注册表。"""
    app = QApplication.instance()
    hints = None
    try:
        if app is not None and hasattr(app, "styleHints"):
            hints = app.styleHints()
        scheme = hints.colorScheme() if (hints is not None and
                                         hasattr(hints, "colorScheme")) else None
        if scheme == Qt.ColorScheme.Dark:
            return DARK
        if scheme == Qt.ColorScheme.Light:
            return LIGHT
    except Exception:
        pass
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        ) as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return LIGHT if int(value) else DARK
    except (OSError, ValueError, ImportError):
        return LIGHT


def resolve_theme(mode: str) -> str:
    """把主题设置解析成实际生效的 light / dark（auto 时跟随系统）。"""
    mode = str(mode or "").lower()
    if mode == AUTO:
        return system_theme()
    return mode if mode in (LIGHT, DARK) else LIGHT


def save_theme(mode: str) -> None:
    mode = mode if mode in _VALID else LIGHT
    _save_state(theme=mode)


def load_glass_style() -> str:
    """读取已保存的玻璃质感；未设置或非法值回退「经典磨砂」。"""
    style = str(_load_state().get("glass_style", GLASS_STANDARD)).lower()
    return style if style in GLASS_STYLES else GLASS_STANDARD


def current_glass_style() -> str:
    """当前玻璃质感（已应用过则读 app 属性，否则读存档）。"""
    app = QApplication.instance()
    style = app.property("ui_glass_style") if app is not None else None
    return style if style in GLASS_STYLES else load_glass_style()


def save_glass_style(style: str) -> None:
    style = style if style in GLASS_STYLES else GLASS_STANDARD
    _save_state(glass_style=style)


_LIGHT_COLORS = {
    "window": "#e9eff9",
    "surface": "#ffffff",
    "card": "#ffffff",
    "card_alt": "#f7f9fc",
    "text": "#1f2532",
    "muted": "#6f7a8c",
    "border": "#dbe2ec",
    "accent": "#2f78c4",
    "accent_soft": "#e8f1fd",
    "hover": "#eef4fd",
    "nav_bg": "#f7f9fd",
    "nav_active": "#e4effc",
    "danger": "#d93025",
    "ok": "#1e8e3e",
    "warn": "#e67e22",
    "log_bg": "#fbfcfe",
    # 应用背景的斜向渐变末端色（给玻璃面板留出景深）
    "bg_tint": "#dbe7f8",
    # 玻璃拟态：半透明底色 / 高光 / 描边 / 进度条轨道
    "glass": "rgba(255,255,255,224)",
    "glass_strong": "rgba(255,255,255,246)",
    "glass_sheen": "rgba(255,255,255,150)",
    "glass_border": "rgba(122,152,196,84)",
    "glass_track": "rgba(118,144,182,52)",
}

_DARK_COLORS = {
    "window": "#13161b",
    "surface": "#1b1f26",
    "card": "#21252d",
    "card_alt": "#1d2129",
    "text": "#e6eaf2",
    "muted": "#8d97a8",
    "border": "#2d333d",
    "accent": "#5aa6ec",
    "accent_soft": "#1e2c3e",
    "hover": "#252b35",
    "nav_bg": "#171b21",
    "nav_active": "#22303f",
    "danger": "#ff6b5e",
    "ok": "#59d27a",
    "warn": "#f0a34a",
    "log_bg": "#12151a",
    "bg_tint": "#0f1620",
    "glass": "rgba(38,45,56,214)",
    "glass_strong": "rgba(46,54,66,236)",
    "glass_sheen": "rgba(255,255,255,46)",
    "glass_border": "rgba(255,255,255,46)",
    "glass_track": "rgba(255,255,255,30)",
}


_RGBA_RE = re.compile(r"rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)")

# 玻璃质感 → 各 glass token 的 alpha 缩放系数（1.0 = 经典磨砂，不缩放）
_GLASS_ALPHA_FACTORS = {
    GLASS_STANDARD: {},
    # 通透玻璃：底更透、高光/描边更亮，背景渐变能透出来
    GLASS_CLEAR: {
        "glass": 0.72,
        "glass_strong": 0.84,
        "glass_sheen": 1.25,
        "glass_border": 1.20,
        "glass_track": 1.25,
    },
    # 实心高对比：底更实、高光收敛、描边更清晰，小屏/低对比屏更易读
    GLASS_SOLID: {
        "glass": 1.12,
        "glass_strong": 1.06,
        "glass_sheen": 0.72,
        "glass_border": 1.15,
        "glass_track": 1.20,
    },
}

_COLOR_CACHE: Dict[Tuple[str, str], Dict[str, str]] = {}


def _scale_alpha(value: str, factor: float) -> str:
    """按系数缩放 rgba(...) 的 alpha（0-255，自动钳制）；非 rgba 原样返回。"""
    match = _RGBA_RE.match(str(value).strip())
    if not match:
        return value
    r, g, b, a = (int(x) for x in match.groups())
    alpha = max(0, min(255, int(round(a * float(factor)))))
    return f"rgba({r},{g},{b},{alpha})"


def _themed_colors(theme: str, style: str) -> Dict[str, str]:
    """明暗基础色 + 玻璃质感缩放后的完整色集（带缓存，返回副本）。"""
    key = (theme, style)
    cached = _COLOR_CACHE.get(key)
    if cached is None:
        cached = dict(_DARK_COLORS if theme == DARK else _LIGHT_COLORS)
        for token, factor in _GLASS_ALPHA_FACTORS.get(style, {}).items():
            if token in cached:
                cached[token] = _scale_alpha(cached[token], factor)
        _COLOR_CACHE[key] = cached
    return dict(cached)


def colors_for(mode: str, style: str = "") -> Dict[str, str]:
    """返回实际生效的色集（明暗 + 当前玻璃质感）。"""
    theme = DARK if (mode or "").lower() == DARK else LIGHT
    glass = str(style or current_glass_style()).lower()
    if glass not in GLASS_STYLES:
        glass = GLASS_STANDARD
    return _themed_colors(theme, glass)


def apply_glass_style(app: QApplication, style: str, persist: bool = True) -> str:
    """切换玻璃质感（standard / clear / solid）并刷新全局 QSS。"""
    style = str(style or "").lower()
    if style not in GLASS_STYLES:
        style = GLASS_STANDARD
    if persist:
        save_glass_style(style)
    if app is not None:
        # 先写属性，_build_qss → colors_for 才能取到新质感
        app.setProperty("ui_glass_style", style)
        theme = app.property("ui_theme")
        if theme not in (LIGHT, DARK):
            theme = resolve_theme(current_mode())
        app.setStyleSheet(_build_qss(theme))
    return style


def qcolor(value: str) -> QColor:
    """把颜色字符串转成 QColor（支持 #RRGGBB 与 rgba(r,g,b,a) 两种写法）。"""
    match = _RGBA_RE.match(str(value).strip())
    if match:
        r, g, b, a = (int(x) for x in match.groups())
        return QColor(r, g, b, a)
    return QColor(value)


def _build_palette(mode: str) -> QPalette:
    c = colors_for(mode)
    pal = QPalette()
    roles = {
        QPalette.ColorRole.Window: c["window"],
        QPalette.ColorRole.WindowText: c["text"],
        QPalette.ColorRole.Base: c["surface"],
        QPalette.ColorRole.AlternateBase: c["card_alt"],
        QPalette.ColorRole.ToolTipBase: c["card"],
        QPalette.ColorRole.ToolTipText: c["text"],
        QPalette.ColorRole.Text: c["text"],
        QPalette.ColorRole.Button: c["card"],
        QPalette.ColorRole.ButtonText: c["text"],
        QPalette.ColorRole.BrightText: "#ffffff",
        QPalette.ColorRole.Highlight: c["accent"],
        QPalette.ColorRole.HighlightedText: "#ffffff",
        QPalette.ColorRole.Link: c["accent"],
        QPalette.ColorRole.PlaceholderText: c["muted"],
        QPalette.ColorRole.Light: c["card_alt"],
        QPalette.ColorRole.Midlight: c["border"],
        QPalette.ColorRole.Mid: c["border"],
        QPalette.ColorRole.Dark: c["border"],
        QPalette.ColorRole.Shadow: "#00000033",
    }
    for role, value in roles.items():
        pal.setColor(role, QColor(value))
    try:
        pal.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor(c["muted"]))
        pal.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor(c["muted"]))
    except Exception:
        pass
    return pal


def _build_qss(mode: str) -> str:
    c = colors_for(mode)
    return f"""
QWidget {{
    font-family: "Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", sans-serif;
    font-size: 12px;
    color: {c['text']};
}}
QMainWindow, QDialog {{ background: {c['window']}; }}
QWidget#AppRoot {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 {c['window']}, stop:0.45 {c['surface']}, stop:1 {c['bg_tint']});
}}
QScrollArea {{ background: transparent; border: none; }}
QScrollArea#PageScroll, QScrollArea#PageScroll > QWidget > QWidget,
QScrollArea#NavScroll, QScrollArea#NavScroll > QWidget > QWidget {{
    background: transparent;
    border: none;
}}
QFrame#BottomDock, QFrame#NavPanel, QFrame#GlassHeader {{
    background: transparent;
    border: none;
}}
QLabel#HeaderIcon {{ background: transparent; font-size: 18px; }}
QLabel#HeaderTitle {{
    color: {c['text']}; font-size: 15px; font-weight: bold; background: transparent;
}}
QLabel#HeaderPill {{
    background: {c['accent_soft']}; color: {c['accent']};
    border-radius: 9px; padding: 1px 8px; font-size: 10px;
}}
QLabel#HeaderStatus {{
    color: {c['muted']}; font-size: 11px; padding-left: 6px;
    background: transparent;
}}
QPushButton#HeaderBtn {{
    background: {c['glass']};
    color: {c['text']};
    border: 1px solid {c['glass_border']};
    border-radius: 10px;
    padding: 4px 12px;
    font-size: 11px;
}}
QPushButton#HeaderBtn:hover {{
    background: {c['glass_strong']}; border-color: {c['accent']};
}}
QPushButton#HeaderBtn:pressed {{ background: {c['accent_soft']}; }}
/* 无边框标题栏上的窗口按钮：最小化 / 最大化 / 关闭（悬停才显示底色） */
QPushButton#WinBtn, QPushButton#WinBtnClose {{
    background: transparent;
    color: {c['muted']};
    border: 1px solid transparent;
    border-radius: 7px;
    padding: 0px;
    margin: 0px;
    font-size: 12px;
}}
QPushButton#WinBtn:hover {{
    background: {c['hover']}; color: {c['text']};
}}
QPushButton#WinBtn:pressed {{ background: {c['accent_soft']}; }}
QPushButton#WinBtnClose:hover {{
    background: {c['danger']}; color: #ffffff; border-color: {c['danger']};
}}
QPushButton#WinBtnClose:pressed {{
    background: {c['danger']}; color: #ffffff; border-color: {c['danger']};
}}
/* 待机页内置浏览器（嵌在待机右侧的滑入面板）：底 + 工具栏按钮 + 只读地址栏 */
QWidget#StandbyBrowser {{
    background: {c['surface']};
    border: 1px solid {c['glass_border']};
    border-radius: 14px;
}}
QPushButton#BrowserBackBtn {{
    padding: 3px 12px; font-size: 12px; border-radius: 8px;
    color: #ffffff; background: {c['accent']};
    border: 1px solid {c['accent']}; font-weight: bold;
}}
QPushButton#BrowserBackBtn:hover {{
    background: {c['accent']}; color: #ffffff;
}}
QPushButton#BrowserBtn {{
    padding: 3px 10px; font-size: 12px; border-radius: 8px;
    color: {c['text']}; background: {c['glass']};
    border: 1px solid {c['glass_border']};
}}
QPushButton#BrowserBtn:hover {{
    background: {c['glass_strong']}; border-color: {c['accent']};
}}
QPushButton#BrowserBtn:disabled {{
    color: {c['muted']}; background: {c['glass_track']};
}}
QLineEdit#BrowserAddress {{
    background: {c['card_alt']}; border: 1px solid {c['border']};
    border-radius: 8px; padding: 3px 8px;
    color: {c['muted']}; font-size: 11px;
}}
QPushButton#NavRailBtn, QPushButton#DockBtn {{
    background: {c['glass']};
    color: {c['muted']};
    border: 1px solid {c['glass_border']};
    border-radius: 9px;
    padding: 3px 10px;
    font-size: 11px;
}}
QPushButton#NavRailBtn:hover, QPushButton#DockBtn:hover {{
    background: {c['glass_strong']}; color: {c['text']}; border-color: {c['accent']};
}}
QLabel#DockTitle {{ color: {c['muted']}; font-size: 11px; font-weight: bold; }}
QLabel#DockStatus {{ color: {c['muted']}; font-size: 11px; }}
QLabel#NavBrand {{ color: {c['text']}; font-size: 13px; font-weight: bold; }}
QLabel#NavBrandSub {{ color: {c['muted']}; font-size: 10px; }}
/* 待机画面：大时钟 + 日期 + 问候语（全部透明底，靠自绘玻璃面板做背景） */
QLabel#StandbyIcon {{ background: transparent; font-size: 30px; }}
QLabel#StandbyClock {{
    background: transparent; color: {c['text']};
    font-size: 44px; font-weight: 600; letter-spacing: 2px;
}}
QLabel#StandbyDate {{
    background: transparent; color: {c['muted']}; font-size: 13px;
}}
QLabel#StandbyWeather {{
    background: transparent; color: {c['text'] };
    font-size: 14px; font-weight: 600; padding-top: 2px;
}}
QLabel#StandbyWeatherSub {{
    background: transparent; color: {c['muted']}; font-size: 11px;
}}
QLabel#StandbySection {{
    background: transparent; color: {c['accent']};
    font-size: 11px; font-weight: bold; padding-top: 6px;
}}
QLabel#StandbyNews {{
    background: transparent; color: {c['text']}; font-size: 11px;
}}
QPushButton#StandbyAction {{
    padding: 4px 10px; font-size: 11px; border-radius: 8px;
    color: {c['text']}; background: {c['glass']};
    border: 1px solid {c['glass_border']};
}}
QPushButton#StandbyAction:hover {{
    background: {c['glass_strong']}; border-color: {c['accent']};
}}
QPushButton#StandbyAction:disabled {{
    color: {c['muted']}; background: {c['glass_track']};
}}
QScrollArea#StandbyScroll {{ background: transparent; border: none; }}
QWidget#StandbyContent {{ background: transparent; }}
QLabel#StandbyTip {{
    background: transparent; color: {c['muted']}; font-size: 11px;
    padding-top: 4px;
}}
/* 展开面板内的小节小标题（便于在单列长页面里识别内容块） */
QLabel#SectionHeader {{
    color: {c['muted']}; font-size: 11px; font-weight: bold;
    padding: 1px 0 3px 0;
    border-bottom: 1px solid {c['border']};
}}
/* 语义色调标签：替代页面里写死的 #666/#888（深色主题下也能清晰可读） */
QLabel#MutedNote {{ color: {c['muted']}; }}
QLabel#MutedHint {{ color: {c['muted']}; font-size: 11px; }}
QLabel#StatusOk {{ color: {c['ok']}; }}
QLabel#StatusOkSmall {{ color: {c['ok']}; font-size: 11px; }}
QLabel#StatusErr {{ color: {c['danger']}; }}
QLabel#StatusErrSmall {{ color: {c['danger']}; font-size: 11px; }}
QLabel#StatusWarn {{ color: {c['warn']}; }}
QLabel#StatusWarnSmall {{ color: {c['warn']}; font-size: 11px; }}
QPushButton {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 {c['glass_strong']}, stop:1 {c['card_alt']});
    color: {c['text']};
    border: 1px solid {c['glass_border']};
    border-radius: 9px;
    padding: 5px 10px;
}}
QPushButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 {c['glass_strong']}, stop:1 {c['hover']});
    border-color: {c['accent']};
}}
QPushButton:pressed {{ background: {c['accent_soft']}; border-color: {c['accent']}; }}
QPushButton:disabled {{ color: {c['muted']}; background: {c['card_alt']}; border-color: {c['border']}; }}
QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QSpinBox, QDoubleSpinBox,
QTimeEdit, QDateEdit {{
    background: {c['card']};
    border: 1px solid {c['border']};
    border-radius: 8px;
    padding: 4px 6px;
    selection-background-color: {c['accent']};
    selection-color: #ffffff;
}}
QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus, QComboBox:focus,
QSpinBox:focus, QDoubleSpinBox:focus {{ border-color: {c['accent']}; }}
QComboBox QAbstractItemView {{
    background: {c['card']};
    border: 1px solid {c['border']};
    selection-background-color: {c['accent_soft']};
    selection-color: {c['text']};
    outline: none;
}}
QGroupBox {{
    background: {c['card']};
    border: 1px solid {c['border']};
    border-radius: 10px;
    margin-top: 12px;
    padding: 10px 8px 8px 8px;
}}
QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 4px; color: {c['muted']}; }}
QCheckBox, QRadioButton {{ spacing: 6px; }}
QCheckBox::indicator, QRadioButton::indicator {{ width: 15px; height: 15px; }}
QTableWidget, QListWidget, QTreeWidget {{
    background: {c['card']};
    alternate-background-color: {c['card_alt']};
    border: 1px solid {c['border']};
    border-radius: 8px;
    gridline-color: {c['border']};
}}
QTableWidget::item:selected, QListWidget::item:selected {{
    background: {c['accent_soft']};
    color: {c['text']};
}}
QListWidget::item {{ padding: 3px 4px; border-radius: 6px; }}
QHeaderView::section {{
    background: {c['card_alt']};
    color: {c['muted']};
    border: none;
    border-bottom: 1px solid {c['border']};
    padding: 4px 6px;
}}
QProgressBar {{
    background: {c['glass_track']};
    border: 1px solid {c['glass_border']};
    border-radius: 9px;
    height: 16px;
    text-align: center;
    color: {c['text']};
}}
QProgressBar::chunk {{
    border-radius: 8px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 {c['accent']}, stop:0.7 {c['accent']}, stop:1 {c['ok']});
}}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: {c['border']}; min-height: 28px; border-radius: 5px; }}
QScrollBar::handle:vertical:hover {{ background: {c['accent']}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; }}
QScrollBar::handle:horizontal {{ background: {c['border']}; min-width: 28px; border-radius: 5px; }}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
QSplitter::handle {{ background: {c['window']}; }}
QSplitter::handle:hover {{ background: {c['accent_soft']}; }}
QTextEdit#GlobalLog {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 {c['log_bg']}, stop:1 {c['card_alt']});
    border: 1px solid {c['glass_border']};
    border-radius: 12px;
    padding: 7px 10px;
    font-family: "Cascadia Mono", "Consolas", "Microsoft YaHei UI", monospace;
    font-size: 12px;
    color: {c['text']};
}}
QTextEdit#GlobalLog QScrollBar:vertical {{ background: transparent; width: 9px; }}
QToolTip {{
    background: {c['card']};
    color: {c['text']};
    border: 1px solid {c['accent']};
    padding: 5px 7px;
}}
QMenu {{ background: {c['card']}; border: 1px solid {c['border']}; border-radius: 8px; }}
QMenu::item {{ padding: 5px 18px 5px 14px; border-radius: 6px; }}
QMenu::item:selected {{ background: {c['accent_soft']}; }}
"""


_TONE_NAMES = {
    ("muted", False): "MutedNote",
    ("muted", True): "MutedHint",
    ("ok", False): "StatusOk",
    ("ok", True): "StatusOkSmall",
    ("warn", False): "StatusWarn",
    ("warn", True): "StatusWarnSmall",
    ("err", False): "StatusErr",
    ("err", True): "StatusErrSmall",
}


def apply_tone(widget, tone: str = "muted", small: bool = False) -> None:
    """给标签套用语义色调（随主题自动更新，替代页面里写死的 #666/#888）。"""
    if widget is None:
        return
    name = _TONE_NAMES.get((str(tone), bool(small)))
    if not name or widget.objectName() == name:
        return
    widget.setObjectName(name)
    try:
        style = widget.style()
        style.unpolish(widget)
        style.polish(widget)
    except Exception:
        pass
    widget.update()


def apply_theme(app: QApplication, mode: str) -> None:
    """应用主题（light / dark / auto）：切换调色板与全局 QSS。

    auto 会先解析成系统当前的深浅色；实际生效的主题写在 app 的
    ``ui_theme`` 属性里，原始选择写在 ``ui_theme_mode``，方便界面回显。
    """
    mode = str(mode or "").lower()
    if mode not in _VALID:
        mode = LIGHT
    theme = resolve_theme(mode)
    try:
        app.setStyle("Fusion")
    except Exception:
        pass
    app.setPalette(_build_palette(theme))
    font = QFont("Microsoft YaHei UI", 12)
    # 关闭子像素（ClearType）抗锯齿：深色底上子像素彩边会让小字号中文发糊，
    # 改为灰度抗锯齿后明暗两套主题下文字都清晰锐利。
    font.setStyleStrategy(
        QFont.StyleStrategy.PreferAntialias
        | QFont.StyleStrategy.NoSubpixelAntialias
    )
    app.setFont(font)
    # 玻璃质感与明暗主题独立：先固定属性，_build_qss → colors_for 才会用同一套质感
    app.setProperty("ui_glass_style", current_glass_style())
    app.setStyleSheet(_build_qss(theme))
    app.setProperty("ui_theme", theme)
    app.setProperty("ui_theme_mode", mode)
