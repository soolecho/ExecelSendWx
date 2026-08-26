"""和风天气拉取 + 消息模板渲染。

- 全局配置存在 %LOCALAPPDATA%\\ExcelSendWx\\weather.json (api_key/base_url/default_city)
- 任务级 default_city 优先；为空时 fallback 到全局 default_city；都空时
  {{weather}} 这类不带参占位符保留原样不渲染，避免误发占位符
- 内存缓存 (city, base_url, api_key) -> data，10 分钟过期，避免短时间反复请求
- 所有日志走 logging 模块，便于 app.log 排查
"""
from __future__ import annotations

import json
import logging
import os
import re
import tempfile
import threading
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional


logger = logging.getLogger(__name__)


# ----------------------------- 配置持久化 -----------------------------

def _weather_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base = Path(local_app_data) / "ExcelSendWx"
    else:
        base = Path(tempfile.gettempdir()) / "ExcelSendWx"
    base.mkdir(parents=True, exist_ok=True)
    return base


WEATHER_CONFIG_FILE = _weather_dir() / "weather.json"

# 和风免费版默认域名；商业版可改成 https://api.qweather.com
DEFAULT_BASE_URL = "https://devapi.qweather.com"


def load_weather_config() -> Dict:
    """读取全局天气配置。返回 dict，缺失字段填默认。"""
    default = {
        "api_key": "",
        "base_url": DEFAULT_BASE_URL,
        "default_city": "",
    }
    try:
        if WEATHER_CONFIG_FILE.exists():
            with WEATHER_CONFIG_FILE.open("r", encoding="utf-8") as fp:
                raw = json.load(fp)
            if isinstance(raw, dict):
                for k, v in default.items():
                    if k not in raw or not isinstance(raw.get(k), str):
                        raw[k] = v
                return raw
    except Exception as exc:
        logger.warning("读取天气配置失败: %s", exc)
    return default


def save_weather_config(api_key: str = "", base_url: str = "",
                       default_city: str = "") -> Dict:
    """原子写入全局天气配置。任一参数留空时保留原值。"""
    current = load_weather_config()
    if api_key:
        current["api_key"] = api_key.strip()
    if base_url:
        current["base_url"] = base_url.strip() or DEFAULT_BASE_URL
    if default_city is not None:
        current["default_city"] = default_city.strip()
    current["updated_at"] = datetime.now().isoformat(timespec="seconds")
    data = dict(current)
    path = WEATHER_CONFIG_FILE
    fd, tmp = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp",
                               dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fp:
            json.dump(data, fp, ensure_ascii=False, indent=2)
            fp.write("\n")
        os.replace(tmp, path)
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    logger.info("天气配置已保存: base_url=%s default_city=%s key_len=%d",
                current.get("base_url"), current.get("default_city"),
                len(current.get("api_key", "")))
    return current


# ----------------------------- 天气拉取 -----------------------------

_CACHE_LOCK = threading.Lock()
_CACHE: Dict[str, tuple] = {}  # key: city|base_url|api_key -> (ts, data)
_CACHE_TTL = 600.0  # 10 分钟


def _cache_key(city: str, base_url: str, api_key: str) -> str:
    return f"{city}|{base_url}|{api_key[:4]}"


def _http_get_json(url: str, timeout: float = 8.0) -> Dict:
    req = urllib.request.Request(url, headers={
        "User-Agent": "ExcelSendWx/1.0 (Python urllib)",
        "Accept": "application/json",
    })
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        charset = resp.headers.get_content_charset() or "utf-8"
        raw = resp.read().decode(charset, errors="replace")
    return json.loads(raw)


def _geo_base_for(base_url: str) -> str:
    """根据用户填的 base_url 决定 GeoAPI 域名。

    - 新版专属域名 (*.qweatherapi.com, 如 xxx.re.qweatherapi.com) →
      所有 API（含 GeoAPI）都走这个域名，和风 2024+ 推荐
    - 老版共享域名 (devapi.qweather.com / api.qweather.com) →
      GeoAPI 必须走 geoapi.qweather.com 子域名
    """
    host = (urllib.parse.urlparse(base_url).hostname or "").lower()
    if host.endswith(".qweatherapi.com"):
        return base_url.rstrip("/")
    return "https://geoapi.qweather.com"


def lookup_city_id(city: str, base_url: str, api_key: str,
                   log_fn=None) -> Optional[str]:
    """城市名 -> 和风 LocationID。失败返回 None。"""
    def _log(msg):
        try:
            logger.info(msg)
            if log_fn:
                log_fn(msg)
        except Exception:
            pass

    if not (city and base_url and api_key):
        _log(f"城市查询跳过：city={city!r} base_url={base_url!r} key_len={len(api_key)}")
        return None
    geo_base = _geo_base_for(base_url)
    url = (f"{geo_base}/v2/city/lookup?"
           f"location={urllib.parse.quote(city)}&key={api_key}")
    _log(f"GET {geo_base}/v2/city/lookup?location={city}&key=***")
    try:
        data = _http_get_json(url)
    except Exception as exc:
        _log(f"❌ 城市查询请求失败 {city}: {exc}")
        return None
    if data.get("code") != "200":
        _log(f"❌ 城市查询返回 code={data.get('code')} msg={data.get('message','')} (city={city})")
        return None
    locations = data.get("location") or []
    if not locations:
        return None
    return str(locations[0].get("id") or "") or None


def fetch_weather(city: str, base_url: str = "",
                  api_key: str = "", log_fn=None) -> Optional[Dict]:
    """
    拉取实时天气。返回 dict:
      {
        "city": "北京",
        "text": "晴",
        "temp": "25",
        "wind": "北风 3级",
        "humidity": "40",
        "update_time": "2026-08-26 08:00"
      }
    失败返回 None。
    """
    def _log(msg):
        try:
            logger.info(msg)
            if log_fn:
                log_fn(msg)
        except Exception:
            pass

    if not city:
        return None
    base_url = (base_url or DEFAULT_BASE_URL).rstrip("/")
    if not api_key:
        _log("❌ 和风 API key 为空，跳过天气拉取")
        return None

    cache_k = _cache_key(city, base_url, api_key)
    now = datetime.now().timestamp()
    with _CACHE_LOCK:
        hit = _CACHE.get(cache_k)
        if hit and (now - hit[0]) < _CACHE_TTL:
            _log(f"命中缓存（{int(_CACHE_TTL - (now - hit[0]))}s 后过期）：{city}")
            return hit[1]

    location_id = lookup_city_id(city, base_url, api_key, log_fn=log_fn)
    if not location_id:
        _log(f"❌ 未找到城市 LocationID：{city}")
        return None
    _log(f"LocationID={location_id}，开始拉取实时天气...")
    url = (f"{base_url}/v7/weather/now?"
           f"location={location_id}&key={api_key}")
    try:
        data = _http_get_json(url)
    except Exception as exc:
        _log(f"❌ 实时天气请求失败 {city}: {exc}")
        return None
    if data.get("code") != "200":
        _log(f"❌ 实时天气返回 code={data.get('code')} msg={data.get('message','')}")
        return None
    now_obj = data.get("now") or {}
    result = {
        "city": city,
        "text": str(now_obj.get("text") or ""),
        "temp": str(now_obj.get("temp") or ""),
        "wind_dir": str(now_obj.get("windDir") or ""),
        "wind_scale": str(now_obj.get("windScale") or ""),
        "wind_speed": str(now_obj.get("windSpeed") or ""),
        "humidity": str(now_obj.get("humidity") or ""),
        "feels_like": str(now_obj.get("feelsLike") or ""),
        "update_time": str(now_obj.get("obsTime") or ""),
    }
    result["wind"] = _format_wind(result["wind_dir"], result["wind_scale"])
    result["summary"] = _format_summary(result)
    with _CACHE_LOCK:
        _CACHE[cache_k] = (now, result)
    return result


def _format_wind(direction: str, scale: str) -> str:
    if not direction:
        return ""
    if scale:
        return f"{direction} {scale}级"
    return direction


def _format_summary(data: Dict) -> str:
    parts = []
    if data.get("text"):
        parts.append(data["text"])
    if data.get("temp"):
        parts.append(f"{data['temp']}°C")
    return " ".join(parts)


def test_connection(api_key: str, base_url: str,
                    city: str = "北京", log_fn=None) -> tuple[bool, str]:
    """UI 测试连接按钮用。返回 (ok, message)。"""
    if not api_key:
        if log_fn: log_fn("❌ API key 为空")
        return False, "API key 为空"
    base_url = (base_url or DEFAULT_BASE_URL).rstrip("/")
    if log_fn: log_fn(f"=== 测试连接开始 city={city} base_url={base_url} ===")
    data = fetch_weather(city, base_url, api_key, log_fn=log_fn)
    if data:
        msg = f"✅ 连接成功：{data.get('city','')} {data.get('summary','')}"
        if log_fn: log_fn(msg)
        return True, msg
    msg = f"❌ 连接失败：请检查 key/域名/网络（测试城市：{city}）"
    if log_fn: log_fn(msg)
    return False, msg


# ----------------------------- 模板渲染 -----------------------------

# 匹配 {{weather}} / {{weather:北京}} / {{temp}} / {{temp:北京}} / {{wind}} / {{date}} 等
_PATTERN = re.compile(r"\{\{\s*([a-zA-Z_]+)(?:\s*:\s*([^}]+?))?\s*\}\}")


def _resolve_city(arg: Optional[str], task_default_city: str,
                  global_default_city: str) -> str:
    city = (arg or "").strip()
    if city:
        return city
    if task_default_city:
        return task_default_city.strip()
    if global_default_city:
        return global_default_city.strip()
    return ""


def render_message(text: str, task_default_city: str = "",
                   api_key: str = "", base_url: str = "",
                   global_default_city: str = "") -> str:
    """把 {{weather}}/{{temp}}/{{wind}}/{{date}}/{{weekday}}/{{news}} 等占位符
    替换成真实值。无城市/网络失败时保留原占位符字符串，避免误发。
    """
    if not text:
        return text
    if "{{" not in text:
        return text

    cache_for_render: Dict[str, Optional[Dict]] = {}

    def get_weather(city: str) -> Optional[Dict]:
        if not city:
            return None
        if city in cache_for_render:
            return cache_for_render[city]
        data = fetch_weather(city, base_url, api_key)
        cache_for_render[city] = data
        return data

    def replacer(match: re.Match) -> str:
        key = match.group(1).lower()
        arg = match.group(2)
        if key in ("weather", "temp", "wind", "humidity", "feels_like",
                   "summary", "city"):
            if not api_key:
                logger.info("天气占位符 %s 但 API key 为空，保留原文", key)
                return match.group(0)
            city = _resolve_city(arg, task_default_city, global_default_city)
            if not city:
                return match.group(0)
            data = get_weather(city)
            if not data:
                return match.group(0)
            if key == "weather":
                return data.get("summary") or data.get("text") or ""
            if key == "city":
                return data.get("city") or city
            return data.get({
                "temp": "temp",
                "wind": "wind",
                "humidity": "humidity",
                "feels_like": "feels_like",
                "summary": "summary",
            }.get(key, "")) or ""
        if key in ("news", "hotsearch", "hot", "热搜"):
            try:
                from modules import news_fetcher
                platform, limit = news_fetcher.parse_news_arg(arg)
                items = news_fetcher.fetch_top(platform=platform, limit=limit)
                if not items:
                    return match.group(0)
                return news_fetcher.format_items(items, platform=platform)
            except Exception as exc:
                logger.warning("新闻占位符渲染失败: %s", exc)
                return match.group(0)
        if key == "date":
            return datetime.now().strftime("%Y-%m-%d")
        if key == "weekday":
            from datetime import date as _date
            names = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
            return names[_date.today().isoweekday() - 1]
        if key == "time":
            return datetime.now().strftime("%H:%M")
        # 未知占位符保留
        return match.group(0)

    return _PATTERN.sub(replacer, text)


def find_placeholders(text: str) -> list:
    """返回消息中所有占位符名称（去重，按出现顺序）。UI 提示用。"""
    seen = []
    for m in _PATTERN.finditer(text or ""):
        name = m.group(1).lower()
        if name not in seen:
            seen.append(name)
    return seen
