"""知乎/头条/B站 热搜拉取 + 消息格式化。

占位符语法 (在 weather_fetcher.render_message 内解析):
  {{news}}              -> 知乎热榜 Top 10 (默认)
  {{news:zhihu}}        -> 知乎热榜 Top 10
  {{news:toutiao}}      -> 头条热搜 Top 10
  {{news:bilibili}}     -> B 站热门视频 Top 10
  {{news:5}}            -> 知乎 Top 5
  {{news:zhihu:5}}      -> 知乎 Top 5
  {{news:toutiao:3}}    -> 头条 Top 3

微博端点都需要 cookie 或返回 4xx/5xx, 已暂时移除。
内存缓存 5 分钟。失败时返回 None, 调用方保留原占位符不误发。
"""
from __future__ import annotations

import json
import logging
import threading
import urllib.parse
import urllib.request
from datetime import datetime
from typing import Dict, List, Optional, Tuple


logger = logging.getLogger(__name__)

_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

_PLATFORMS = {"zhihu", "toutiao", "bilibili"}

_CACHE_LOCK = threading.Lock()
_CACHE: Dict[str, Tuple[float, List[Dict]]] = {}
_CACHE_TTL = 300.0  # 5 分钟


def _cache_key(platform: str) -> str:
    return f"news|{platform}"


def _http_get_json(url: str, timeout: float = 8.0) -> Dict:
    req = urllib.request.Request(url, headers={
        "User-Agent": _UA,
        "Accept": "application/json,text/plain,*/*",
        "Accept-Language": "zh-CN,zh;q=0.9",
    })
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        charset = resp.headers.get_content_charset() or "utf-8"
        raw = resp.read().decode(charset, errors="replace")
    return json.loads(raw)


def _normalize(items: List[Dict], platform: str) -> List[Dict]:
    """统一字段: rank / title / hot / url"""
    out = []
    for i, x in enumerate(items, 1):
        title = x.get("title") or x.get("word") or ""
        if not title:
            continue
        hot = x.get("hot") or x.get("hot_value") or 0
        try:
            hot = int(hot)
        except (TypeError, ValueError):
            hot = 0
        url = x.get("url") or ""
        out.append({
            "rank": i,
            "title": str(title).strip(),
            "hot": hot,
            "url": url,
            "platform": platform,
        })
    return out


def fetch_zhihu(limit: int = 10, log_fn=None) -> List[Dict]:
    """知乎热榜。"""
    url = "https://api.zhihu.com/topstory/hot-lists/total?limit=20"
    try:
        data = _http_get_json(url)
    except Exception as exc:
        if log_fn: log_fn(f"❌ 知乎热榜请求失败: {exc}")
        logger.warning("知乎热榜请求失败: %s", exc)
        return []
    arr = data.get("data") or []
    items = []
    for x in arr:
        target = x.get("target") or {}
        title = target.get("title") or ""
        if not title:
            continue
        hot_text = x.get("detail_text") or ""
        hot_num = 0
        for ch in hot_text:
            if ch.isdigit():
                hot_num = hot_num * 10 + int(ch)
            else:
                break
        items.append({
            "title": title,
            "hot": hot_num * 10000,  # 知乎 detail_text 形如 "123 万热度"
            "url": f"https://www.zhihu.com/question/{target.get('id','')}",
        })
    return _normalize(items[:limit], "zhihu")


def fetch_toutiao(limit: int = 10, log_fn=None) -> List[Dict]:
    """头条热搜。"""
    url = "https://www.toutiao.com/hot-event/hot-board/?origin=toutiao_pc"
    try:
        data = _http_get_json(url)
    except Exception as exc:
        if log_fn: log_fn(f"❌ 头条热搜请求失败: {exc}")
        logger.warning("头条热搜请求失败: %s", exc)
        return []
    arr = data.get("data") or []
    items = []
    for x in arr:
        title = x.get("Title") or ""
        if not title:
            continue
        try:
            hot = int(x.get("HotValue") or 0)
        except (TypeError, ValueError):
            hot = 0
        items.append({
            "title": title,
            "hot": hot,
            "url": x.get("Url") or "",
        })
    return _normalize(items[:limit], "toutiao")


def fetch_bilibili(limit: int = 10, log_fn=None) -> List[Dict]:
    """B 站热门视频榜 (popular 端点, 无需 referer)。"""
    url = "https://api.bilibili.com/x/web-interface/popular?ps=20&pn=1"
    try:
        data = _http_get_json(url)
    except Exception as exc:
        if log_fn: log_fn(f"❌ B站热门请求失败: {exc}")
        logger.warning("B站热门请求失败: %s", exc)
        return []
    arr = (data.get("data") or {}).get("list") or []
    items = []
    for x in arr:
        title = x.get("title") or ""
        if not title:
            continue
        stat = x.get("stat") or {}
        try:
            hot = int(stat.get("view") or 0)
        except (TypeError, ValueError):
            hot = 0
        bvid = x.get("bvid") or ""
        items.append({
            "title": title,
            "hot": hot,
            "url": f"https://www.bilibili.com/video/{bvid}",
        })
    return _normalize(items[:limit], "bilibili")


_FETCHERS = {
    "zhihu": fetch_zhihu,
    "toutiao": fetch_toutiao,
    "bilibili": fetch_bilibili,
}

_PLATFORM_LABEL = {
    "zhihu": "知乎热榜",
    "toutiao": "头条热搜",
    "bilibili": "B站热门",
}


def fetch_top(platform: str = "zhihu", limit: int = 10,
              log_fn=None) -> List[Dict]:
    """统一入口, 带内存缓存。"""
    platform = (platform or "zhihu").strip().lower()
    if platform not in _PLATFORMS:
        if log_fn: log_fn(f"❌ 未知热搜平台: {platform}, 可选: {list(_PLATFORMS)}")
        return []
    if limit < 1:
        limit = 10
    if limit > 50:
        limit = 50

    cache_k = _cache_key(platform)
    now = datetime.now().timestamp()
    with _CACHE_LOCK:
        hit = _CACHE.get(cache_k)
        if hit and (now - hit[0]) < _CACHE_TTL:
            return hit[1][:limit]

    fetcher = _FETCHERS[platform]
    items = fetcher(limit=limit, log_fn=log_fn)
    with _CACHE_LOCK:
        _CACHE[cache_k] = (now, items)
    return items


def format_items(items: List[Dict], platform: str = "weibo") -> str:
    """把热搜 List 渲染成多行文本消息。"""
    if not items:
        return ""
    label = _PLATFORM_LABEL.get(platform, platform)
    lines = [f"📰 今日{label} Top {len(items)}"]
    for it in items:
        hot_str = ""
        if it.get("hot"):
            h = it["hot"]
            if h >= 100000000:
                hot_str = f" ({h/100000000:.2f}亿)"
            elif h >= 10000:
                hot_str = f" ({h/10000:.1f}万)"
            else:
                hot_str = f" ({h})"
        lines.append(f"{it['rank']}. {it['title']}{hot_str}")
    return "\n".join(lines)


def parse_news_arg(arg: Optional[str]) -> Tuple[str, int]:
    """解析 {{news:zhihu:5}} 参数: 平台 + limit。返回 (platform, limit)。"""
    platform = "zhihu"
    limit = 10
    if not arg:
        return platform, limit
    for part in arg.split(":"):
        part = (part or "").strip().lower()
        if not part:
            continue
        if part.isdigit():
            n = int(part)
            if 1 <= n <= 50:
                limit = n
        elif part in _PLATFORMS:
            platform = part
    return platform, limit
