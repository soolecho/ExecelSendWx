"""知乎/头条/B站/微博 热搜拉取 + 消息格式化。

占位符语法 (在 weather_fetcher.render_message 内解析):
  {{news}}              -> 知乎热榜 Top 10 (默认)
  {{news:zhihu}}        -> 知乎热榜 Top 10
  {{news:toutiao}}      -> 头条热搜 Top 10
  {{news:bilibili}}     -> B 站热门视频 Top 10
  {{news:weibo}}        -> 微博热搜 Top 10
  {{news:5}}            -> 知乎 Top 5
  {{news:zhihu:5}}      -> 知乎 Top 5
  {{news:weibo:3}}      -> 微博 Top 3

微博采用三级兜底源: 官方ajax -> 60s API(viki.moe) -> codelife(tophub聚合)。
所有平台条目均带可点击链接 (微博为搜索页链接, 点开直接搜索)。
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

_PLATFORMS = {"zhihu", "toutiao", "bilibili", "weibo"}

_CACHE_LOCK = threading.Lock()
_CACHE: Dict[str, Tuple[float, List[Dict]]] = {}
_CACHE_TTL = 300.0  # 5 分钟


def _cache_key(platform: str) -> str:
    return f"news|{platform}"


def _http_get_json(url: str, timeout: float = 8.0,
                   referer: Optional[str] = None) -> Dict:
    headers = {
        "User-Agent": _UA,
        "Accept": "application/json,text/plain,*/*",
        "Accept-Language": "zh-CN,zh;q=0.9",
    }
    if referer:
        headers["Referer"] = referer
    req = urllib.request.Request(url, headers=headers)
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


def _parse_hot_text(value) -> int:
    """解析热度文本: '129万' -> 1290000, '1.2亿' -> 120000000, '12345' -> 12345。"""
    try:
        s = str(value).strip()
        if s.endswith("亿"):
            return int(float(s[:-1]) * 100000000)
        if s.endswith("万"):
            return int(float(s[:-1]) * 10000)
        return int(s)
    except (TypeError, ValueError):
        return 0


def _weibo_search_url(word: str) -> str:
    return "https://s.weibo.com/weibo?q=" + urllib.parse.quote(f"#{word}#")


def fetch_weibo(limit: int = 10, log_fn=None) -> List[Dict]:
    """微博热搜。三级兜底源: 官方ajax -> 60s API -> codelife 聚合。

    官方端点存在 IP 频控 (间歇性 403), 因此必须多源兜底。
    """
    errors = []

    # 1) 微博官方 ajax (数据最全, 含真实热度数值)
    try:
        data = _http_get_json("https://weibo.com/ajax/side/hotSearch",
                              referer="https://weibo.com")
        rt = (data.get("data") or {}).get("realtime") or []
        items = []
        for x in rt:
            word = (x.get("word") or "").strip()
            if not word:
                continue
            try:
                hot = int(x.get("num") or 0)
            except (TypeError, ValueError):
                hot = 0
            items.append({
                "title": word,
                "hot": hot,
                "url": _weibo_search_url(word),
            })
        if items:
            return _normalize(items[:limit], "weibo")
        errors.append("官方端点返回为空")
    except Exception as exc:
        errors.append(f"官方: {exc}")

    # 2) 60s API (viki.moe 镜像)
    try:
        data = _http_get_json("https://60s.viki.moe/v2/weibo")
        items = []
        for x in (data.get("data") or []):
            title = (x.get("title") or x.get("word") or "").strip()
            if not title:
                continue
            items.append({
                "title": title,
                "hot": _parse_hot_text(x.get("hot_value")),
                "url": x.get("link") or _weibo_search_url(title),
            })
        if items:
            if log_fn:
                log_fn("ℹ️ 微博官方端点不可用({}), 已切换备用源".format(errors[-1]))
            return _normalize(items[:limit], "weibo")
        errors.append("60s API 返回为空")
    except Exception as exc:
        errors.append(f"60s: {exc}")

    # 3) codelife 聚合 (tophub 微博热榜)
    try:
        data = _http_get_json("https://api.codelife.cc/api/top/list?lang=cn&id=KqndgxeLl9",
                              referer="https://tophub.today/")
        items = []
        for x in (data.get("data") or []):
            title = (x.get("title") or "").strip()
            if not title:
                continue
            items.append({
                "title": title,
                "hot": _parse_hot_text(x.get("hotValue")),
                "url": x.get("link") or _weibo_search_url(title),
            })
        if items:
            if log_fn:
                log_fn("ℹ️ 微博官方端点不可用({}), 已切换备用源".format(errors[-1]))
            return _normalize(items[:limit], "weibo")
        errors.append("codelife 返回为空")
    except Exception as exc:
        errors.append(f"codelife: {exc}")

    msg = "❌ 微博热搜所有源均失败: " + "; ".join(errors)
    if log_fn:
        log_fn(msg)
    logger.warning("微博热搜拉取失败: %s", msg)
    return []


_FETCHERS = {
    "zhihu": fetch_zhihu,
    "toutiao": fetch_toutiao,
    "bilibili": fetch_bilibili,
    "weibo": fetch_weibo,
}

_PLATFORM_LABEL = {
    "zhihu": "知乎热榜",
    "toutiao": "头条热搜",
    "bilibili": "B站热门",
    "weibo": "微博热搜",
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
    """把热搜 List 渲染成多行文本消息, 每条附可点击链接。

    微信文本消息中的 URL 会自动识别为可点击链接;
    微博条目链接为搜索页, 点开直接显示该词条的搜索结果。
    """
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
        url = (it.get("url") or "").strip()
        if url:
            lines.append(f"   🔗 {url}")
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
