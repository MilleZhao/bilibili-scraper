# -*- coding: utf-8 -*-
"""
B站通用搜索引擎 -- 零依赖，仅用标准库。
通过 B站搜索 API，按关键词 + 分区 + 排序检索视频。

供 server.py 调用的核心函数:
    search_videos(keywords, tid, sort_order, pages, on_progress) -> list[dict]
    filter_by_keywords(candidates, filter_words) -> list[dict]
    get_cookie() -> str
    get_wbi_key() -> str
"""

import argparse
import csv
import hashlib
import json
import os
import random
import re
import sys
import time
import urllib.parse
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

BASE = "https://api.bilibili.com"
SEARCH_API = f"{BASE}/x/web-interface/search/type"
NAV_API = f"{BASE}/x/web-interface/nav"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Referer": "https://www.bilibili.com/",
    "Origin": "https://www.bilibili.com",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}

# 搜索用（由调用方传入，这里留空占位）
KEYWORDS = []

SORT_OPTIONS = {
    "click": "最多播放",
    "pubdate": "最新发布",
    "stow": "最多收藏",
    "dm": "最多弹幕",
}

# B站主流分区（tid -> 名称），search_videos 中 tid=0 表示全站
CATEGORIES = {
    0: "全站",
    1: "动画", 13: "番剧", 167: "国创",
    3: "音乐", 129: "舞蹈", 4: "游戏",
    36: "知识", 188: "科技", 234: "运动",
    223: "汽车", 160: "生活", 211: "美食",
    217: "动物圈", 119: "鬼畜", 155: "时尚",
    5: "娱乐", 181: "影视", 177: "纪录片",
    23: "电影", 11: "电视剧",
}

_CACHE = {}

def log(msg):
    stamp = time.strftime("%H:%M:%S")
    print(f"[{stamp}] {msg}", flush=True)

def matches_keywords(text, words):
    """检查 text 是否包含 words 中任意一个关键词（小写匹配）"""
    if not words:
        return True
    t = (text or "").lower()
    return any(w.lower() in t for w in words)

# ---- Cookie ----

COOKIE_KEYS = ("SESSDATA", "BILI_JCT", "DEDEUSERID", "BUVID3")

def _find_my_cookies():
    here = Path(__file__).resolve().parent
    for base in [here, here.parent]:
        for root, dirs, files in os.walk(str(base)):
            if "my_cookies.py" in files:
                return Path(root) / "my_cookies.py"
    return None

def load_cookies():
    parts = []
    for key in COOKIE_KEYS:
        val = os.environ.get(key, "")
        if val:
            if key == "SESSDATA":
                parts.append(f"SESSDATA={val}")
            else:
                parts.append(f"{key}={val}")
    if parts:
        return "; ".join(parts)
    cp = _find_my_cookies()
    if cp:
        ns = {}
        exec(cp.read_text(encoding="utf-8-sig"), ns)
        for key in COOKIE_KEYS:
            val = ns.get(key, "")
            if val:
                if key == "SESSDATA":
                    parts.append(f"SESSDATA={val}")
                else:
                    parts.append(f"{key}={val}")
    return "; ".join(parts)

def get_cookie():
    if "cookie" not in _CACHE:
        _CACHE["cookie"] = load_cookies()
    return _CACHE["cookie"]

# ---- Wbi Signing ----

MIXIN = [
    46, 47, 18, 2, 53, 8, 23, 32, 15, 50, 10, 31, 58, 3, 45, 35,
    27, 43, 5, 49, 33, 9, 42, 19, 29, 28, 14, 39, 12, 38, 41, 13,
    37, 48, 7, 16, 24, 55, 40, 61, 26, 17, 0, 1, 60, 51, 30, 4,
    22, 25, 54, 21, 56, 59, 6, 63, 57, 62, 11, 36, 20, 34, 44, 52,
]

def _mix(img_key, sub_key):
    s = img_key + sub_key
    return "".join(s[i] for i in MIXIN[:32])

def _fetch_wbi_key():
    headers = dict(HEADERS)
    cookie = get_cookie()
    if cookie:
        headers["Cookie"] = cookie
    try:
        req = Request(NAV_API, headers=headers)
        with urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8", errors="replace"))
        wbi = (data.get("data", {}) or {}).get("wbi_img", {}) or {}
        img_url = wbi.get("img_url", "")
        sub_url = wbi.get("sub_url", "")
        if img_url and sub_url:
            img_key = img_url.rsplit("/", 1)[-1].split(".")[0]
            sub_key = sub_url.rsplit("/", 1)[-1].split(".")[0]
            return _mix(img_key, sub_key)
    except Exception as e:
        log(f"Wbi key fetch failed: {e}")
    return ""

def get_wbi_key():
    if "wbi" not in _CACHE:
        _CACHE["wbi"] = _fetch_wbi_key()
    return _CACHE["wbi"]

def sign_params(params):
    key = get_wbi_key()
    if not key:
        return params
    p = dict(params)
    p["wts"] = int(time.time())
    keys = sorted(p.keys())
    raw = "&".join(f"{k}={urllib.parse.quote(str(p[k]), safe='')}" for k in keys)
    p["w_rid"] = hashlib.md5((raw + key).encode()).hexdigest()
    return p

# ---- HTTP ----

def request_json(url, params, retries=5):
    for attempt in range(retries):
        if attempt > 0:
            time.sleep(random.uniform(1.5, 4.0) * (attempt + 1))
        signed = sign_params(params)
        qs = urlencode(signed, doseq=True)
        full_url = f"{url}?{qs}"
        headers = dict(HEADERS)
        cookie = get_cookie()
        if cookie:
            headers["Cookie"] = cookie
        req = Request(full_url, headers=headers, method="GET")
        try:
            with urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8", errors="replace"))
            if isinstance(data, dict):
                code = data.get("code", -1)
                if code not in (0, None):
                    msg = data.get("message", "") or data.get("msg", "")
                    if code in (-403, -412, -352):
                        wait = min(45, (attempt + 1) * 8)
                        log(f"rate-limited (code={code}), wait {wait}s")
                        time.sleep(wait)
                        continue
                    log(f"API code={code}: {msg}")
                return data
        except (HTTPError, URLError, OSError) as e:
            if attempt >= retries - 1:
                log(f"giving up on this request after {retries} retries"); return {}
            log(f"network error (attempt {attempt+1}): {e}")
    return {}

# ---- Search ----

def search_videos(keywords, tid=0, sort_order="click", pages=8, on_progress=None):
    """通用搜索。keywords 为空时浏览分区热门。"""
    if not keywords:
        keywords = [""]  # empty keyword = browse category
    """通用搜索：关键词 + 分区 + 排序 -> 视频列表"""
    seen = set()
    candidates = []
    total = len(keywords) * pages
    done = 0
    for kw in keywords:
        for page in range(1, pages + 1):
            done += 1
            data = request_json(SEARCH_API, {
                "search_type": "video",
                "keyword": kw,
                "page": page,
                "order": sort_order,
                "tid": tid,
            })
            results = (data.get("data", {}) or {}).get("result", []) or []
            if not results:
                if on_progress:
                    on_progress({"kw": kw, "page": page, "done": done, "total": total, "candidates": len(candidates)})
                break
            for item in results:
                bvid = item.get("bvid", "")
                if not bvid or bvid in seen:
                    continue
                seen.add(bvid)
                candidates.append(item)
            if on_progress:
                on_progress({"kw": kw, "page": page, "done": done, "total": total, "candidates": len(candidates)})
            time.sleep(random.uniform(0.3, 0.7))
        if on_progress and len(candidates) % 50 == 0:
            log(f"candidates accumulated: {len(candidates)}")
    return candidates


def filter_by_keywords(candidates, filter_words):
    """对搜索结果做标题/描述的二次关键词过滤"""
    if not filter_words:
        return candidates
    return [item for item in candidates if (
        matches_keywords(item.get("title", ""), filter_words) or
        matches_keywords(item.get("description", "") or item.get("desc", ""), filter_words) or
        matches_keywords(item.get("tag", ""), filter_words)
    )]




import re

def _strip_html(text):
    """去掉 B站搜索结果中的 <em> 等 HTML 标签"""
    return re.sub(r"<[^>]+>", "", text)
def _parse_duration(val):
    """兼容 MM:SS 字符串和整数秒"""
    if val is None:
        return 0
    if isinstance(val, int):
        return val
    if isinstance(val, str) and ":" in val:
        parts = val.split(":")
        return int(parts[0]) * 60 + int(parts[1])
    try:
        return int(val)
    except (ValueError, TypeError):
        return 0
def search_and_filter(keywords, tid=0, sort_order="click", pages=8, target=200, filter_words=None, on_progress=None):
    """一步完成搜索+过滤+排序+截断，返回标准化结果列表"""
    candidates = search_videos(keywords, tid, sort_order, pages, on_progress=on_progress)
    if filter_words:
        candidates = filter_by_keywords(candidates, filter_words)
    candidates.sort(key=lambda x: int(x.get("play") or 0), reverse=True)
    selected = candidates[:target]
    results = []
    for item in selected:
        stat = item.get("stat", {}) or {}
        owner = item.get("owner", {}) or {}
        results.append({
            "bvid": item.get("bvid", ""),
            "aid": item.get("aid", 0),
            "title": _strip_html((item.get("title") or "").strip()),
            "author": (owner.get("name") or "").strip(),
            "views": int(item.get("play") or 0),
            "duration": _parse_duration(item.get("duration")),
            "category": item.get("typename", "") or item.get("tname", "") or CATEGORIES.get(tid, ""),
            "pubdate": item.get("pubdate", 0),
        })
    return results