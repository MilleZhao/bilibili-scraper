# -*- coding: utf-8 -*-
"""
B站爬取引擎 -- 零依赖，仅用标准库。
对单个 BV 号爬取：视频信息 + 评论（含子回复） + 弹幕。

供 server.py 调用:
    get_video_info(bvid) -> dict
    collect_comments(aid, max_pages, on_progress) -> list[dict]
    collect_danmaku(cid, max_segments, on_progress) -> list[dict]
    save_outputs(vi, comments, danmaku, out_dir) -> None
"""

import csv
import gzip
import json
import os
import random
import re
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

BASE_URL = "https://api.bilibili.com"
VIDEO_INFO = f"{BASE_URL}/x/web-interface/view"
REPLY_MAIN = f"{BASE_URL}/x/v2/reply/main"
DM_SEG = f"{BASE_URL}/x/v2/dm/web/seg.so"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Referer": "https://www.bilibili.com/",
    "Origin": "https://www.bilibili.com",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}

COOKIE_KEYS = ("SESSDATA", "BILI_JCT", "DEDEUSERID", "BUVID3")
_cache = {}

# ---- Helpers ----

def ts_to_iso(ts):
    if not ts:
        return ""
    if isinstance(ts, str):
        return ts[:19]
    return datetime.fromtimestamp(int(ts), tz=timezone.utc).astimezone().isoformat(timespec="seconds")

def safe_int(v, default=0):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default

# ---- Cookie ----

def _find_my_cookies():
    here = Path(__file__).resolve().parent
    for root, dirs, files in os.walk(str(here)):
        if "my_cookies.py" in files:
            return Path(root) / "my_cookies.py"
    for root, dirs, files in os.walk(str(here.parent)):
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
            elif key == "BILI_JCT":
                parts.append(f"bili_jct={val}")
            elif key == "DEDEUSERID":
                parts.append(f"DedeUserID={val}")
            elif key == "BUVID3":
                parts.append(f"buvid3={val}")
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
                elif key == "BILI_JCT":
                    parts.append(f"bili_jct={val}")
                elif key == "DEDEUSERID":
                    parts.append(f"DedeUserID={val}")
                elif key == "BUVID3":
                    parts.append(f"buvid3={val}")
    return "; ".join(parts)

def get_cookie():
    if "cookie" not in _cache:
        _cache["cookie"] = load_cookies()
    return _cache["cookie"]

# ---- HTTP ----

def _request(url, params=None, timeout=20):
    if params:
        sep = "&" if "?" in url else "?"
        url = f"{url}{sep}{urlencode(params, doseq=True)}"
    headers = dict(HEADERS)
    cookie = get_cookie()
    if cookie:
        headers["Cookie"] = cookie
    req = Request(url, headers=headers, method="GET")
    return urlopen(req, timeout=timeout).read()

def fetch_json(url, params=None, retries=5):
    for attempt in range(retries):
        try:
            data = _request(url, params)
            result = json.loads(data.decode("utf-8"))
            code = result.get("code", -1)
            if code != 0:
                msg = str(result.get("message") or result.get("msg") or "")
                if code in (-403, -412, -352) or any(kw in msg for kw in ("frequency", "limit", "restricted")):
                    wait = min(60, (attempt + 1) * 10)
                    time.sleep(wait)
                    continue
                raise RuntimeError(f"API error code={code}: {msg}")
            return result
        except (URLError, HTTPError, OSError) as e:
            if attempt == retries - 1:
                raise
            time.sleep(random.uniform(2, 5))
    raise RuntimeError(f"Failed after {retries} retries: {url}")

# ---- Video Info ----

def get_video_info(bvid):
    result = fetch_json(VIDEO_INFO, {"bvid": bvid})
    data = result.get("data", {}) or {}
    stat = data.get("stat", {}) or {}
    owner = data.get("owner", {}) or {}
    pub_ts = data.get("pubdate", 0)
    return {
        "bvid": data.get("bvid", bvid),
        "aid": data.get("aid", 0),
        "cid": data.get("cid", 0),
        "title": (data.get("title") or "").strip(),
        "author": (owner.get("name") or "").strip(),
        "author_mid": owner.get("mid", 0),
        "duration_sec": data.get("duration", 0),
        "pubdate": ts_to_iso(pub_ts) if pub_ts else "",
        "tname": data.get("tname", ""),
        "view": stat.get("view", 0),
        "danmaku_count": stat.get("danmaku", 0),
        "reply_count": stat.get("reply", 0),
        "like": stat.get("like", 0),
        "coin": stat.get("coin", 0),
        "favorite": stat.get("favorite", 0),
        "share": stat.get("share", 0),
    }

# ---- Comments ----

def collect_comments(aid, max_pages=50, on_progress=None):
    all_comments = []
    cursor = 0
    page = 1
    if max_pages <= 0:
        max_pages = 9999
    while page <= max_pages:
        try:
            result = fetch_json(REPLY_MAIN, {
                "type": 1, "oid": aid, "mode": 3, "plat": 1,
                "next": cursor, "ps": 49,
            })
        except Exception:
            break
        reply_data = result.get("data", {}) or {}
        replies = reply_data.get("replies") or []
        if not replies:
            break
        for reply in replies:
            member = reply.get("member", {}) or {}
            content_obj = reply.get("content", {}) or {}
            rpid = reply.get("rpid", 0)
            all_comments.append({
                "rpid": rpid,
                "parent_rpid": 0,
                "is_sub": False,
                "username": (member.get("uname") or "").strip(),
                "user_mid": member.get("mid", 0),
                "content": (content_obj.get("message") or "").replace("\n", " ").strip(),
                "timestamp": ts_to_iso(reply.get("ctime", 0)),
                "likes": reply.get("like", 0),
                "reply_count": reply.get("rcount", 0),
            })
            for sub in (reply.get("replies") or []):
                sub_member = sub.get("member", {}) or {}
                sub_content = sub.get("content", {}) or {}
                all_comments.append({
                    "rpid": sub.get("rpid", 0),
                    "parent_rpid": rpid,
                    "is_sub": True,
                    "username": (sub_member.get("uname") or "").strip(),
                    "user_mid": sub_member.get("mid", 0),
                    "content": (sub_content.get("message") or "").replace("\n", " ").strip(),
                    "timestamp": ts_to_iso(sub.get("ctime", 0)),
                    "likes": sub.get("like", 0),
                    "reply_count": 0,
                })
        if on_progress:
            on_progress(len(all_comments), None)
        cursor = safe_int((reply_data.get("cursor", {}) or {}).get("next"), 0)
        if cursor <= 0:
            break
        page += 1
        time.sleep(random.uniform(0.05, 0.15))
    return all_comments

# ---- Danmaku (hand-rolled protobuf parser) ----

def _read_varint(data, offset):
    result = 0
    shift = 0
    while offset < len(data):
        byte = data[offset]
        offset += 1
        result |= (byte & 0x7F) << shift
        shift += 7
        if not (byte & 0x80):
            break
    return result, offset

def _parse_danmaku_elem(data):
    dm = {"id": 0, "progress": 0, "mode": 0, "fontsize": 0, "color": "ffffff",
          "mid_hash": "", "content": "", "ctime": 0, "weight": 0,
          "action": "", "pool": 0, "id_str": "", "attr": 0}
    offset = 0
    while offset < len(data):
        tag, offset = _read_varint(data, offset)
        field_num = tag >> 3
        wire_type = tag & 0x07
        if wire_type == 0:
            val, offset = _read_varint(data, offset)
            mapping = {1: "id", 2: "progress", 3: "mode", 4: "fontsize",
                       5: "color", 8: "ctime", 9: "weight", 11: "pool", 13: "attr"}
            if field_num in mapping:
                dm[mapping[field_num]] = val
        elif wire_type == 2:
            length, offset = _read_varint(data, offset)
            val_data = data[offset:offset + length]
            offset += length
            mapping = {6: "mid_hash", 7: "content", 10: "action", 12: "id_str"}
            if field_num in mapping:
                try:
                    dm[mapping[field_num]] = val_data.decode("utf-8")
                except UnicodeDecodeError:
                    dm[mapping[field_num]] = val_data.decode("utf-8", errors="replace")
        else:
            break
    if dm.get("color"):
        dm["color"] = format(dm["color"], "06x")
    if dm.get("progress"):
        dm["progress"] = round(dm["progress"] / 1000.0, 3)
    return dm, offset

def _parse_danmaku_protobuf(data):
    result = []
    offset = 0
    while offset < len(data):
        try:
            tag, offset = _read_varint(data, offset)
        except (IndexError, ValueError):
            break
        field_num = tag >> 3
        wire_type = tag & 0x07
        if field_num == 1 and wire_type == 2:
            length, offset = _read_varint(data, offset)
            inner = data[offset:offset + length]
            offset += length
            inner_off = 0
            while inner_off < len(inner):
                try:
                    dm, consumed = _parse_danmaku_elem(inner[inner_off:])
                    if consumed == 0:
                        break
                    if dm.get("content"):
                        result.append(dm)
                    inner_off += consumed
                except Exception:
                    break
        else:
            if wire_type == 0:
                _, offset = _read_varint(data, offset)
            elif wire_type == 2:
                length, offset = _read_varint(data, offset)
                offset += length
            else:
                break
    return result

def collect_danmaku(cid, max_segments=1200, on_progress=None):
    all_dm = []
    segment = 1
    total_empty = 0
    EMPTY_LIMIT = 3 if max_segments <= 100 else 12
    while segment <= max_segments:
        try:
            raw = _request(f"{DM_SEG}?type=1&oid={cid}&segment_index={segment}")
        except Exception:
            break
        if len(raw) >= 2 and raw[0] == 0x1F and raw[1] == 0x8B:
            raw = gzip.decompress(raw)
        if not raw:
            total_empty += 1
            if total_empty >= EMPTY_LIMIT:
                break
            segment += 1
            continue
        try:
            dm_list = _parse_danmaku_protobuf(raw)
        except Exception:
            total_empty += 1
            segment += 1
            continue
        if dm_list:
            for dm in dm_list:
                dm["segment"] = segment
            all_dm.extend(dm_list)
            total_empty = 0
            if on_progress:
                on_progress(len(all_dm), None)
        else:
            total_empty += 1
        if total_empty >= EMPTY_LIMIT:
            break
        segment += 1
        time.sleep(random.uniform(0.05, 0.12))
    return all_dm

# ---- Save ----

def save_json(data, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def _append_consolidated(out_root, filename, new_items):
    '''Append items to a consolidated JSON array file, thread-safe.'''
    import json as _json
    from pathlib import Path as _Path
    fpath = _Path(out_root) / filename
    existing = []
    if fpath.exists():
        try:
            existing = _json.loads(fpath.read_text(encoding='utf-8'))
        except:
            existing = []
    existing.extend(new_items)
    fpath.write_text(_json.dumps(existing, ensure_ascii=False, indent=2), encoding='utf-8')


def save_outputs(video_info, comments, danmaku, out_dir):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    save_json({k: v for k, v in video_info.items() if k != "pubdate_ts"}, out_dir / "video_info.json")
    save_json(comments, out_dir / "comments.json")
    save_json(danmaku, out_dir / "danmaku.json")
    # CSV
    vi = {k: v for k, v in video_info.items() if k != "pubdate_ts"}
    with (out_dir / "video_info.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(vi.keys()))
        w.writeheader(); w.writerow(vi)
    cf = ["rpid", "parent_rpid", "is_sub", "username", "user_mid", "content", "timestamp", "likes", "reply_count"]
    with (out_dir / "comments.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cf, extrasaction="ignore")
        w.writeheader(); w.writerows(comments)
    df = ["content", "progress", "mode", "fontsize", "color", "ctime", "pool", "mid_hash", "weight", "segment"]
    with (out_dir / "danmaku.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=df, extrasaction="ignore")
        w.writeheader(); w.writerows(danmaku)
    # Also append to consolidated files for data-cleaning scripts
    _append_consolidated(out_dir.parent, "comments_all.json", comments)
    _append_consolidated(out_dir.parent, "danmaku_all.json", danmaku)
