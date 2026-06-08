# -*- coding: utf-8 -*-
"""B站通用筛选爬取平台 -- Web 服务端（轮询模式，零依赖）"""
from __future__ import annotations

import json
import sys
import threading
import time
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn
from pathlib import Path
from urllib.parse import urlparse

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from bilibili_search_engine import (
    search_and_filter, CATEGORIES, SORT_OPTIONS,
    get_cookie as engine_get_cookie, get_wbi_key,
)
from bilibili_scraper import (
    get_video_info, collect_comments, collect_danmaku, save_outputs,
)

# Global event store (list of dicts), protected by lock
events: list[dict] = []
events_lock = threading.Lock()
event_id_counter = [0]

def emit(event_type: str, data: dict):
    with events_lock:
        event_id_counter[0] += 1
        events.append({"id": event_id_counter[0], "event": event_type, "data": data})
        # Keep last 200 events only
        while len(events) > 200:
            events.pop(0)

# ---- Search Worker ----
def search_worker(keywords, tid, sort_order, target, fast, filter_words):
    pages = 3 if fast else 8
    def progress_cb(info):
        emit("search_progress", info)
    emit("status", {"msg": "搜索中..."})
    try:
        results = search_and_filter(
            keywords, tid=tid, sort_order=sort_order,
            pages=pages, target=target, filter_words=filter_words,
            on_progress=progress_cb,
        )
        emit("search_done", {"results": results, "total": len(results)})
    except Exception as e:
        emit("search_error", {"error": str(e)})

# ---- Scrape Worker ----
def scrape_worker(bvids, options):
    opts = options or {}
    do_vi = opts.get("video_info", True)
    do_cm = opts.get("comments", True)
    do_dm = opts.get("danmaku", True)
    cm_pages = opts.get("comment_pages", 50)
    dm_segs = opts.get("danmaku_segments", 1200)

    output_root = HERE / "outputs"
    output_root.mkdir(parents=True, exist_ok=True)
    total = len(bvids)
    emit("status", {"msg": f"开始爬取 {total} 个视频..."})

    for idx, bvid in enumerate(bvids):
        emit("scrape_progress", {"bvid": bvid, "phase": "init", "video_index": idx+1, "video_total": total})
        try:
            vi = get_video_info(bvid) if do_vi else {"bvid": bvid, "aid": 0, "cid": 0}
            aid = vi.get("aid", 0)
            cid = vi.get("cid", 0)
            title = vi.get("title", "")

            comments = []
            if do_cm and aid:
                def cm_cb(cur, _):
                    emit("scrape_progress", {"bvid": bvid, "phase": "comments", "count": cur, "video_index": idx+1, "video_total": total})
                comments = collect_comments(aid, max_pages=cm_pages, on_progress=cm_cb)

            danmaku = []
            if do_dm and cid:
                def dm_cb(cur, _):
                    emit("scrape_progress", {"bvid": bvid, "phase": "danmaku", "count": cur, "video_index": idx+1, "video_total": total})
                danmaku = collect_danmaku(cid, max_segments=dm_segs, on_progress=dm_cb)

            out_dir = output_root / bvid
            save_outputs(vi, comments, danmaku, out_dir)
            emit("video_done", {"bvid": bvid, "title": title, "comments": len(comments), "danmaku": len(danmaku)})
        except Exception as e:
            emit("video_error", {"bvid": bvid, "error": str(e)})
        time.sleep(1.0 + (idx % 3) * 0.5)

    emit("scrape_done", {"output_dir": str(output_root)})
    emit("status", {"msg": "爬取完成"})


class ThreadingHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True

# ---- HTTP Handler ----
class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def _json(self, data, code=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _html(self):
        p = HERE / "index.html"
        if not p.exists():
            self.send_error(404); return
        body = p.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urlparse(self.path).path
        if path in ("/", "/index.html"):
            return self._html()
        if path == "/api/categories":
            return self._json([{"tid": k, "name": v} for k, v in CATEGORIES.items()])
        if path == "/api/status":
            return self._json({"cookie_ok": bool(engine_get_cookie()), "wbi_ok": bool(get_wbi_key())})
        if path == "/api/events":
            qs = urlparse(self.path).query
            since = 0
            if qs:
                from urllib.parse import parse_qs
                p = parse_qs(qs)
                since = int(p.get("since", [0])[0])
            with events_lock:
                new = [e for e in events if e["id"] > since]
            return self._json({"events": new, "latest_id": event_id_counter[0]})
        self.send_error(404)

    def do_POST(self):
        path = urlparse(self.path).path
        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length)) if length else {}

        if path == "/api/search":
            kw = body.get("keywords", [])
            tid = body.get("tid", 0)
            sort = body.get("sort", "click")
            target = body.get("target", 200)
            fast = body.get("fast", False)
            filter_words = body.get("filter_words", None)
            threading.Thread(target=search_worker, args=(kw, tid, sort, target, fast, filter_words), daemon=True).start()
            return self._json({"ok": True})

        if path == "/api/scrape":
            bvids = body.get("bvids", [])
            options = body.get("options", {})
            threading.Thread(target=scrape_worker, args=(bvids, options), daemon=True).start()
            return self._json({"ok": True})

        self.send_error(404)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

def main():
    port = 8765
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"\n  B站通用筛选爬取平台 -> http://127.0.0.1:{port}\n")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  shutdown")
        server.shutdown()

if __name__ == "__main__":
    main()
