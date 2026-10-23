"""mask_server.py — 黒塗りの範囲を先生がブラウザで決めるための一時サーバー。

- 自分の PC からしか繋がらない（127.0.0.1 だけで待ち受け、Host ヘッダーも確かめる）
- ポートは起動のたびに OS が空きを選ぶ
- URL に推測できない合言葉（トークン）を付け、無いアクセスは断る
- 先生が承認したら、または時間切れになったら止まる

答案画像はこのサーバーから先生のブラウザにだけ送る。LLM はこの画像を取りに行かない
（手順書と AGENTS.md の約束）。
"""
import json
import secrets
import threading
import time
from collections import OrderedDict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from saiten_llm import mask as maskmod

PAGE_HTML = Path(__file__).with_name("mask_page.html")
MAX_BODY = 1_000_000
CACHE_PAGES = 24


class _AlignedCache:
    """傾き補正した用紙の画像を、少しだけ手元に置いておく（同じ答案を何度も補正しない）"""

    def __init__(self, files):
        self.files = files
        self._items = OrderedDict()
        self._lock = threading.Lock()

    def get(self, index):
        with self._lock:
            if index in self._items:
                self._items.move_to_end(index)
                return self._items[index]
        image, _aligned = maskmod.load_aligned(self.files[index])
        with self._lock:
            self._items[index] = image
            while len(self._items) > CACHE_PAGES:
                self._items.popitem(last=False)
        return image


class LocalPageServer:
    """先生のブラウザにだけ画面を出す一時サーバーの共通部分。

    派生クラスが page_html・handle_get・handle_post を決める。
    handle_get は (状態コード, 本文, Content-Type) か None（見つからない）を返す。
    handle_post は JSON にする dict か None を返す
    """
    page_html = PAGE_HTML

    def __init__(self, timeout_minutes=120):
        self.token = secrets.token_urlsafe(24)
        self.approved = threading.Event()
        self.deadline = time.monotonic() + timeout_minutes * 60
        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), self._handler_class())
        self.httpd.daemon_threads = True

    @property
    def port(self):
        return self.httpd.server_address[1]

    @property
    def url(self):
        return f"http://127.0.0.1:{self.port}/?t={self.token}"

    def handle_get(self, path, query) -> "tuple | None":
        return None

    def handle_post(self, path, body) -> "dict | None":
        return None

    # --- HTTP -------------------------------------------------------

    def _handler_class(self):
        server = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, fmt, *args):   # 画面をアクセス記録で埋めない
                pass

            def _allowed(self, query):
                host = self.headers.get("Host", "")
                if host not in (f"127.0.0.1:{server.port}", f"localhost:{server.port}"):
                    return False
                given = (query.get("t") or [""])[0]
                return secrets.compare_digest(given, server.token)

            def _send(self, status, body, ctype):
                self.send_response(status)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("X-Content-Type-Options", "nosniff")
                self.send_header("Referrer-Policy", "no-referrer")
                self.end_headers()
                self.wfile.write(body)

            def _json(self, data, status=200):
                self._send(status, json.dumps(data, ensure_ascii=False).encode("utf-8"),
                           "application/json; charset=utf-8")

            def do_GET(self):
                url = urlparse(self.path)
                query = parse_qs(url.query)
                if not self._allowed(query):
                    return self._json({"ok": False, "error": "このリンクは使えません"}, 403)
                try:
                    if url.path == "/":
                        html = server.page_html.read_text(encoding="utf-8").replace(
                            "__TOKEN__", server.token)
                        return self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
                    res = server.handle_get(url.path, query)
                    if res is not None:
                        status, body, ctype = res
                        if ctype == "json":
                            return self._json(body, status)
                        return self._send(status, body, ctype)
                except (KeyError, ValueError) as e:
                    return self._json({"ok": False, "error": str(e)}, 400)
                return self._json({"ok": False, "error": "見つかりません"}, 404)

            def do_POST(self):
                url = urlparse(self.path)
                query = parse_qs(url.query)
                if not self._allowed(query):
                    return self._json({"ok": False, "error": "このリンクは使えません"}, 403)
                length = int(self.headers.get("Content-Length") or 0)
                if length > MAX_BODY:
                    return self._json({"ok": False, "error": "送られた内容が大きすぎます"}, 413)
                try:
                    body = json.loads(self.rfile.read(length) or b"{}")
                    res = server.handle_post(url.path, body)
                    if res is not None:
                        return self._json(res)
                except ValueError as e:
                    return self._json({"ok": False, "error": str(e)}, 400)
                return self._json({"ok": False, "error": "見つかりません"}, 404)

        return Handler

    # --- 起動と停止 -------------------------------------------------

    def serve_until_done(self, poll=0.5):
        """承認されるか時間切れまで待つ。Returns: 承認されたら True"""
        thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)
        thread.start()
        try:
            while not self.approved.is_set() and time.monotonic() < self.deadline:
                self.approved.wait(poll)
            # 承認の応答がブラウザに届くよう、少し待ってから止める
            time.sleep(0.5)
        finally:
            self.httpd.shutdown()
            self.httpd.server_close()
        return self.approved.is_set()


class MaskServer(LocalPageServer):
    def __init__(self, project, timeout_minutes=120):
        self.project = project
        self.files = project.scan_files()
        if not self.files:
            raise FileNotFoundError("答案の画像がありません。先に pdf コマンドで画像にしてください")
        self.names = [p.name for p in self.files]
        self.cache = _AlignedCache(self.files)
        saved = project.data.get("mask") or {}
        self.spec = {"regions": saved.get("regions", []), "extra": saved.get("extra", {})}
        super().__init__(timeout_minutes)

    # --- 処理 -------------------------------------------------------

    def state(self):
        return {"ok": True, "pages": [{"name": n} for n in self.names],
                "spec": self.spec}

    def image(self, index, width, masked):
        img = self.cache.get(index)
        if masked:
            img = maskmod.apply_mask(img, maskmod.rects_for(self.spec, self.names[index]))
        return maskmod.encode_jpeg(img, max_width=width)

    def save(self, body, approve=False):
        spec = maskmod.validate_spec(body, set(self.names))
        self.spec = spec
        maskmod.save_spec(self.project, spec, approved=approve)
        if approve:
            self.approved.set()
        return {"ok": True}

    def handle_get(self, path, query):
        if path == "/api/state":
            return 200, self.state(), "json"
        if path == "/img":
            index = int(query["i"][0])
            if not 0 <= index < len(self.names):
                return 404, {"ok": False, "error": "答案がありません"}, "json"
            width = max(100, min(2000, int(query.get("w", ["1200"])[0])))
            masked = query.get("m", ["0"])[0] == "1"
            return 200, self.image(index, width, masked), "image/jpeg"
        return None

    def handle_post(self, path, body):
        if path == "/api/save":
            return self.save(body)
        if path == "/api/approve":
            return self.save(body, approve=True)
        return None
