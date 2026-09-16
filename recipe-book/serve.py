#!/usr/bin/env python3
"""Local kitchen-book server. Serves the site and saves edits to disk."""
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import json
import re
import base64

ROOT = Path(__file__).resolve().parent
PORT = 8765


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-cache")
        super().end_headers()

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)

        if self.path == "/api/save":
            data = json.loads(body.decode("utf-8"))
            recipes = data.get("recipes")
            if not isinstance(recipes, list):
                return self._send(400, {"ok": False, "error": "recipes array required"})
            (ROOT / "recipes.json").write_text(
                json.dumps(recipes, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            (ROOT / "recipes.js").write_text(
                "window.RECIPES = " + json.dumps(recipes, ensure_ascii=False) + ";\n",
                encoding="utf-8",
            )
            saved_photos = 0
            photos = data.get("photos") or {}
            for key, data_url in photos.items():
                if self._write_photo(key, data_url):
                    saved_photos += 1
            return self._send(200, {"ok": True, "count": len(recipes), "photos": saved_photos})

        m = re.match(r"^/api/photo/(\d+)$", self.path)
        if m:
            item_id = int(m.group(1))
            ctype = self.headers.get("Content-Type", "")
            if "application/json" in ctype:
                payload = json.loads(body.decode("utf-8"))
                ok = self._write_photo(item_id, payload.get("data") or "")
            else:
                ok = self._write_bytes(item_id, body)
            if not ok:
                return self._send(400, {"ok": False, "error": "could not save photo"})
            return self._send(200, {"ok": True, "id": item_id})

        self._send(404, {"ok": False, "error": "unknown endpoint"})

    def _write_photo(self, key, data_url):
        if not data_url or "," not in str(data_url):
            return False
        raw = base64.b64decode(str(data_url).split(",", 1)[1])
        return self._write_bytes(key, raw)

    def _write_bytes(self, key, raw):
        try:
            item_id = int(key)
        except (TypeError, ValueError):
            return False
        if not 1 <= item_id <= 999:
            return False
        if not raw:
            return False
        path = ROOT / "images" / f"{item_id:03d}.jpg"
        path.write_bytes(raw)
        return True

    def _send(self, code, obj):
        payload = json.dumps(obj).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def log_message(self, fmt, *args):
        if "/api/" in str(args[0] if args else ""):
            super().log_message(fmt, *args)


if __name__ == "__main__":
    httpd = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"TCM recipe book: http://127.0.0.1:{PORT}/")
    print("Edit on this computer, then Save to disk. Ctrl+C to stop.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nstopped")
