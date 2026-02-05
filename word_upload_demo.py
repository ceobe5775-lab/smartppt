from __future__ import annotations

import argparse
import html
import os
import threading
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

UPLOAD_DIR = Path("uploads")
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB


def is_allowed_word_file(filename: str) -> bool:
    """Return True when filename ends with .doc or .docx (case-insensitive)."""
    lowered = filename.lower()
    return lowered.endswith(".doc") or lowered.endswith(".docx")


class WordUploadHandler(BaseHTTPRequestHandler):
    def _render_form(self, message: str = "") -> bytes:
        escaped_message = html.escape(message)
        return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <title>Word 上传最小演示</title>
</head>
<body>
  <h1>Word 上传最小演示</h1>
  <p>支持 .doc / .docx，单文件最大 10MB。</p>
  <p style="color:#0a7">{escaped_message}</p>
  <form method="post" enctype="multipart/form-data">
    <input type="file" name="file" accept=".doc,.docx" required />
    <button type="submit">上传测试</button>
  </form>
</body>
</html>
""".encode("utf-8")

    def do_GET(self) -> None:
        body = self._render_form()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self) -> None:
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            self._send_form_with_message("上传失败：请求内容为空。")
            return

        if content_length > MAX_FILE_SIZE + 20_000:
            self._send_form_with_message("上传失败：请求过大（超过 10MB 限制）。")
            return

        content_type = self.headers.get("Content-Type", "")
        boundary_key = "boundary="
        if "multipart/form-data" not in content_type or boundary_key not in content_type:
            self._send_form_with_message("上传失败：请使用 multipart/form-data 表单上传。")
            return

        boundary = content_type.split(boundary_key, 1)[1].strip().encode()
        body = self.rfile.read(content_length)

        marker = b'filename="'
        marker_index = body.find(marker)
        if marker_index == -1:
            self._send_form_with_message("上传失败：未检测到文件字段。")
            return

        filename_start = marker_index + len(marker)
        filename_end = body.find(b'"', filename_start)
        filename = body[filename_start:filename_end].decode("utf-8", errors="ignore")
        filename = os.path.basename(filename)

        if not filename:
            self._send_form_with_message("上传失败：文件名为空。")
            return

        if not is_allowed_word_file(filename):
            self._send_form_with_message("上传失败：仅允许 .doc 或 .docx 文件。")
            return

        file_data_start = body.find(b"\r\n\r\n", filename_end)
        if file_data_start == -1:
            self._send_form_with_message("上传失败：无法解析上传内容。")
            return

        file_data_start += 4
        end_boundary = b"\r\n--" + boundary
        file_data_end = body.find(end_boundary, file_data_start)
        if file_data_end == -1:
            self._send_form_with_message("上传失败：文件结束边界未找到。")
            return

        file_bytes = body[file_data_start:file_data_end]
        if len(file_bytes) > MAX_FILE_SIZE:
            self._send_form_with_message("上传失败：文件超过 10MB 限制。")
            return

        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        save_path = UPLOAD_DIR / filename
        save_path.write_bytes(file_bytes)

        self._send_form_with_message(
            f"上传成功：{filename}（{len(file_bytes)} 字节），已保存到 {save_path}"
        )

    def _send_form_with_message(self, message: str) -> None:
        body = self._render_form(message)
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def run_server(host: str = "0.0.0.0", port: int = 8000, open_browser: bool = False) -> None:
    print(f"Starting server at http://{host}:{port}")

    if open_browser:
        browse_host = "127.0.0.1" if host == "0.0.0.0" else host
        browse_url = f"http://{browse_host}:{port}"
        print(f"Opening browser: {browse_url}")
        threading.Timer(0.6, lambda: webbrowser.open(browse_url)).start()

    with HTTPServer((host, port), WordUploadHandler) as httpd:
        httpd.serve_forever()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run a minimal Word upload demo server.")
    parser.add_argument("--host", default="0.0.0.0", help="Host to bind, default: 0.0.0.0")
    parser.add_argument("--port", type=int, default=8000, help="Port to bind, default: 8000")
    parser.add_argument(
        "--open-browser",
        action="store_true",
        help="Automatically open default browser after server starts.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    run_server(host=args.host, port=args.port, open_browser=args.open_browser)
