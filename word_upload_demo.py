from __future__ import annotations

import argparse
import html
import json
import os
import threading
import webbrowser
import zipfile
from dataclasses import dataclass
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Iterable
from urllib.parse import parse_qs, urlencode, urlparse
from xml.etree import ElementTree as ET

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB per file
MAX_FILES = 50

DOCX_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


@dataclass
class ParagraphBlock:
    text: str
    is_heading: bool = False


def is_allowed_word_file(filename: str) -> bool:
    """Return True when filename ends with .doc or .docx (case-insensitive)."""
    lowered = filename.lower()
    return lowered.endswith(".doc") or lowered.endswith(".docx")


def sanitize_filename(filename: str) -> str:
    return os.path.basename(filename).replace("..", "").strip()


def extract_docx_paragraphs(file_path: Path) -> list[ParagraphBlock]:
    """Extract text blocks from .docx and mark heading paragraphs when style says Heading*."""
    with zipfile.ZipFile(file_path, "r") as archive:
        try:
            document_xml = archive.read("word/document.xml")
        except KeyError as exc:
            raise ValueError("DOCX 结构异常：缺少 word/document.xml") from exc

    root = ET.fromstring(document_xml)
    blocks: list[ParagraphBlock] = []

    for paragraph in root.findall(".//w:body/w:p", DOCX_NS):
        texts = [node.text or "" for node in paragraph.findall(".//w:t", DOCX_NS)]
        text = "".join(texts).strip()
        if not text:
            continue

        style_node = paragraph.find("./w:pPr/w:pStyle", DOCX_NS)
        style_val = ""
        if style_node is not None:
            style_val = style_node.attrib.get(f"{{{DOCX_NS['w']}}}val", "")

        is_heading = style_val.lower().startswith("heading")
        blocks.append(ParagraphBlock(text=text, is_heading=is_heading))

    return blocks


def paginate_blocks(blocks: Iterable[ParagraphBlock], max_chars_per_page: int = 800) -> list[dict]:
    pages: list[dict] = []
    current_title = "未命名页"
    current_lines: list[str] = []

    def flush_page() -> None:
        nonlocal current_lines, current_title
        if not current_lines:
            return
        content = "\n".join(current_lines).strip()
        if not content:
            return
        pages.append(
            {
                "title": current_title,
                "content": content,
                "char_count": len(content),
            }
        )
        current_lines = []

    for block in blocks:
        if block.is_heading:
            flush_page()
            current_title = block.text
            continue

        prospective = "\n".join(current_lines + [block.text]).strip()
        if current_lines and len(prospective) > max_chars_per_page:
            flush_page()
        current_lines.append(block.text)

    flush_page()

    if not pages:
        pages.append({"title": current_title, "content": "", "char_count": 0})

    for idx, page in enumerate(pages, start=1):
        page["page_no"] = idx

    return pages


def parse_and_paginate_word(file_path: Path) -> dict:
    suffix = file_path.suffix.lower()
    if suffix == ".doc":
        return {
            "status": "unsupported",
            "reason": "V1 暂不解析 .doc（二进制老格式），建议转为 .docx 后上传。",
            "pages": [],
        }

    blocks = extract_docx_paragraphs(file_path)
    pages = paginate_blocks(blocks)
    return {
        "status": "ok",
        "reason": "",
        "pages": pages,
        "total_pages": len(pages),
        "total_chars": sum(page["char_count"] for page in pages),
    }


class WordUploadHandler(BaseHTTPRequestHandler):
    def _render_form(self, message: str = "", result_json: str = "") -> bytes:
        escaped_message = html.escape(message)
        escaped_result = html.escape(result_json)
        return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <title>Word 批量上传与分页解析（V1）</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 2rem; }}
    .ok {{ color: #0a7; }}
    pre {{ background: #f6f8fa; padding: 1rem; border-radius: 8px; white-space: pre-wrap; }}
  </style>
</head>
<body>
  <h1>Word 批量上传与分页解析（V1）</h1>
  <p>支持 <strong>.doc / .docx</strong>，单文件最大 10MB，一次最多 50 个文件（可用于 20~50 份验收）。</p>
  <p>上传后会输出分页 JSON（每页 title/content/char_count/page_no）。</p>
  <p class="ok">{escaped_message}</p>
  <form method="post" enctype="multipart/form-data" action="/upload">
    <input type="file" name="files" accept=".doc,.docx" multiple required />
    <button type="submit">上传并解析</button>
  </form>
  <h2>最近一次解析结果</h2>
  <pre>{escaped_result}</pre>
</body>
</html>
""".encode("utf-8")

    def do_GET(self) -> None:
        query = parse_qs(urlparse(self.path).query)
        message = query.get("message", [""])[0]
        result_path = query.get("result", [""])[0]
        result_json = ""

        if result_path:
            candidate = OUTPUT_DIR / sanitize_filename(result_path)
            if candidate.exists():
                result_json = candidate.read_text(encoding="utf-8")

        body = self._render_form(message=message, result_json=result_json)
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self) -> None:
        if self.path != "/upload":
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        ctype = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in ctype:
            self._respond_with_page("上传失败：请使用 multipart/form-data。")
            return

        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            self._respond_with_page("上传失败：请求体为空。")
            return

        body = self.rfile.read(content_length)
        files = self._extract_uploaded_files(body, ctype)

        if not files:
            self._respond_with_page("上传失败：未找到文件字段 files。")
            return

        if len(files) > MAX_FILES:
            self._respond_with_page(f"上传失败：单次最多上传 {MAX_FILES} 个文件。")
            return

        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        results: list[dict] = []

        for filename, file_bytes in files:
            safe_name = sanitize_filename(filename)
            if not safe_name:
                continue

            if len(file_bytes) > MAX_FILE_SIZE:
                results.append(
                    {
                        "file": safe_name,
                        "status": "rejected",
                        "reason": f"文件超过 {MAX_FILE_SIZE} 字节限制",
                    }
                )
                continue

            if not is_allowed_word_file(safe_name):
                results.append(
                    {
                        "file": safe_name,
                        "status": "rejected",
                        "reason": "仅允许 .doc/.docx",
                    }
                )
                continue

            save_path = UPLOAD_DIR / safe_name
            save_path.write_bytes(file_bytes)

            try:
                parsed = parse_and_paginate_word(save_path)
                parsed["file"] = safe_name
                results.append(parsed)
            except Exception as exc:  # noqa: BLE001
                results.append(
                    {
                        "file": safe_name,
                        "status": "error",
                        "reason": f"解析失败: {exc}",
                        "pages": [],
                    }
                )

        output_name = "latest_result.json"
        output_path = OUTPUT_DIR / output_name
        output_path.write_text(
            json.dumps(
                {
                    "total_files": len(results),
                    "results": results,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        self._redirect_with_message(
            f"处理完成：{len(results)} 个文件，结果已写入 {output_path}",
            output_name,
        )

    def _extract_uploaded_files(self, body: bytes, content_type: str) -> list[tuple[str, bytes]]:
        boundary_key = "boundary="
        if boundary_key not in content_type:
            return []

        boundary = content_type.split(boundary_key, 1)[1].strip().strip('"').encode("utf-8")
        delimiter = b"--" + boundary
        parts = body.split(delimiter)
        files: list[tuple[str, bytes]] = []

        for part in parts:
            if b"Content-Disposition" not in part:
                continue
            if b'name="files"' not in part and b'name="file"' not in part:
                continue

            header_end = part.find(b"\r\n\r\n")
            if header_end == -1:
                continue

            headers = part[:header_end]
            payload = part[header_end + 4 :]
            payload = payload.rstrip(b"\r\n")

            marker = b'filename="'
            marker_idx = headers.find(marker)
            if marker_idx == -1:
                continue

            fname_start = marker_idx + len(marker)
            fname_end = headers.find(b'"', fname_start)
            filename = headers[fname_start:fname_end].decode("utf-8", errors="ignore")
            files.append((filename, payload))

        return files

    def _redirect_with_message(self, message: str, result_name: str) -> None:
        location = build_redirect_location(message, result_name)
        self.send_response(HTTPStatus.SEE_OTHER)
        self.send_header("Location", location)
        self.end_headers()

    def _respond_with_page(self, message: str) -> None:
        body = self._render_form(message=message)
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def build_redirect_location(message: str, result_name: str) -> str:
    query = urlencode({"message": message, "result": result_name})
    return f"/?{query}"


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
