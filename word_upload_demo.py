from __future__ import annotations

import argparse
import html
import json
import os
import re
import subprocess
import threading
import webbrowser
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse
from xml.etree import ElementTree as ET

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB per file
MAX_FILES = 50

DOCX_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
ENGINE_VERSION = os.getenv("ENGINE_VERSION", "v2-knowledge-point")
BUILD_TIME = os.getenv("BUILD_TIME", datetime.now(timezone.utc).isoformat(timespec="seconds"))


@dataclass(frozen=True)
class EngineConfig:
    max_chars_per_page: int = 260
    max_bullets_per_page: int = 6
    short_title_char_limit: int = 20
    max_bullet_chars: int = 90


ENGINE_CONFIG = EngineConfig()
SECTION_TITLE_RE = re.compile(r"^(?:[一二三四五六七八九十]+、|\d+[\.、])")
PERSON_START_RE = re.compile(r"^([\u4e00-\u9fff]{2,3})(?:作为|是|则是)")
QUOTE_RE = re.compile(r'["“].+["”]')


@dataclass
class ParagraphBlock:
    text: str
    is_heading: bool = False



def get_git_sha() -> str:
    try:
        return (
            subprocess.check_output(["git", "rev-parse", "--short", "HEAD"], stderr=subprocess.DEVNULL)
            .decode("utf-8")
            .strip()
        )
    except Exception:  # noqa: BLE001
        return "unknown"


def build_metadata() -> dict[str, str]:
    return {
        "engine_version": ENGINE_VERSION,
        "git_sha": get_git_sha(),
        "build_time": BUILD_TIME,
    }


def is_allowed_word_file(filename: str) -> bool:
    lowered = filename.lower()
    return lowered.endswith(".doc") or lowered.endswith(".docx")


def sanitize_filename(filename: str) -> str:
    return os.path.basename(filename).replace("..", "").strip()


def is_section_title(text: str) -> bool:
    text = text.strip()
    if SECTION_TITLE_RE.match(text):
        return True
    if "：" in text and len(text.split("：", 1)[0]) <= ENGINE_CONFIG.short_title_char_limit:
        return True
    return False


def detect_person_topic(text: str) -> str:
    match = PERSON_START_RE.match(text.strip())
    if match:
        return match.group(1)
    return ""


def is_quote_line(text: str) -> bool:
    stripped = text.strip()
    return bool(QUOTE_RE.search(stripped))


def split_to_bullets(text: str, config: EngineConfig = ENGINE_CONFIG) -> list[str]:
    text = text.strip()
    if not text:
        return []
    chunks = re.split(r"[。！？!?；;]", text)
    primary = [chunk.strip(" ，、\n\t") for chunk in chunks if chunk.strip()]

    bullets: list[str] = []
    for chunk in primary:
        if len(chunk) <= config.max_bullet_chars:
            bullets.append(chunk)
            continue

        secondary = [part.strip(" ，、\n\t") for part in re.split(r"[，、]", chunk) if part.strip()]
        current = ""
        for part in secondary:
            candidate = f"{current}，{part}" if current else part
            if len(candidate) > config.max_bullet_chars and current:
                bullets.append(current)
                current = part
            else:
                current = candidate
        if current:
            bullets.append(current)

    return bullets if bullets else [text]


def extract_docx_paragraphs(file_path: Path) -> list[ParagraphBlock]:
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

        blocks.append(ParagraphBlock(text=text, is_heading=style_val.lower().startswith("heading")))

    return blocks


def init_page(title: str, topic: str, page_type: str, first_signal: str, chunk_id: int) -> dict:
    return {
        "title": title,
        "topic": topic,
        "page_type": page_type,
        "bullets": [],
        "quotes": [],
        "content": "",
        "char_count": 0,
        "evidence": {"signals": [first_signal] if first_signal else [], "source_chunks": [chunk_id]},
    }


def finalize_page(page: dict) -> dict:
    page["content"] = "\n".join(page["bullets"] + page["quotes"]).strip()
    page["char_count"] = len(page["content"])
    page["quality_score"] = score_page(page)
    return page


def score_page(page: dict) -> int:
    score = 100
    if page["char_count"] > ENGINE_CONFIG.max_chars_per_page:
        score -= 25
    if not (1 <= len(page["bullets"]) <= ENGINE_CONFIG.max_bullets_per_page):
        score -= 15
    if page["page_type"] == "quote" and not page["quotes"]:
        score -= 20
    if page["page_type"] == "person_profile" and not page["topic"]:
        score -= 10
    return max(0, score)


def paginate_blocks(blocks: list[ParagraphBlock], config: EngineConfig = ENGINE_CONFIG) -> list[dict]:
    pages: list[dict] = []
    current = init_page("课程导入", "总览", "bullets", "init", 0)
    current_person = ""

    def flush() -> None:
        nonlocal current
        if not current["bullets"] and not current["quotes"]:
            return
        pages.append(finalize_page(current))

    for idx, block in enumerate(blocks, start=1):
        text = block.text.strip()
        if not text:
            continue

        person = detect_person_topic(text)
        section_hit = block.is_heading or is_section_title(text)
        quote_hit = is_quote_line(text)

        if section_hit:
            flush()
            current = init_page(text, text.split("：", 1)[0], "section_cover", "section", idx)
            current_person = ""
            if "：" in text:
                title, subtitle = text.split("：", 1)
                current["title"] = title.strip()
                if subtitle.strip():
                    current["bullets"].append(subtitle.strip())
            continue

        if person and person != current_person and (current["bullets"] or current["quotes"]):
            flush()
            current = init_page(f"{person}：核心知识点", person, "person_profile", "person_switch", idx)
            current_person = person
        elif person and not current_person:
            current["topic"] = person
            current["page_type"] = "person_profile"
            current["evidence"]["signals"].append("person_switch")
            current_person = person

        if quote_hit:
            if current["bullets"] and current["page_type"] != "quote":
                flush()
                quote_title = f"{current['topic']}：代表诗句" if current["topic"] else "代表诗句"
                current = init_page(quote_title, current.get("topic", ""), "quote", "quote_block", idx)
            current["quotes"].append(text)
            current["evidence"]["signals"].append("quote_block")
            current["evidence"]["source_chunks"].append(idx)
            continue

        for bullet in split_to_bullets(text, config):
            projected = "\n".join(current["bullets"] + [bullet] + current["quotes"])
            if (
                current["bullets"]
                and (len(projected) > config.max_chars_per_page or len(current["bullets"]) >= config.max_bullets_per_page)
            ):
                flush()
                follow_title = f"{current['topic']}（续）" if current["topic"] else "知识点续页"
                current = init_page(follow_title, current.get("topic", ""), "bullets", "length", idx)
            current["bullets"].append(bullet)
            current["evidence"]["source_chunks"].append(idx)

    flush()

    if not pages:
        pages = [finalize_page(init_page("空文档", "", "summary", "empty", 0))]

    for i, page in enumerate(pages, start=1):
        page["page_no"] = i

    return pages


def parse_and_paginate_word(file_path: Path) -> dict:
    suffix = file_path.suffix.lower()
    if suffix == ".doc":
        return {
            "status": "unsupported",
            "reason": "V2 暂不解析 .doc（二进制老格式），建议转为 .docx 后上传。",
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
        "avg_score": round(sum(page["quality_score"] for page in pages) / len(pages), 2) if pages else 0,
    }


def build_report(results: list[dict], metadata: dict[str, str]) -> str:
    status_counts: dict[str, int] = {}
    total_pages = 0
    total_avg_score = 0.0
    scored_files = 0

    for item in results:
        status = item.get("status", "unknown")
        status_counts[status] = status_counts.get(status, 0) + 1
        total_pages += item.get("total_pages", 0)
        if "avg_score" in item:
            total_avg_score += float(item["avg_score"])
            scored_files += 1

    avg_score = round(total_avg_score / scored_files, 2) if scored_files else 0
    lines = [
        f"engine_version: {metadata['engine_version']}",
        f"git_sha: {metadata['git_sha']}",
        f"build_time: {metadata['build_time']}",
        f"total_files: {len(results)}",
        f"status_counts: {status_counts}",
        f"total_pages: {total_pages}",
        f"avg_score: {avg_score}",
    ]
    return "\n".join(lines) + "\n"


class WordUploadHandler(BaseHTTPRequestHandler):
    def _render_form(self, message: str = "", result_json: str = "") -> bytes:
        escaped_message = html.escape(message)
        escaped_result = html.escape(result_json)
        metadata = build_metadata()
        return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <title>Word 知识点分页引擎（V2）</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 2rem; }}
    .ok {{ color: #0a7; }}
    .meta {{ color: #555; font-size: 0.95rem; }}
    .actions a {{ margin-right: 1rem; }}
    pre {{ background: #f6f8fa; padding: 1rem; border-radius: 8px; white-space: pre-wrap; }}
  </style>
</head>
<body>
  <h1>Word 知识点分页引擎（V2）</h1>
  <p class="meta">Engine: {html.escape(metadata['engine_version'])} | Git SHA: {html.escape(metadata['git_sha'])} | Build: {html.escape(metadata['build_time'])}</p>
  <p>支持 .doc/.docx，单次最多 50 份。输出包含 page_type/topic/bullets/quotes/evidence。</p>
  <p class="ok">{escaped_message}</p>
  <form method="post" enctype="multipart/form-data" action="/upload">
    <input type="file" name="files" accept=".doc,.docx" multiple required />
    <button type="submit">上传并分页</button>
  </form>
  <p class="actions"><a href="/download?file=latest_result.json">下载 latest_result.json</a><a href="/download?file=latest_report.txt">下载 latest_report.txt</a></p>
  <h2>最近一次解析结果（JSON）</h2>
  <pre>{escaped_result}</pre>
</body>
</html>
""".encode("utf-8")

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/download":
            self._download_output(parse_qs(parsed.query).get("file", [""])[0])
            return

        query = parse_qs(parsed.query)
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
                results.append({"file": safe_name, "status": "rejected", "reason": f"文件超过 {MAX_FILE_SIZE} 字节限制"})
                continue

            if not is_allowed_word_file(safe_name):
                results.append({"file": safe_name, "status": "rejected", "reason": "仅允许 .doc/.docx"})
                continue

            save_path = UPLOAD_DIR / safe_name
            save_path.write_bytes(file_bytes)

            try:
                parsed = parse_and_paginate_word(save_path)
                parsed["file"] = safe_name
                results.append(parsed)
            except Exception as exc:  # noqa: BLE001
                results.append({"file": safe_name, "status": "error", "reason": f"解析失败: {exc}", "pages": []})

        metadata = build_metadata()
        payload = {"metadata": metadata, "total_files": len(results), "results": results}

        output_name = "latest_result.json"
        output_path = OUTPUT_DIR / output_name
        output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

        report_path = OUTPUT_DIR / "latest_report.txt"
        report_path.write_text(build_report(results, metadata), encoding="utf-8")

        self._redirect_with_message(f"处理完成：{len(results)} 个文件，结果已写入 {output_path}", output_name)

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
            payload = part[header_end + 4 :].rstrip(b"\r\n")
            marker = b'filename="'
            marker_idx = headers.find(marker)
            if marker_idx == -1:
                continue

            fname_start = marker_idx + len(marker)
            fname_end = headers.find(b'"', fname_start)
            filename = headers[fname_start:fname_end].decode("utf-8", errors="ignore")
            files.append((filename, payload))

        return files

    def _download_output(self, filename: str) -> None:
        safe_name = sanitize_filename(filename)
        if not safe_name:
            self.send_error(HTTPStatus.BAD_REQUEST, "缺少文件名")
            return

        candidate = OUTPUT_DIR / safe_name
        if not candidate.exists() or not candidate.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, "文件不存在")
            return

        content = candidate.read_bytes()
        if safe_name.endswith(".json"):
            ctype = "application/json; charset=utf-8"
        else:
            ctype = "text/plain; charset=utf-8"

        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Content-Disposition", f'attachment; filename="{safe_name}"')
        self.end_headers()
        self.wfile.write(content)

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
    parser.add_argument("--open-browser", action="store_true", help="Automatically open default browser after server starts.")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    run_server(host=args.host, port=args.port, open_browser=args.open_browser)
