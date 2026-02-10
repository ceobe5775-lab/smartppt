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

from engine import load_rules, paginate_and_classify


# -----------------------------
# Constants & basic config
# -----------------------------
UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB per file
MAX_FILES = 50

DOCX_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
ENGINE_VERSION = os.getenv("ENGINE_VERSION", "v2-knowledge-point")
BUILD_TIME = os.getenv("BUILD_TIME", datetime.now(timezone.utc).isoformat(timespec="seconds"))


@dataclass(frozen=True)
class EngineConfig:
    """规则化的分页配置（当前仍然是硬编码，后续可接 rules.yaml）."""

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


# -----------------------------
# Metadata helpers
# -----------------------------
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


# -----------------------------
# Basic validators / signals
# -----------------------------
def is_allowed_word_file(filename: str) -> bool:
    """Return True when filename ends with .doc or .docx (case-insensitive)."""
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
    """将一段长句拆成若干 bullet，尽量控制在 max_bullet_chars 以内。"""
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


# -----------------------------
# DOCX parsing
# -----------------------------
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


# -----------------------------
# Pagination core
# -----------------------------
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


# -----------------------------
# Product-level hard constraints
# -----------------------------
def prune_empty_pages(pages: list[dict]) -> list[dict]:
    """
    产品级硬过滤器（必须写死）：
    除 标题页 / 章节页 外：
    如果 bullets + quotes + content 为空 -> 整页直接删除，不允许出现。
    """
    kept: list[dict] = []
    for p in pages:
        if p.get("layout") in ("标题页", "章节页"):
            kept.append(p)
            continue

        has_text = bool(
            p.get("bullets")
            or p.get("quotes")
            or (p.get("content") and str(p.get("content")).strip())
        )
        if not has_text:
            continue  # ❌ 直接丢掉
        kept.append(p)
    return kept


def renumber_page_no(pages: list[dict]) -> list[dict]:
    for i, p in enumerate(pages, start=1):
        p["page_no"] = i
    return pages


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

    # 先用旧的提取逻辑把段落抽出来
    blocks = extract_docx_paragraphs(file_path)
    # 拼成纯文本给 engine 统一分页（走 rules.yaml + 150 字规则）
    plain_text = "\n".join(b.text for b in blocks)

    engine_result = paginate_and_classify(plain_text, None)
    pages = engine_result["pages"]

    # 为了兼容现有质量评分和统计逻辑，这里补充 quality_score 字段
    for page in pages:
        # 如果后面需要更细的评分，可以在 score_page 内部再调参数
        page["quality_score"] = score_page(
            {
                "char_count": page.get("char_count", 0),
                "bullets": page.get("bullets", []),
                "quotes": page.get("quotes", []),
                "page_type": page.get("page_type", ""),
                "topic": page.get("topic", ""),
            }
        )

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


def compute_health_for_pages(pages: list[dict]) -> dict:
    """
    基于单次解析得到的 pages 做健康度检查。
    """
    rules = load_rules()

    all_under_limit = all(p.get("char_count", 0) <= rules.max_chars_per_page for p in pages)
    all_have_layout = all(p.get("layout") for p in pages)

    # 老师出镜不连续
    no_consecutive_teacher_only = True
    for i in range(1, len(pages)):
        if pages[i - 1].get("page_type") == "teacher_only" and pages[i].get("page_type") == "teacher_only":
            no_consecutive_teacher_only = False
            break

    # layout 连续不超过 4（仅统计全屏/半屏/小头像，老师出镜 & 章节页不计入）
    tracked = {"全屏", "半屏", "小头像"}
    no_layout_run_over_4 = True
    run_layout: str | None = None
    run_len = 0
    for p in pages:
        layout = p.get("layout", "")
        pt = p.get("page_type", "")

        if pt in ("teacher_only", "section_page", "title_page"):
            run_layout = None
            run_len = 0
            continue

        if layout not in tracked:
            run_layout = None
            run_len = 0
            continue

        if layout == run_layout:
            run_len += 1
        else:
            run_layout = layout
            run_len = 1

        if run_len > 4:
            no_layout_run_over_4 = False
            break

    # 章节页是否只展示标题（不带 bullets/quotes）
    section_pages = [p for p in pages if p.get("page_type") in ("section_page", "section_cover")]
    section_title_only = all(
        not p.get("bullets") and not p.get("quotes") for p in section_pages
    ) if section_pages else True

    return {
        "all_pages_under_150": all_under_limit,
        "all_pages_have_layout": all_have_layout,
        "no_consecutive_teacher_only": no_consecutive_teacher_only,
        "no_layout_run_over_4": no_layout_run_over_4,
        "section_pages_title_only": section_title_only,
    }


def build_product_snapshot(metadata: dict[str, str], latest_result_path: Path) -> dict:
    """
    产品状态快照，用于快速确认当前“引擎 + 规则 + 输出健康度”。
    """
    rules = load_rules()

    # 默认健康检查结果（如果 latest_result.json 不存在，就认为尚未通过验证）
    health = {
        "all_pages_under_150": False,
        "all_pages_have_layout": False,
        "no_consecutive_teacher_only": False,
        "no_layout_run_over_4": False,
        "section_pages_title_only": False,
    }

    if latest_result_path.exists():
        try:
            data = json.loads(latest_result_path.read_text(encoding="utf-8"))
            results = data.get("results", [])
            pages: list[dict] = []
            for item in results:
                pages.extend(item.get("pages", []))

            if pages:
                health = compute_health_for_pages(pages)
        except Exception:
            # latest_result.json 结构异常时，保持默认值，方便排查
            pass

    snapshot = {
        "engine_version": rules.version,
        "git_commit": metadata.get("git_sha", ""),
        "rules": {
            "max_chars_per_page": rules.max_chars_per_page,
            "layout_mapping": {
                f"{rules.full_screen_min}+": "全屏",
                f"{rules.small_avatar_min}-{rules.small_avatar_max}": "小头像",
                f"{rules.half_screen_min}-{rules.half_screen_max}": "半屏",
                "0": "老师出镜",
            },
        },
        "latest_result": str(latest_result_path.as_posix()),
        "healthcheck": health,
    }
    return snapshot


def build_preview_html(results: list[dict]) -> str:
    """
    基于本次解析结果构建一个简单的 HTML 预览，方便肉眼检查。
    """
    lines: list[str] = [
        "<!doctype html>",
        "<html lang='zh-CN'>",
        "<head>",
        "  <meta charset='utf-8' />",
        "  <title>SmartPPT 预览</title>",
        "  <style>",
        "    body { font-family: Arial, sans-serif; margin: 2rem; }",
        "    h1 { margin-bottom: 1.5rem; }",
        "    .file { margin-bottom: 2rem; }",
        "    .page { border: 1px solid #ddd; padding: 0.75rem 1rem; margin-bottom: 0.5rem; border-radius: 6px; }",
        "    .layout { font-weight: bold; color: #555; }",
        "    .meta { color: #888; font-size: 0.9rem; }",
        "  </style>",
        "</head>",
        "<body>",
        "  <h1>SmartPPT 分页预览</h1>",
    ]

    for item in results:
        fname = item.get("file", "<unknown>")
        pages = item.get("pages", [])
        health = item.get("healthcheck", {})

        lines.append("  <div class='file'>")
        lines.append(f"    <h2>文件：{html.escape(fname)}</h2>")
        if health:
            lines.append("    <p class='meta'>Healthcheck：")
            flags = [f"{k}={v}" for k, v in health.items()]
            lines.append(" | ".join(flags) + "</p>")

        for p in pages:
            page_no = p.get("page_no", "?")
            layout = p.get("layout", "")
            page_type = p.get("page_type", "")
            char_count = p.get("char_count", 0)
            bullets = p.get("bullets", [])
            quotes = p.get("quotes", [])
            title = p.get("title", "")

            lines.append("    <div class='page'>")
            lines.append(
                f"      <div class='layout'>[第 {page_no} 页] "
                f"{html.escape(layout)} "
                f"<span class='meta'>({html.escape(page_type)}, {char_count} 字)</span></div>"
            )
            if title and page_type in ("section_page", "title_page"):
                lines.append(f"      <p>标题：{html.escape(title)}</p>")

            if bullets:
                lines.append("      <ul>")
                for b in bullets:
                    lines.append(f"        <li>{html.escape(str(b))}</li>")
                lines.append("      </ul>")

            if quotes:
                lines.append("      <blockquote>")
                for q in quotes:
                    lines.append(f"        <p>{html.escape(str(q))}</p>")
                lines.append("      </blockquote>")

            lines.append("    </div>")

        lines.append("  </div>")

    lines.append("</body>")
    lines.append("</html>")
    return "\n".join(lines)


# -----------------------------
# HTTP server
# -----------------------------
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
  <p class="actions">
    <a href="/download?file=latest_result.json">下载 latest_result.json</a>
    <a href="/download?file=latest_report.txt">下载 latest_report.txt</a>
    <a href="/download?file=preview.html">查看分页预览（HTML）</a>
  </p>
  <h2>最近一次解析结果（JSON）</h2>
  <pre>{escaped_result}</pre>
</body>
</html>
""".encode("utf-8")

    # ---- GET / ----
    def do_GET(self) -> None:  # noqa: N802
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

    # ---- POST /upload ----
    def do_POST(self) -> None:  # noqa: N802
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

                # 产品级硬过滤：在最终输出前移除“没字但占一页”的非法空页，并重新编号 page_no
                pages = prune_empty_pages(parsed.get("pages", []))
                renumber_page_no(pages)
                parsed["pages"] = pages
                parsed["total_pages"] = len(pages)
                parsed["total_chars"] = sum(p.get("char_count", 0) for p in pages)
                parsed["avg_score"] = round(
                    (sum(float(p.get("quality_score", 0)) for p in pages) / len(pages)), 2
                ) if pages else 0

                results.append(parsed)
            except Exception as exc:  # noqa: BLE001
                results.append({"file": safe_name, "status": "error", "reason": f"解析失败: {exc}", "pages": []})

        metadata = build_metadata()
        # 为每个文件结果补充 stats / healthcheck，方便下游直接使用
        for item in results:
            pages = item.get("pages", [])
            item.setdefault("stats", {})
            engine_health = compute_health_for_pages(pages) if pages else {}
            item["healthcheck"] = engine_health

        payload = {"metadata": metadata, "total_files": len(results), "results": results}

        output_name = "latest_result.json"
        output_path = OUTPUT_DIR / output_name
        output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

        report_path = OUTPUT_DIR / "latest_report.txt"
        report_path.write_text(build_report(results, metadata), encoding="utf-8")

        # 生成/更新产品状态快照（放在项目根目录）
        snapshot = build_product_snapshot(metadata, output_path)
        snapshot_path = Path("product_snapshot.json")
        snapshot_path.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8")

        # 生成预览 HTML，方便肉眼检查分页与版式
        preview_html = build_preview_html(results)
        preview_path = OUTPUT_DIR / "preview.html"
        preview_path.write_text(preview_html, encoding="utf-8")

        self._redirect_with_message(f"处理完成：{len(results)} 个文件，结果已写入 {output_path}", output_name)

    # ---- helpers ----
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
        elif safe_name.endswith(".html"):
            ctype = "text/html; charset=utf-8"
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


# -----------------------------
# URL helpers & CLI
# -----------------------------
def build_redirect_location(message: str, result_name: str) -> str:
    # 通过 urlencode 确保 Location 可被 latin-1 安全编码（单测有覆盖）
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
    parser = argparse.ArgumentParser(description="Run Word knowledge-point pagination demo server.")
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

