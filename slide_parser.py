"""Utilities for parsing lecture slide directives from .docx files."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List

from docx import Document

# Matches paragraphs like: "P3老师出镜", "P7 全屏", "p12小头像"
_SLIDE_HEADER_RE = re.compile(r"^\s*[Pp]\s*(\d+)\s*([\u4e00-\u9fffA-Za-z0-9_\-（）()【】\[\]·、，,。：:；;！？!?《》\s]+?)\s*$")


def parse_docx_slides(docx_path: str | Path) -> List[Dict[str, object]]:
    """Parse a .docx lecture script into a structured list of slide dictionaries.

    A new slide starts when a paragraph begins with the pattern "P<page><layout>",
    for example "P3老师出镜" or "P7 全屏".

    Non-header paragraphs are appended to the current slide's content. Blank
    paragraphs are ignored.
    """
    document = Document(str(docx_path))

    slides: List[Dict[str, object]] = []
    current_slide: Dict[str, object] | None = None
    content_lines: List[str] = []

    def flush_current_slide() -> None:
        nonlocal current_slide, content_lines
        if current_slide is None:
            return
        current_slide["content"] = "\n".join(content_lines).strip()
        slides.append(current_slide)
        current_slide = None
        content_lines = []

    for paragraph in document.paragraphs:
        raw_text = paragraph.text or ""
        text = raw_text.strip()
        if not text:
            continue

        header_match = _SLIDE_HEADER_RE.match(text)
        if header_match:
            flush_current_slide()
            page_num = int(header_match.group(1))
            layout = header_match.group(2).strip()
            current_slide = {
                "page_num": page_num,
                "layout": layout,
                "content": "",
            }
            continue

        if current_slide is not None:
            content_lines.append(text)

    flush_current_slide()
    return slides
