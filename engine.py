from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

import yaml


@dataclass(frozen=True)
class Rules:
    version: str
    max_chars_per_page: int

    # Layout thresholds
    full_screen_min: int
    small_avatar_min: int
    small_avatar_max: int
    half_screen_min: int
    half_screen_max: int

    # Labels
    label_title_page: str
    label_section_page: str
    label_teacher_only: str

    # Semantic split
    topic_split_enabled: bool
    similarity_threshold: float

    # Heuristics
    teacher_only_keywords: tuple[str, ...]

    @staticmethod
    def from_dict(rules: dict[str, Any]) -> "Rules":
        engine = rules.get("engine", {})
        layout = rules.get("layout", {})
        topic_split = rules.get("topic_split", {})
        heuristics = rules.get("heuristics", {})

        full = layout.get("full_screen", {})
        small = layout.get("small_avatar", {})
        half = layout.get("half_screen", {})

        return Rules(
            version=str(engine.get("version", "v2")),
            max_chars_per_page=int(engine.get("max_chars_per_page", 150)),
            full_screen_min=int(full.get("min_bullets", 6)),
            small_avatar_min=int(small.get("min_bullets", 4)),
            small_avatar_max=int(small.get("max_bullets", 5)),
            half_screen_min=int(half.get("min_bullets", 1)),
            half_screen_max=int(half.get("max_bullets", 3)),
            label_title_page=str(layout.get("title_page", "标题页")),
            label_section_page=str(layout.get("section_page", "章节页")),
            label_teacher_only=str(layout.get("teacher_only", "老师出镜")),
            topic_split_enabled=bool(topic_split.get("enabled", True)),
            similarity_threshold=float(topic_split.get("similarity_threshold", 0.58)),
            teacher_only_keywords=tuple(heuristics.get("teacher_only_keywords", [])),
        )


SECTION_TITLE_RE = re.compile(r"^(?:[一二三四五六七八九十]+、|\d+[\.、])")
QUOTE_RE = re.compile(r'["“].+["”]')


def load_rules(path: str = "rules.yaml") -> Rules:
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return Rules.from_dict(data)


def paginate_and_classify(text: str, rules_dict: dict[str, Any] | None = None) -> dict[str, Any]:
    """
    高层入口：
    - 接收一段纯文本（通常是 docx 提取后的正文拼接）
    - 根据 rules.yaml 里的规则分页，并给出 layout 建议
    """
    rules = Rules.from_dict(rules_dict) if rules_dict is not None else load_rules()
    blocks = _split_to_blocks(text)
    pages = _paginate(blocks, rules)

    for p in pages:
        p["layout"] = _choose_layout(p, rules)

    stats = {
        "total_pages": len(pages),
        "max_chars_per_page": rules.max_chars_per_page,
        "avg_chars": round(sum(p["char_count"] for p in pages) / len(pages), 2) if pages else 0,
    }
    return {"engine_version": rules.version, "pages": pages, "stats": stats}


# -----------------------------
# 内部工具函数
# -----------------------------
def _split_to_blocks(text: str) -> list[str]:
    lines = [ln.strip() for ln in text.replace("\r\n", "\n").split("\n")]
    return [ln for ln in lines if ln]


def _is_section_title(line: str) -> bool:
    s = line.strip()
    if SECTION_TITLE_RE.match(s):
        return True
    if "：" in s:
        head = s.split("：", 1)[0]
        if 1 <= len(head) <= 20:
            return True
    return False


def _is_quote_line(line: str) -> bool:
    return bool(QUOTE_RE.search(line.strip()))


def _split_to_bullets(line: str) -> list[str]:
    s = line.strip()
    if not s:
        return []
    parts = re.split(r"[；;。]", s)
    bullets = [p.strip() for p in parts if p.strip()]
    if len(bullets) <= 1 and len(s) > 80:
        parts2 = re.split(r"[，,]", s)
        bullets = [p.strip() for p in parts2 if p.strip()]
    return bullets[:3] if bullets else [s]


def _looks_teacher_only(line: str, rules: Rules) -> bool:
    s = line.strip()
    if not s:
        return False
    for kw in rules.teacher_only_keywords:
        if kw and kw in s:
            return True
    if len(s) <= 18 and any(x in s for x in ("下面", "接着", "然后", "接下来")):
        return True
    return False


def _jaccard_similarity(a: str, b: str) -> float:
    def bigrams(s: str) -> set[str]:
        s = re.sub(r"\s+", "", s)
        if len(s) < 2:
            return {s} if s else set()
        return {s[i : i + 2] for i in range(len(s) - 1)}

    A, B = bigrams(a), bigrams(b)
    if not A or not B:
        return 0.0
    return len(A & B) / max(1, len(A | B))


def _avg_similarity_to_page(candidate: str, page_bullets: list[str]) -> float:
    if not page_bullets:
        return 1.0
    sims = [_jaccard_similarity(candidate, b) for b in page_bullets[-3:]]
    return sum(sims) / len(sims)


def _new_page(title: str, page_type: str, topic: str = "", first_signal: str = "") -> dict[str, Any]:
    return {
        "title": title,
        "page_type": page_type,
        "topic": topic,
        "bullets": [],
        "quotes": [],
        "char_count": 0,
        "content": "",
        "evidence": {"signals": [first_signal] if first_signal else [], "split_reason": []},
    }


def _finalize_page(p: dict[str, Any]) -> dict[str, Any]:
    lines = []
    lines.extend(p["bullets"])
    lines.extend(p["quotes"])
    p["content"] = "\n".join(lines).strip()
    p["char_count"] = len(p["content"])
    return p


def _paginate(blocks: list[str], rules: Rules) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    cur = _new_page("开场", "teacher_only", topic="", first_signal="init")

    for line in blocks:
        text = line.strip()
        if not text:
            continue

        if _is_section_title(text):
            if cur["bullets"] or cur["quotes"] or cur["page_type"] != "teacher_only":
                pages.append(_finalize_page(cur))
            cur = _new_page(text, "section_page", topic=text.split("：", 1)[0], first_signal="section")
            pages.append(_finalize_page(cur))
            cur = _new_page("开场", "teacher_only", topic="", first_signal="after_section")
            continue

        if _is_quote_line(text):
            if cur["bullets"]:
                pages.append(_finalize_page(cur))
                cur = _new_page("引用", "quote", topic=cur.get("topic", ""), first_signal="quote_block")
            cur["quotes"].append(text)
            cur["evidence"]["signals"].append("quote_block")
            if cur["char_count"] > rules.max_chars_per_page:
                pages.append(_finalize_page(cur))
                cur = _new_page("引用（续）", "quote", topic=cur.get("topic", ""), first_signal="char_limit")
            continue

        if _looks_teacher_only(text, rules):
            if cur["page_type"] != "teacher_only" and (cur["bullets"] or cur["quotes"]):
                pages.append(_finalize_page(cur))
                cur = _new_page("老师出镜", "teacher_only", topic="", first_signal="teacher_only")
            for b in _split_to_bullets(text):
                cur["bullets"].append(b)
                if cur["char_count"] > rules.max_chars_per_page:
                    pages.append(_finalize_page(cur))
                    cur = _new_page("老师出镜（续）", "teacher_only", topic="", first_signal="char_limit")
            continue

        for b in _split_to_bullets(text):
            if rules.topic_split_enabled and cur["bullets"]:
                sim = _avg_similarity_to_page(b, cur["bullets"])
                if sim < rules.similarity_threshold:
                    pages.append(_finalize_page(cur))
                    cur = _new_page("知识点", "bullets", topic=cur.get("topic", ""), first_signal="topic_diverge")
                    cur["evidence"]["split_reason"].append("topic_diverge")

            cur["bullets"].append(b)
            if cur["char_count"] > rules.max_chars_per_page:
                cur["bullets"].pop()
                pages.append(_finalize_page(cur))
                cur = _new_page("知识点（续）", "bullets", topic=cur.get("topic", ""), first_signal="char_limit")
                cur["evidence"]["split_reason"].append("char_limit")
                cur["bullets"].append(b)

    if cur["bullets"] or cur["quotes"]:
        pages.append(_finalize_page(cur))

    for i, p in enumerate(pages, start=1):
        p["page_no"] = i

    return pages


def _choose_layout(page: dict[str, Any], rules: Rules) -> str:
    pt = page.get("page_type", "")
    bullet_count = len(page.get("bullets", []))

    if pt == "section_page":
        return rules.label_section_page
    if pt == "title_page":
        return rules.label_title_page
    if pt == "teacher_only":
        return rules.label_teacher_only

    if bullet_count == 0:
        return rules.label_teacher_only

    if bullet_count >= rules.full_screen_min:
        return "全屏"
    if rules.small_avatar_min <= bullet_count <= rules.small_avatar_max:
        return "小头像"
    if rules.half_screen_min <= bullet_count <= rules.half_screen_max:
        return "半屏"

    return "半屏"


