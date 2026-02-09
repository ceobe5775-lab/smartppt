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

    # 先根据知识点数量等规则给出初始 layout
    for p in pages:
        p["layout"] = _choose_layout(p, rules)

    # 产品级约束：限制同一 layout 的连续次数、避免连续老师出镜
    enforce_layout_run_limit(pages, max_run=4)
    enforce_no_consecutive_teacher_only(pages)

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
    lines: list[str] = []
    lines.extend(p["bullets"])
    lines.extend(p["quotes"])
    p["content"] = "\n".join(lines).strip()
    p["char_count"] = len(p["content"])
    return p


def _projected_len(p: dict[str, Any]) -> int:
    return len("\n".join(p["bullets"] + p["quotes"]).strip())


def _split_long_text(s: str, max_len: int) -> list[str]:
    """
    极端兜底：单条 bullet/quote 本身就超过 max_len。
    不做“智能改写”，只做硬切片，保证 char_count 不超标。
    """
    s = s.strip()
    if not s:
        return []
    if len(s) <= max_len:
        return [s]
    chunks = []
    i = 0
    while i < len(s):
        chunks.append(s[i : i + max_len])
        i += max_len
    return chunks


def _append_bullet_with_limit(pages: list[dict[str, Any]], cur: dict[str, Any], bullet: str, rules: Rules) -> dict[str, Any]:
    # 如果 bullet 本身超长，先切片
    for piece in _split_long_text(bullet, rules.max_chars_per_page):
        cur["bullets"].append(piece)
        if _projected_len(cur) > rules.max_chars_per_page:
            # 回退这条，先落盘当前页，再开新页放进去
            cur["bullets"].pop()
            pages.append(_finalize_page(cur))
            nxt_title = f"{cur.get('title', '知识点')}（续）"
            cur = _new_page(nxt_title, cur.get("page_type", "bullets"), topic=cur.get("topic", ""), first_signal="char_limit")
            cur["evidence"]["split_reason"].append("char_limit")
            cur["bullets"].append(piece)
    return cur


def _append_quote_with_limit(pages: list[dict[str, Any]], cur: dict[str, Any], quote: str, rules: Rules) -> dict[str, Any]:
    for piece in _split_long_text(quote, rules.max_chars_per_page):
        cur["quotes"].append(piece)
        if _projected_len(cur) > rules.max_chars_per_page:
            cur["quotes"].pop()
            pages.append(_finalize_page(cur))
            cur = _new_page("引用（续）", "quote", topic=cur.get("topic", ""), first_signal="char_limit")
            cur["evidence"]["split_reason"].append("char_limit")
            cur["quotes"].append(piece)
    return cur


def _paginate(blocks: list[str], rules: Rules) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    cur = _new_page("开场", "teacher_only", topic="", first_signal="init")

    for line in blocks:
        text = line.strip()
        if not text:
            continue

        # 章节页：只展示标题，不排版；并切断上下文
        if _is_section_title(text):
            if cur["bullets"] or cur["quotes"] or cur["page_type"] != "teacher_only":
                pages.append(_finalize_page(cur))

            sec = _new_page(text, "section_page", topic=text.split("：", 1)[0], first_signal="section")
            pages.append(_finalize_page(sec))

            cur = _new_page("开场", "teacher_only", topic="", first_signal="after_section")
            continue

        # 引用页：尽量独立
        if _is_quote_line(text):
            if cur["bullets"] and cur["page_type"] != "quote":
                pages.append(_finalize_page(cur))
                cur = _new_page("引用", "quote", topic=cur.get("topic", ""), first_signal="quote_block")

            cur["evidence"]["signals"].append("quote_block")
            cur = _append_quote_with_limit(pages, cur, text, rules)
            continue

        # 老师出镜/寒暄页
        if _looks_teacher_only(text, rules):
            if cur["page_type"] != "teacher_only" and (cur["bullets"] or cur["quotes"]):
                pages.append(_finalize_page(cur))
                cur = _new_page("老师出镜", "teacher_only", topic="", first_signal="teacher_only")

            for b in _split_to_bullets(text):
                cur = _append_bullet_with_limit(pages, cur, b, rules)
            continue

        # 知识点页
        for b in _split_to_bullets(text):
            # 不相关尽量拆页（轻量相似度）
            if rules.topic_split_enabled and cur["bullets"]:
                sim = _avg_similarity_to_page(b, cur["bullets"])
                if sim < rules.similarity_threshold:
                    pages.append(_finalize_page(cur))
                    cur = _new_page("知识点", "bullets", topic=cur.get("topic", ""), first_signal="topic_diverge")
                    cur["evidence"]["split_reason"].append("topic_diverge")

            if cur["page_type"] == "teacher_only":
                # 从老师出镜进入知识点
                pages.append(_finalize_page(cur))
                cur = _new_page("知识点", "bullets", topic="", first_signal="enter_knowledge")

            cur = _append_bullet_with_limit(pages, cur, b, rules)

    if cur["bullets"] or cur["quotes"] or cur["page_type"] in ("section_page", "teacher_only", "quote"):
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


def _allowed_layouts_for_page(page: dict[str, Any]) -> list[str]:
    """
    给每页定义“允许的替代版式”，保证不会胡乱换。
    - bullets>=6：全屏优先，但允许降到小头像
    - bullets 4-5：小头像优先，但允许半屏
    - bullets 1-3：半屏优先，但允许小头像
    - bullets==0：老师出镜（不参与本函数）
    """
    layout = page.get("layout", "")
    pt = page.get("page_type", "")
    if pt in ("teacher_only", "section_page", "title_page"):
        return [layout]

    bullets = len(page.get("bullets", []))
    if bullets >= 6:
        return ["全屏", "小头像"]
    if 4 <= bullets <= 5:
        return ["小头像", "半屏"]
    if 1 <= bullets <= 3:
        return ["半屏", "小头像"]
    return [layout]


def enforce_layout_run_limit(pages: list[dict[str, Any]], max_run: int = 4) -> None:
    """
    产品级约束：
    - 全屏/半屏/小头像 任一 layout 连续不得超过 max_run 次。
    策略：
    - 当某个 layout 已连续达到 max_run，再遇到同 layout：
      尝试把当前页切换到它的“允许替代版式”中的另一个。
    """
    tracked = {"全屏", "半屏", "小头像"}

    run_layout: str | None = None
    run_len = 0

    for p in pages:
        layout = p.get("layout", "")
        pt = p.get("page_type", "")

        # 章节页 / 标题页 / 老师出镜：不计入连续次数
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

        if run_len <= max_run:
            continue

        # 超过 max_run：尝试换版式（在允许集合内换）
        allowed = _allowed_layouts_for_page(p)
        alt = next((x for x in allowed if x != layout), None)

        if alt:
            p["layout"] = alt
            p.setdefault("evidence", {}).setdefault("signals", []).append("layout_run_break")
            p.setdefault("evidence", {}).setdefault("split_reason", []).append(f"layout_run>{max_run}")
            # 换完之后，从新 layout 重新开始计数
            run_layout = alt
            run_len = 1
        else:
            # 没有可替代 layout，保留原样并记录
            p.setdefault("evidence", {}).setdefault("signals", []).append("layout_run_break_failed")


def enforce_no_consecutive_teacher_only(pages: list[dict[str, Any]]) -> None:
    """
    产品级约束：
    - 不允许连续两页 page_type == teacher_only。
    策略：
    - 如果出现连续 teacher_only：
      * 若当前页有内容（bullets/quotes），则降级为知识点页（半屏）
      * 若无内容，则保留老师出镜，但记录信号（极端兜底）
    """
    for i in range(1, len(pages)):
        prev = pages[i - 1]
        cur = pages[i]

        if prev.get("page_type") == "teacher_only" and cur.get("page_type") == "teacher_only":
            has_content = bool(cur.get("bullets") or cur.get("quotes"))
            cur.setdefault("evidence", {}).setdefault("signals", [])
            cur.setdefault("evidence", {}).setdefault("split_reason", [])

            if has_content:
                # 降级为普通知识点页，默认半屏
                cur["page_type"] = "bullets"
                cur["layout"] = "半屏"
                cur["evidence"]["signals"].append("downgrade_from_teacher_only")
                cur["evidence"]["split_reason"].append("no_consecutive_teacher_only")
            else:
                # 极端情况：空老师页，保留但打标
                cur["layout"] = "老师出镜"
                cur["evidence"]["signals"].append("teacher_only_keep_empty")


