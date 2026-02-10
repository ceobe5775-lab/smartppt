from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

import yaml

from ai_hooks import Intent, safe_ai_classify


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

    # Importance scoring
    keep_keywords: tuple[str, ...]
    drop_keywords: tuple[str, ...]
    anchor_patterns: tuple[str, ...]
    knowledge_min_score: int
    teacher_only_max_score: int

    @staticmethod
    def from_dict(rules: dict[str, Any]) -> "Rules":
        engine = rules.get("engine", {})
        layout = rules.get("layout", {})
        topic_split = rules.get("topic_split", {})
        heuristics = rules.get("heuristics", {})
        importance = rules.get("importance", {})
        thresholds = rules.get("thresholds", {})

        full = layout.get("full_screen", {})
        small = layout.get("small_avatar", {})
        half = layout.get("half_screen", {})

        return Rules(
            version=str(engine.get("version", "v2")),
            max_chars_per_page=int(engine.get("max_chars_per_page", 180)),
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
            keep_keywords=tuple(importance.get("keep_keywords", [])),
            drop_keywords=tuple(importance.get("drop_keywords", [])),
            anchor_patterns=tuple(importance.get("anchor_patterns", [])),
            knowledge_min_score=int(thresholds.get("knowledge_min_score", 2)),
            teacher_only_max_score=int(thresholds.get("teacher_only_max_score", 0)),
        )


# -----------------------------
# Layout 枚举（产品级接口约束）
# -----------------------------
ALLOWED_LAYOUTS = {
    "全屏",
    "老师出镜",
    "半屏",
    "小头像",
    "标题页",
    "章节页",
}

SECTION_TITLE_RE = re.compile(r"^(?:[一二三四五六七八九十]+、|\d+[\.、])")
QUOTE_RE = re.compile(r'[""].+[""]')
PUNCT_FOR_SPLIT = "。！？；，,"

# 标签解析正则
TAG_RE = re.compile(r"^【(?P<tag>标题页|章节页|老师出镜|要点|例子|引用|可略)】\s*")

# -----------------------------
# 主知识点锚点（产品级硬规则：知识点块 > 页面）
# -----------------------------
# 目标：一个页面最多 1 个“主知识点”
# 说明：不依赖模型，仅靠强规则；命中即认为“开启新知识点块”
KNOWLEDGE_ANCHOR_PATTERNS: tuple[str, ...] = (
    r"^[\u4e00-\u9fff]{2,4}是",  # 人物/概念 + 判断句
    r"^[\u4e00-\u9fff]{2,4}的",  # 人物/概念 + 的...
    r"^[\u4e00-\u9fff]{2,4}则",  # X则...
    r"^南北朝时期",
    r"^这一时期",
    r"^分为",
    r"^主要有",
    r"^代表作",
    r"^特点是",
    r"^贡献在于",
)


def is_main_knowledge_anchor(text: str) -> bool:
    s = (text or "").strip()
    if not s:
        return False
    return any(re.search(p, s) for p in KNOWLEDGE_ANCHOR_PATTERNS)


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

    # 产品级约束：跨页凝聚力修正（把引子句从上一页末尾搬到下一页开头）
    enforce_topic_cohesion(pages, rules.max_chars_per_page)

    # 搬运后可能需要重新计算 layout（如果 bullets 数量变了）
    for p in pages:
        p["layout"] = _choose_layout(p, rules)

    # 产品级约束：统一校验 layout 枚举、清理标题页/章节页结构、添加 page_tag
    for p in pages:
        p["layout"] = enforce_layout(p["layout"])
        enforce_page_structure(p)
        p["page_tag"] = f"【P{p['page_no']}{p['layout']}】"

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


def parse_tag(line: str) -> tuple[str | None, str]:
    """
    解析行首标签，返回 (tag, clean_text)。
    支持的标签：标题页、章节页、老师出镜、要点、例子、引用、可略
    """
    m = TAG_RE.match(line.strip())
    if not m:
        return None, line.strip()
    tag = m.group("tag")
    rest = line[m.end() :].strip()
    return tag, rest


def score_line(text: str, keep_keywords: tuple[str, ...], drop_keywords: tuple[str, ...]) -> int:
    """
    词表打分（只在无标签时使用）：
    - 命中 keep_keywords → +1
    - 命中 drop_keywords → -1
    """
    s = 0
    for k in keep_keywords:
        if k and k in text:
            s += 1
    for k in drop_keywords:
        if k and k in text:
            s -= 1
    return s


def _matches_anchor_pattern(text: str, patterns: tuple[str, ...]) -> bool:
    """检查文本是否匹配知识点锚点模式（触发新知识点块）。"""
    for pattern in patterns:
        if re.match(pattern, text.strip()):
            return True
    return False


def classify_block(
    tag: str | None,
    text: str,
    score: int,
    rules: Rules,
) -> tuple[str, bool]:
    """
    根据标签/词表打分决定 block_type 和是否强制新知识点块。
    返回 (block_type, force_new_topic)
    """
    # 标签优先
    if tag == "标题页":
        return "title", True
    if tag == "章节页":
        return "section", True
    if tag == "老师出镜":
        return "teacher_only", False
    if tag == "要点":
        return "knowledge", True  # 新要点强制新页
    if tag == "引用":
        return "quote", False
    if tag == "例子":
        return "example", False
    if tag == "可略":
        return "drop", False

    # 无标签：走词表打分 + 锚点检测
    force_new = _matches_anchor_pattern(text, rules.anchor_patterns)

    if score <= rules.teacher_only_max_score:
        return "teacher_only", force_new
    if score >= rules.knowledge_min_score:
        return "knowledge", force_new

    return "teacher_only", force_new


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
    """
    仅当整行主要是“引用句”时才认为是 quote：
    - 以引号开头、以引号结尾
    - 总长度不要太长（避免整段叙述被当成引用）
    """
    s = line.strip()
    if not s:
        return False
    # 保留一个较宽松的正则备选（将来如果需要更复杂模式）
    if QUOTE_RE.fullmatch(s):
        return True
    return (s.startswith(("“", '"')) and s.endswith(("”", '"')) and len(s) <= 120)


def _split_to_bullets(line: str) -> list[str]:
    s = line.strip()
    if not s:
        return []

    # 如果包含成对的中文引号，整行必须作为一个 bullet，不在引号内拆分。
    # 即使引号内内容很长，也交给后续的 _split_long_text 处理（那里有引号保护逻辑）。
    if "“" in s and "”" in s:
        first_q = s.find("“")
        last_q = s.rfind("”")
        if 0 <= first_q < last_q:
            # 只要引号成对存在，就整行保留为一个 bullet
            return [s]

    parts = re.split(r"[；;。]", s)
    bullets = [p.strip() for p in parts if p.strip()]
    if len(bullets) <= 1 and len(s) > 80:
        parts2 = re.split(r"[，,]", s)
        bullets = [p.strip() for p in parts2 if p.strip()]
    # 单行最多拆成 5 个要点，提升版式多样性
    return bullets[:5] if bullets else [s]


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
        "items": [],  # [{text, intent}]
        "char_count": 0,
        "content": "",
        "_has_main_anchor": False,  # internal flag: enforce at most 1 main knowledge anchor per page
        "evidence": {"signals": [first_signal] if first_signal else [], "split_reason": []},
    }


def _finalize_page(p: dict[str, Any]) -> dict[str, Any]:
    lines: list[str] = []
    lines.extend(p["bullets"])
    lines.extend(p["quotes"])
    p["content"] = "\n".join(lines).strip()
    p["char_count"] = len(p["content"])
    # intent_mix: 该页由哪些展示意图构成（SHOW / SUPPORT / SAY）
    intents = sorted({it.get("intent") for it in p.get("items", []) if it.get("intent")})
    if intents:
        p["intent_mix"] = intents
    # internal flags should not leak to downstream payload
    p.pop("_has_main_anchor", None)
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

    chunks: list[str] = []
    while len(s) > max_len:
        # 默认切在 max_len
        cut = max_len
        window = s[:max_len]

        # 在当前窗口内向前找一个“更自然”的切分点（标点）
        for j in range(max_len - 1, -1, -1):
            if window[j] in PUNCT_FOR_SPLIT:
                cut = j + 1
                break

        # 如果 window 中引号不平衡，说明切点在引号内部，优先在闭引号之后拆分
        open_q = window.count("“")
        close_q = window.count("”")
        if open_q > close_q:
            # 优先策略：在整个字符串中找最近的闭引号，在闭引号之后拆分
            full_close_q = s.find("”", max_len - 1)
            if full_close_q != -1 and full_close_q < len(s):
                # 在闭引号之后找标点
                after_close = s[full_close_q + 1 : full_close_q + 50]
                punct_pos = -1
                for i, char in enumerate(after_close):
                    if char in PUNCT_FOR_SPLIT:
                        punct_pos = full_close_q + 1 + i + 1
                        break
                if punct_pos != -1:
                    cut = punct_pos
                else:
                    # 如果闭引号后没有标点，就在闭引号之后直接切
                    cut = full_close_q + 1
            else:
                # 如果找不到闭引号，尝试在开引号之前找标点
                last_open = window.rfind("“", 0, cut)
                if last_open != -1:
                    search_end = last_open
                    adjusted_cut = cut
                    for j in range(search_end - 1, max(0, search_end - 30), -1):
                        if window[j] in PUNCT_FOR_SPLIT:
                            adjusted_cut = j + 1
                            break
                    if adjusted_cut == cut:
                        adjusted_cut = last_open
                    if 0 < adjusted_cut < cut:
                        cut = adjusted_cut

        # 最终切分
        chunk = s[:cut].strip()
        if not chunk:
            # 兜底：避免死循环，强制前进
            chunk = s[:max_len].strip()
            cut = max_len
        chunks.append(chunk)
        s = s[cut:].lstrip()

    if s:
        chunks.append(s)
    return chunks


def _append_bullet_with_limit(
    pages: list[dict[str, Any]],
    cur: dict[str, Any],
    bullet: str,
    rules: Rules,
    intent: str,
) -> dict[str, Any]:
    # 如果 bullet 本身超长，先切片
    for piece in _split_long_text(bullet, rules.max_chars_per_page):
        cur["bullets"].append(piece)
        cur.setdefault("items", []).append({"text": piece, "intent": intent})
        if _projected_len(cur) > rules.max_chars_per_page:
            # 回退这条，先落盘当前页，再开新页放进去
            cur["bullets"].pop()
            cur["items"].pop()
            pages.append(_finalize_page(cur))
            nxt_title = f"{cur.get('title', '知识点')}（续）"
            cur = _new_page(nxt_title, cur.get("page_type", "bullets"), topic=cur.get("topic", ""), first_signal="char_limit")
            cur["evidence"]["split_reason"].append("char_limit")
            cur["bullets"].append(piece)
            cur.setdefault("items", []).append({"text": piece, "intent": intent})
    return cur


def _append_quote_with_limit(
    pages: list[dict[str, Any]],
    cur: dict[str, Any],
    quote: str,
    rules: Rules,
    intent: str,
) -> dict[str, Any]:
    for piece in _split_long_text(quote, rules.max_chars_per_page):
        cur["quotes"].append(piece)
        cur.setdefault("items", []).append({"text": piece, "intent": intent})
        if _projected_len(cur) > rules.max_chars_per_page:
            cur["quotes"].pop()
            cur["items"].pop()
            pages.append(_finalize_page(cur))
            cur = _new_page("引用（续）", "quote", topic=cur.get("topic", ""), first_signal="char_limit")
            cur["evidence"]["split_reason"].append("char_limit")
            cur["quotes"].append(piece)
            cur.setdefault("items", []).append({"text": piece, "intent": intent})
    return cur


def _paginate(blocks: list[str], rules: Rules) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    cur = _new_page("开场", "teacher_only", topic="", first_signal="init")

    for line in blocks:
        text = line.strip()
        if not text:
            continue

        # 标签解析和 block 分类（优先级最高）
        tag, clean_text = parse_tag(text)
        score = score_line(clean_text, rules.keep_keywords, rules.drop_keywords) if tag is None else 0
        block_type, force_new_topic = classify_block(tag, clean_text, score, rules)

        # 可选的 AI 判定层（仅在无标签时生效）：
        # - intent: SHOW / SUPPORT / SAY
        # - is_anchor: 是否建议开启新知识点块
        ai_intent: Intent | None = None
        if tag is None:
            ai_intent, ai_anchor = safe_ai_classify(clean_text)
            if ai_intent is not None:
                # 用 intent + is_anchor 对规则分类做“软覆盖”（规则仍然兜底）
                if ai_intent == "SHOW":
                    block_type = "knowledge"
                elif ai_intent == "SUPPORT":
                    # 默认归为 example，后续会标记 intent=SUPPORT
                    block_type = "example"
                elif ai_intent == "SAY":
                    block_type = "teacher_only"
                if ai_anchor:
                    force_new_topic = True

        # 如果标签是标题页/章节页，直接处理
        if block_type == "title":
            if cur["bullets"] or cur["quotes"] or cur["page_type"] != "teacher_only":
                pages.append(_finalize_page(cur))
            title_page = _new_page(clean_text, "title_page", topic="", first_signal="title_tag")
            pages.append(_finalize_page(title_page))
            cur = _new_page("开场", "teacher_only", topic="", first_signal="after_title")
            continue

        if block_type == "section":
            if cur["bullets"] or cur["quotes"] or cur["page_type"] != "teacher_only":
                pages.append(_finalize_page(cur))
            sec = _new_page(clean_text, "section_page", topic=clean_text.split("：", 1)[0], first_signal="section_tag")
            pages.append(_finalize_page(sec))
            cur = _new_page("开场", "teacher_only", topic="", first_signal="after_section")
            continue

        # 如果 block_type 是 drop（可略），跳过
        if block_type == "drop":
            continue

        # 知识点锚点检测：强制新页
        if force_new_topic and (cur["bullets"] or cur["quotes"]):
            pages.append(_finalize_page(cur))
            # 尝试从文本中提取 topic（人物/概念名）
            topic = ""
            if _matches_anchor_pattern(clean_text, rules.anchor_patterns):
                # 提取可能的 topic（简单启发式：前 2-4 个中文字）
                match = re.match(r"^([\u4e00-\u9fff]{2,4})", clean_text)
                if match:
                    topic = match.group(1)
            cur = _new_page("知识点", "bullets", topic=topic, first_signal="anchor_trigger")
            cur["evidence"]["signals"].append("anchor_trigger")

        # 章节页：只展示标题，不排版；并切断上下文（保留原有逻辑作为兜底）
        if _is_section_title(clean_text):
            if cur["bullets"] or cur["quotes"] or cur["page_type"] != "teacher_only":
                pages.append(_finalize_page(cur))

            sec = _new_page(text, "section_page", topic=text.split("：", 1)[0], first_signal="section")
            pages.append(_finalize_page(sec))

            cur = _new_page("开场", "teacher_only", topic="", first_signal="after_section")
            continue

        # 引用页：尽量独立（标签优先，否则用原有检测）
        if block_type == "quote" or (block_type == "knowledge" and _is_quote_line(clean_text)):
            if cur["bullets"] and cur["page_type"] != "quote":
                pages.append(_finalize_page(cur))
                cur = _new_page("引用", "quote", topic=cur.get("topic", ""), first_signal="quote_block")

            cur["evidence"]["signals"].append("quote_block")
            # 引用类内容 → SUPPORT
            cur = _append_quote_with_limit(pages, cur, clean_text, rules, intent="SUPPORT")
            continue

        # 老师出镜/寒暄页（标签优先，否则用原有检测）
        if block_type == "teacher_only" or (block_type != "knowledge" and _looks_teacher_only(clean_text, rules)):
            if cur["page_type"] != "teacher_only" and (cur["bullets"] or cur["quotes"]):
                pages.append(_finalize_page(cur))
                cur = _new_page("老师出镜", "teacher_only", topic="", first_signal="teacher_only")

            for b in _split_to_bullets(clean_text):
                # 老师出镜 / 寒暄页 → SAY
                cur = _append_bullet_with_limit(pages, cur, b, rules, intent="SAY")
            continue

        # 知识点页（knowledge / example）
        if block_type in ("knowledge", "example"):
            for b in _split_to_bullets(clean_text):
                # 产品级硬规则：同一页最多 1 个“主知识点锚点”
                # 若当前页已经出现过主锚点，再遇到新的主锚点 -> 立即落盘开新页（不管字数）
                if is_main_knowledge_anchor(b):
                    if cur.get("_has_main_anchor") and (cur["bullets"] or cur["quotes"]):
                        pages.append(_finalize_page(cur))
                        cur = _new_page("知识点", "bullets", topic=cur.get("topic", ""), first_signal="main_anchor_conflict")
                        cur["evidence"]["split_reason"].append("main_anchor_conflict")
                    cur["_has_main_anchor"] = True

                # 不相关尽量拆页（轻量相似度）
                if rules.topic_split_enabled and cur["bullets"]:
                    sim = _avg_similarity_to_page(b, cur["bullets"])
                    if sim < rules.similarity_threshold:
                        pages.append(_finalize_page(cur))
                        cur = _new_page("知识点", "bullets", topic=cur.get("topic", ""), first_signal="topic_diverge")
                        cur["evidence"]["split_reason"].append("topic_diverge")

                if cur["page_type"] == "teacher_only":
                    # 从老师出镜进入知识点
                    if cur["bullets"] or cur["quotes"]:
                        pages.append(_finalize_page(cur))
                    cur = _new_page("知识点", "bullets", topic=cur.get("topic", ""), first_signal="enter_knowledge")
                    cur["evidence"]["split_reason"].append("enter_knowledge")

                # 知识点主体 → SHOW；例子说明 → SUPPORT
                intent = "SHOW" if block_type == "knowledge" else "SUPPORT"
                cur = _append_bullet_with_limit(pages, cur, b, rules, intent=intent)

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
    if pt == "quote":
        # 引用页默认使用半屏版式（可在后续映射到专门模板）
        return "半屏"

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


# -----------------------------
# 跨页凝聚力后处理器（引子句搬运）
# -----------------------------
LEADIN_PATTERNS = ("尤其是", "比如", "例如", "包括", "代表作", "分为", "主要是", "其一", "其二")


def _looks_leadin_bullet(s: str) -> bool:
    """判断一条 bullet 是否是“引子句”（未完结，需要接下文）。"""
    s = (s or "").strip()
    if not s:
        return False
    if len(s) <= 40 and any(k in s for k in LEADIN_PATTERNS):
        return True
    if s.endswith(("，", "：", "（", "(", "——")):
        return True
    return False


def _recalc_page(p: dict[str, Any]) -> None:
    """重新计算页面的 content 和 char_count（搬运后需要更新）。"""
    lines: list[str] = []
    lines.extend(p.get("bullets", []))
    lines.extend(p.get("quotes", []))
    p["content"] = "\n".join(lines).strip()
    p["char_count"] = len(p["content"])


def enforce_topic_cohesion(pages: list[dict[str, Any]], max_chars: int) -> None:
    """
    产品级“凝聚力”修正：
    - 如果上一页末尾是“引子句”，且下一页是同类知识点页，则把引子句搬到下一页开头
    - 前提：搬运后两页都不超过 max_chars
    """
    for i in range(len(pages) - 1):
        a = pages[i]
        b = pages[i + 1]

        # 只对知识点页做（避免动章节页/标题页/老师出镜）
        if a.get("layout") in ("章节页", "标题页", "老师出镜"):
            continue
        if b.get("layout") in ("章节页", "标题页", "老师出镜"):
            continue

        a_bullets = a.get("bullets", [])
        b_bullets = b.get("bullets", [])

        if not a_bullets:
            continue

        tail = a_bullets[-1]
        if not _looks_leadin_bullet(tail):
            continue

        # 尝试搬运 tail 到下一页开头
        a_bullets.pop()
        b_bullets.insert(0, tail)

        _recalc_page(a)
        _recalc_page(b)

        if a["char_count"] <= max_chars and b["char_count"] <= max_chars:
            a.setdefault("evidence", {}).setdefault("signals", []).append("cohesion_move_to_next")
            b.setdefault("evidence", {}).setdefault("signals", []).append("cohesion_receive_from_prev")
            continue

        # 不满足就回滚
        b_bullets.pop(0)
        a_bullets.append(tail)
        _recalc_page(a)
        _recalc_page(b)


def enforce_layout(layout: str) -> str:
    """
    产品级约束：layout 必须是严格枚举值。
    - 非法 layout → 兜底为"半屏"（Demo 阶段）
    - 严格模式可改为 raise ValueError
    """
    if layout not in ALLOWED_LAYOUTS:
        # Demo 阶段：兜底为半屏
        return "半屏"
        # 严格模式（产品稳定期）：
        # raise ValueError(f"Invalid layout: {layout}, must be one of {ALLOWED_LAYOUTS}")
    return layout


def enforce_page_structure(p: dict[str, Any]) -> None:
    """
    产品级约束：标题页/章节页必须干净（无正文内容）。
    - 标题页/章节页：bullets/quotes/content/char_count 必须为空
    """
    if p.get("layout") in ("标题页", "章节页"):
        p["bullets"] = []
        p["quotes"] = []
        p["content"] = ""
        p["char_count"] = 0


