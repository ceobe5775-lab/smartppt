from __future__ import annotations

"""
AI 判定层 Hook（最小骨架）

职责边界：
- 仅负责对单段文本做两类判定：
  1) intent: SHOW / SUPPORT / SAY
  2) is_anchor: 是否建议作为“新知识点块”的起点

重要约束：
- 只返回 JSON 友好的字典；engine 侧通过 safe_ai_classify 做严格兜底。
- 置信度不足 / 调用异常 / 返回非法值时，engine 必须回退到规则层逻辑。
"""

import json
import os
import urllib.request
from typing import Any, Literal, TypedDict


Intent = Literal["SHOW", "SUPPORT", "SAY"]


class AIClassifyResult(TypedDict, total=False):
    intent: Intent          # SHOW / SUPPORT / SAY
    is_anchor: bool         # 是否建议开启新知识点
    confidence: float       # 0.0 ~ 1.0


def ai_classify(text: str) -> AIClassifyResult:
    """
    HTTP 调用实现（仅用标准库，无需安装第三方依赖）。

    你只需要把默认地址里的 "https://your-ai-host/intent-anchor"
    替换成你们真实的服务地址，或设置环境变量 AI_CLASSIFY_ENDPOINT。

    约定返回字段不变：
    {
      "intent": "SHOW" | "SUPPORT" | "SAY",
      "is_anchor": true | false,
      "confidence": 0.0 ~ 1.0
    }
    """
    endpoint = os.getenv("AI_CLASSIFY_ENDPOINT", "https://your-ai-host/intent-anchor").strip()
    if not endpoint:
        return AIClassifyResult()

    payload = json.dumps({"text": text}).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=payload,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    # 超时、解析失败等异常由 safe_ai_classify 捕获并回退规则
    with urllib.request.urlopen(req, timeout=2.0) as resp:  # noqa: S310
        body = resp.read().decode("utf-8", errors="replace")
    data = json.loads(body) or {}

    return AIClassifyResult(
        intent=data.get("intent"),
        is_anchor=bool(data.get("is_anchor", False)),
        confidence=float(data.get("confidence", 0.0) or 0.0),
    )


def safe_ai_classify(text: str, *, min_confidence: float = 0.6) -> tuple[Intent | None, bool]:
    """
    安全封装：
    - 任何异常 / 低置信度 / 非法值 → (None, False)，engine 回退规则。
    """
    try:
        result = ai_classify(text) or {}
    except Exception:  # noqa: BLE001
        return None, False

    conf = float(result.get("confidence", 0.0) or 0.0)
    if conf < min_confidence:
        return None, False

    intent = result.get("intent")
    is_anchor = bool(result.get("is_anchor", False))

    if intent not in ("SHOW", "SUPPORT", "SAY"):
        return None, False

    return intent, is_anchor


