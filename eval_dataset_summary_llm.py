#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Optional LLM-as-judge evaluation for `dataset_summary`.

This module is imported (optionally) by `evaluation.py`.

Interface requirement (must match the call site in `evaluation.py`):

    llm_metrics = eval_dataset_summary_llm(gt_obj, pred_obj)

So this module MUST export `eval_dataset_summary_llm(gt_obj, pred_obj, ...)`.

No hard-coded keys: by default, the OpenAI Python SDK reads `OPENAI_API_KEY`
from the environment. (You can still pass a client explicitly if you want.)

Returned metrics (flat dict):
  - summary_llm_coverage:      float in [0,1]
  - summary_llm_faithfulness:  float in [0,1]
  - summary_llm_score:         float in [0,1]  (= mean of the above)
  - summary_llm_justification: short English string
"""

from __future__ import annotations

import json
import os
import re
from typing import Any, Dict, Optional


def safe_div(num: float, den: float) -> float:
    if den == 0:
        return 0.0
    return num / den


def _clamp01(x: float) -> float:
    if x < 0.0:
        return 0.0
    if x > 1.0:
        return 1.0
    return x


def _extract_summary_text(obj: Dict[str, Any]) -> str:
    """Concatenate dataset_summary.plain_description and dataset_summary.notes."""
    summ = obj.get("dataset_summary", {}) or {}
    parts = [
        str(summ.get("plain_description", "") or "").strip(),
        str(summ.get("notes", "") or "").strip(),
    ]
    parts = [p for p in parts if p]
    return "\n".join(parts).strip()


def build_summary_llm_prompt(gt_text: str, pred_text: str) -> str:
    return (
        "You are evaluating a model-generated dataset summary.\n\n"
        "[Gold reference summary]\n"
        f"{gt_text}\n\n"
        "[Predicted summary]\n"
        f"{pred_text}\n\n"
        "Judge similarity in TERMS OF INFORMATION CONTENT (not wording).\n\n"
        "Return ONLY a JSON object with:\n"
        "- coverage: number in [0,1], how much important info from GOLD is present in PRED.\n"
        "- faithfulness: number in [0,1], factual faithfulness of PRED relative to GOLD.\n"
        "  Deduct for hallucinations, contradictions, or unsupported claims.\n"
        "- justification: 1-3 short English sentences explaining the scores.\n\n"
        "Output JSON only, no markdown, no extra text."
    ).strip()


_JSON_OBJ_RE = re.compile(r"\{.*\}", re.DOTALL)


def _parse_json_object(text: str) -> Dict[str, Any]:
    """Best-effort parse a JSON object from model output."""
    if not isinstance(text, str):
        text = str(text)

    cleaned = text.strip()
    # strip common code fences
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    cleaned = cleaned.strip()

    try:
        parsed = json.loads(cleaned)
        if isinstance(parsed, dict):
            return parsed
    except Exception:
        pass

    m = _JSON_OBJ_RE.search(cleaned)
    if m:
        snippet = m.group(0)
        try:
            parsed2 = json.loads(snippet)
            if isinstance(parsed2, dict):
                return parsed2
        except Exception:
            pass

    return {}


def eval_dataset_summary_llm(
    gt_obj: Dict[str, Any],
    pred_obj: Dict[str, Any],
    model: str = "gpt-4.1-mini",
    client: Optional[Any] = None,
    temperature: float = 0.0,
    max_tokens: int = 256,
) -> Dict[str, Any]:
    """Use an LLM judge to compare dataset summaries.

    NOTE: The first two args MUST be (gt_obj, pred_obj) to match `evaluation.py`.
    """
    gt_text = _extract_summary_text(gt_obj)
    pred_text = _extract_summary_text(pred_obj)

    # Fast paths that do not require API calls
    if not gt_text and not pred_text:
        return {
            "summary_llm_coverage": 0.0,
            "summary_llm_faithfulness": 0.0,
            "summary_llm_score": 0.0,
            "summary_llm_justification": "Both gold and predicted summaries are empty.",
        }

    if gt_text and not pred_text:
        cov = 0.0
        faith = 1.0
        return {
            "summary_llm_coverage": cov,
            "summary_llm_faithfulness": faith,
            "summary_llm_score": safe_div(cov + faith, 2.0),
            "summary_llm_justification": "Predicted summary is empty while gold summary is non-empty.",
        }

    if not gt_text and pred_text:
        return {
            "summary_llm_coverage": 0.0,
            "summary_llm_faithfulness": 0.0,
            "summary_llm_score": 0.0,
            "summary_llm_justification": "Gold summary is empty but predicted summary is non-empty.",
        }

    prompt = build_summary_llm_prompt(gt_text, pred_text)
    model = os.getenv("SUMMARY_LLM_MODEL", model)

    if client is None:
        from openai import OpenAI  # type: ignore
        client = OpenAI()  # reads OPENAI_API_KEY from env

    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=temperature,
            max_tokens=max_tokens,
            response_format={"type": "json_object"},
        )
    except TypeError:
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=temperature,
            max_tokens=max_tokens,
        )

    content = getattr(resp.choices[0].message, "content", None) or ""
    parsed = _parse_json_object(content)

    try:
        coverage = float(parsed.get("coverage", 0.0))
    except Exception:
        coverage = 0.0

    try:
        faithfulness = float(parsed.get("faithfulness", 0.0))
    except Exception:
        faithfulness = 0.0

    justification = str(parsed.get("justification", "") or "").strip()
    if not justification:
        justification = "Failed to parse justification from judge output."

    coverage = _clamp01(coverage)
    faithfulness = _clamp01(faithfulness)
    score = safe_div(coverage + faithfulness, 2.0)

    return {
        "summary_llm_coverage": coverage,
        "summary_llm_faithfulness": faithfulness,
        "summary_llm_score": score,
        "summary_llm_justification": justification,
    }

