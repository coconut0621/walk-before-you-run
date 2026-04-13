#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Stage-0 evaluator for messy workbook/XLSX datasets with explicit source_ref.

Design goals:
- Keep the original metric style from evaluation.py.
- Replace physical_name-only table matching with source_ref-aware fuzzy matching.
- Avoid CSV-specific oracle logic from sanity_check.py.
- Support bulk evaluation via a companion script.

Main ideas:
1) Match GT tables to predicted tables using a weighted similarity over:
   - source_ref
   - table names / logical names
   - matched column coverage
   - row count closeness
   - sample-row token overlap
   - role histogram similarity
2) Evaluate columns within each matched table using fuzzy name alignment.
3) Evaluate relations after projecting predicted endpoints/keys onto GT tables/columns.
4) Compare profiling/outlier fields directly against GT on matched columns/tables.

Usage:
    python evaluation_xlsx_source_ref.py --gt stage0_gt.json --pred pred.json --pretty
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

try:
    from eval_dataset_summary_llm import eval_dataset_summary_llm  # type: ignore
    HAS_LLM_EVAL = True
except Exception:
    eval_dataset_summary_llm = None  # type: ignore
    HAS_LLM_EVAL = False


# ------------------------- basic helpers -------------------------


def load_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)



def safe_div(num: float, den: float) -> float:
    return 0.0 if den == 0 else num / den



def norm_name(s: Any) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = s.replace("_", " ")
    s = re.sub(r"\s+", " ", s)
    return s



def norm_compact(s: Any) -> str:
    s = norm_name(s)
    return re.sub(r"[^a-z0-9]+", "", s)



def tokenize(s: Any) -> List[str]:
    s = norm_name(s)
    return re.findall(r"[a-z0-9]+", s)



def jaccard(a: Set[str], b: Set[str]) -> float:
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)



def seq_ratio(a: str, b: str) -> float:
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()



def exactish_equal(a: Any, b: Any) -> bool:
    aa = norm_compact(a)
    bb = norm_compact(b)
    return aa != "" and aa == bb



def maybe_float(x: Any) -> Optional[float]:
    try:
        if x is None or x == "":
            return None
        return float(x)
    except Exception:
        return None



def _avg(xs: List[float]) -> float:
    return float(sum(xs) / len(xs)) if xs else 0.0



def _score_abs_diff(pred: Optional[float], gt: Optional[float], abs_tol: float) -> Optional[float]:
    if gt is None:
        return None
    if pred is None:
        return 0.0
    err = abs(pred - gt)
    if err <= abs_tol:
        return 1.0
    return max(0.0, 1.0 - (err - abs_tol) / max(abs_tol, 1e-12))



def _score_rel_diff(pred: Optional[float], gt: Optional[float], rel_tol: float) -> Optional[float]:
    if gt is None:
        return None
    if pred is None:
        return 0.0
    scale = max(abs(gt), 1.0)
    rel_err = abs(pred - gt) / scale
    if rel_err <= rel_tol:
        return 1.0
    return max(0.0, 1.0 - (rel_err - rel_tol) / max(rel_tol, 1e-12))



def _concat_summary_text(obj: Dict[str, Any]) -> str:
    s = obj.get("dataset_summary", {}) or {}
    return f"{s.get('plain_description', '')} {s.get('notes', '')}".strip()


# ------------------------- lexical summary metrics -------------------------


def _basic_tokenize(text: str) -> Set[str]:
    text = norm_name(text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    toks = text.split()
    stopwords = {
        "the", "a", "an", "and", "or", "of", "for", "to", "in", "on", "with",
        "this", "that", "these", "those", "is", "are", "was", "were", "be",
        "as", "by", "at", "from", "it", "its", "into", "such", "also",
    }
    return {t for t in toks if t and t not in stopwords}



def compute_dataset_summary_lexical(gt_obj: Dict[str, Any], pred_obj: Dict[str, Any]) -> Dict[str, float]:
    gt_text = _concat_summary_text(gt_obj)
    pred_text = _concat_summary_text(pred_obj)

    gt_tokens = _basic_tokenize(gt_text)
    pred_tokens = _basic_tokenize(pred_text)

    inter = gt_tokens & pred_tokens
    union = gt_tokens | pred_tokens
    p = safe_div(len(inter), len(pred_tokens))
    r = safe_div(len(inter), len(gt_tokens))
    f1 = safe_div(2 * p * r, p + r) if (p + r) > 0 else 0.0
    j = safe_div(len(inter), len(union))
    return {
        "summary_precision": p,
        "summary_recall": r,
        "summary_f1": f1,
        "summary_jaccard": j,
    }



def compute_summary_length_metrics(gt_obj: Dict[str, Any], pred_obj: Dict[str, Any]) -> Dict[str, float]:
    gt_text = _concat_summary_text(gt_obj)
    pred_text = _concat_summary_text(pred_obj)

    lg = len(gt_text.strip())
    lp = len(pred_text.strip())

    if lg == 0 and lp == 0:
        length_score = 1.0
    elif lg == 0 and lp > 0:
        length_score = 0.0
    else:
        ratio = lp / max(1.0, float(lg))
        if 0.5 <= ratio <= 2.0:
            length_score = 1.0
        elif ratio < 0.5:
            length_score = max(0.0, ratio / 0.5)
        else:
            length_score = max(0.0, 2.0 / ratio)

    return {
        "summary_length_chars_gt": float(lg),
        "summary_length_chars_pred": float(lp),
        "summary_length_ratio_pred_to_gt": float(lp / max(1.0, float(lg))) if lg > 0 else 0.0,
        "summary_length_score": float(length_score),
    }


# ------------------------- matching data structures -------------------------


@dataclass
class ColumnMatch:
    gt_idx: int
    pred_idx: int
    sim: float


@dataclass
class TableMatch:
    gt_idx: int
    pred_idx: int
    sim: float


# ------------------------- fingerprint helpers -------------------------


def get_table_aliases(t: Dict[str, Any]) -> Set[str]:
    vals = {
        norm_compact(t.get("physical_name", "")),
        norm_compact(t.get("logical_name", "")),
        norm_compact(t.get("table_id", "")),
    }
    src = t.get("source_ref", {}) or {}
    if isinstance(src, dict):
        vals.update({
            norm_compact(src.get("sheet_name", "")),
            norm_compact(src.get("anchor_text", "")),
            norm_compact(src.get("block_id", "")),
            norm_compact(src.get("extraction_kind", "")),
            norm_compact(src.get("cell_region", "")),
        })
    return {v for v in vals if v}



def get_col_aliases(c: Dict[str, Any]) -> Set[str]:
    vals = {
        norm_compact(c.get("physical_name", "")),
        norm_compact(c.get("logical_name", "")),
    }
    return {v for v in vals if v}



def get_sample_value_tokens(table_obj: Dict[str, Any]) -> Set[str]:
    out: Set[str] = set()
    for row in table_obj.get("sample_rows", []) or []:
        if not isinstance(row, dict):
            continue
        for _, v in row.items():
            if v is None:
                continue
            s = str(v)
            if len(s) > 64:
                continue
            for tok in tokenize(s):
                out.add(tok)
    return out



def rowcount_similarity(gt_t: Dict[str, Any], pred_t: Dict[str, Any]) -> float:
    a = maybe_float(gt_t.get("n_rows"))
    b = maybe_float(pred_t.get("n_rows"))
    if a is None or b is None:
        return 0.0
    if a == b:
        return 1.0
    scale = max(abs(a), abs(b), 1.0)
    return max(0.0, 1.0 - abs(a - b) / scale)



def source_ref_field_similarity(gt_src: Dict[str, Any], pred_src: Dict[str, Any], field: str) -> Optional[float]:
    ga = gt_src.get(field)
    pa = pred_src.get(field)
    if ga is None and pa is None:
        return None
    if field in {"sheet_name", "block_id", "extraction_kind", "cell_region"}:
        if not ga and not pa:
            return None
        return 1.0 if exactish_equal(ga, pa) else seq_ratio(norm_compact(ga), norm_compact(pa))
    if not ga and not pa:
        return None
    return seq_ratio(norm_compact(ga), norm_compact(pa))



def source_ref_similarity(gt_t: Dict[str, Any], pred_t: Dict[str, Any]) -> Optional[float]:
    gt_src = gt_t.get("source_ref")
    pred_src = pred_t.get("source_ref")
    if not isinstance(gt_src, dict) or not isinstance(pred_src, dict):
        return None

    parts = []
    weights = []
    for field, w in [
        ("sheet_name", 0.30),
        ("anchor_text", 0.20),
        ("block_id", 0.20),
        ("extraction_kind", 0.20),
        ("cell_region", 0.10),
    ]:
        sim = source_ref_field_similarity(gt_src, pred_src, field)
        if sim is not None:
            parts.append(sim)
            weights.append(w)
    if not weights:
        return None
    return sum(p * w for p, w in zip(parts, weights)) / sum(weights)


# ------------------------- assignment DP -------------------------


def best_assignment(sim_matrix: List[List[float]]) -> List[Tuple[int, int, float]]:
    n = len(sim_matrix)
    m = len(sim_matrix[0]) if n else 0
    if n == 0 or m == 0:
        return []

    transposed = False
    if n > m:
        sim_matrix = [list(x) for x in zip(*sim_matrix)]
        n, m = m, n
        transposed = True

    from functools import lru_cache

    @lru_cache(None)
    def dp(i: int, mask: int) -> Tuple[float, Tuple[Tuple[int, int], ...]]:
        if i == n:
            return 0.0, ()
        best_score, best_pairs = dp(i + 1, mask)
        for j in range(m):
            if (mask >> j) & 1:
                continue
            s = sim_matrix[i][j]
            score2, pairs2 = dp(i + 1, mask | (1 << j))
            score2 += s
            if score2 > best_score:
                best_score = score2
                best_pairs = pairs2 + ((i, j),)
        return best_score, best_pairs

    _, pairs = dp(0, 0)
    out: List[Tuple[int, int, float]] = []
    for i, j in pairs:
        if transposed:
            out.append((j, i, sim_matrix[i][j]))
        else:
            out.append((i, j, sim_matrix[i][j]))
    return sorted(out)


# ------------------------- column matching -------------------------


def column_name_similarity(gt_c: Dict[str, Any], pred_c: Dict[str, Any]) -> float:
    gt_aliases = get_col_aliases(gt_c)
    pred_aliases = get_col_aliases(pred_c)
    if not gt_aliases and not pred_aliases:
        return 0.0
    if gt_aliases & pred_aliases:
        return 1.0
    best = 0.0
    for a in gt_aliases:
        for b in pred_aliases:
            best = max(best, seq_ratio(a, b))
    return best



def column_similarity(gt_c: Dict[str, Any], pred_c: Dict[str, Any]) -> float:
    name_sim = column_name_similarity(gt_c, pred_c)
    dtype_eq = 1.0 if exactish_equal(gt_c.get("dtype"), pred_c.get("dtype")) and gt_c.get("dtype") else 0.0
    role_eq = 1.0 if exactish_equal(gt_c.get("role"), pred_c.get("role")) and gt_c.get("role") else 0.0
    nullable_eq = 1.0 if gt_c.get("is_nullable") == pred_c.get("is_nullable") else 0.0
    unique_eq = 1.0 if gt_c.get("is_unique") == pred_c.get("is_unique") else 0.0
    return min(1.0, 0.70 * name_sim + 0.10 * dtype_eq + 0.10 * role_eq + 0.05 * nullable_eq + 0.05 * unique_eq)



def match_columns(
    gt_table: Dict[str, Any],
    pred_table: Dict[str, Any],
    threshold: float = 0.55,
) -> Tuple[List[ColumnMatch], Dict[int, int], Dict[int, int]]:
    gt_cols = gt_table.get("columns", []) or []
    pred_cols = pred_table.get("columns", []) or []
    sim_matrix = [[column_similarity(g, p) for p in pred_cols] for g in gt_cols]
    pairs = best_assignment(sim_matrix)

    matches: List[ColumnMatch] = []
    gt2pred: Dict[int, int] = {}
    pred2gt: Dict[int, int] = {}
    for gi, pj, sim in pairs:
        if sim >= threshold:
            matches.append(ColumnMatch(gi, pj, sim))
            gt2pred[gi] = pj
            pred2gt[pj] = gi
    return matches, gt2pred, pred2gt


# ------------------------- table matching -------------------------


def table_name_similarity(gt_t: Dict[str, Any], pred_t: Dict[str, Any]) -> float:
    gt_aliases = get_table_aliases(gt_t)
    pred_aliases = get_table_aliases(pred_t)
    if gt_aliases & pred_aliases:
        return 1.0
    best = 0.0
    for a in gt_aliases:
        for b in pred_aliases:
            best = max(best, seq_ratio(a, b))
    return best



def role_hist_similarity(gt_t: Dict[str, Any], pred_t: Dict[str, Any]) -> float:
    def hist(t: Dict[str, Any]) -> Dict[str, int]:
        d: Dict[str, int] = {}
        for c in t.get("columns", []) or []:
            r = norm_compact(c.get("role", ""))
            if r:
                d[r] = d.get(r, 0) + 1
        return d

    a = hist(gt_t)
    b = hist(pred_t)
    keys = set(a) | set(b)
    if not keys:
        return 0.0
    num = sum(min(a.get(k, 0), b.get(k, 0)) for k in keys)
    den = sum(max(a.get(k, 0), b.get(k, 0)) for k in keys)
    return safe_div(num, den)



def table_similarity(gt_t: Dict[str, Any], pred_t: Dict[str, Any]) -> float:
    src_sim = source_ref_similarity(gt_t, pred_t)
    name_sim = table_name_similarity(gt_t, pred_t)
    col_matches, _, _ = match_columns(gt_t, pred_t, threshold=0.0)
    gt_cols = gt_t.get("columns", []) or []
    pred_cols = pred_t.get("columns", []) or []
    col_cov = safe_div(sum(m.sim for m in col_matches), max(len(gt_cols), len(pred_cols), 1))
    row_sim = rowcount_similarity(gt_t, pred_t)
    samp_sim = jaccard(get_sample_value_tokens(gt_t), get_sample_value_tokens(pred_t))
    role_sim = role_hist_similarity(gt_t, pred_t)

    parts: List[Tuple[float, float]] = []
    parts.append((col_cov, 0.42))
    parts.append((row_sim, 0.10))
    parts.append((samp_sim, 0.08))
    parts.append((role_sim, 0.08))
    if src_sim is None:
        parts.append((name_sim, 0.32))
    else:
        parts.append((src_sim, 0.24))
        parts.append((name_sim, 0.08))

    return sum(v * w for v, w in parts) / sum(w for _, w in parts)



def match_tables(
    gt_tables: List[Dict[str, Any]],
    pred_tables: List[Dict[str, Any]],
    threshold: float = 0.58,
) -> Tuple[List[TableMatch], Dict[int, int], Dict[int, int], List[List[float]]]:
    sim_matrix = [[table_similarity(g, p) for p in pred_tables] for g in gt_tables]
    pairs = best_assignment(sim_matrix)

    matches: List[TableMatch] = []
    gt2pred: Dict[int, int] = {}
    pred2gt: Dict[int, int] = {}
    for gi, pj, sim in pairs:
        if sim >= threshold:
            matches.append(TableMatch(gi, pj, sim))
            gt2pred[gi] = pj
            pred2gt[pj] = gi
    return matches, gt2pred, pred2gt, sim_matrix


# ------------------------- relation evaluation -------------------------


def resolve_relation_endpoint_table(rel: Dict[str, Any], tables: List[Dict[str, Any]], side: str) -> Optional[int]:
    key_id = f"{side}_table_id"
    key_name1 = f"{side}_table"
    key_name2 = f"{side}_table_name"

    if rel.get(key_id) is not None:
        want = str(rel.get(key_id))
        for i, t in enumerate(tables):
            if str(t.get("table_id")) == want:
                return i

    want_name = rel.get(key_name1) or rel.get(key_name2)
    if want_name:
        want_norm = norm_compact(want_name)
        for i, t in enumerate(tables):
            aliases = get_table_aliases(t)
            if want_norm in aliases:
                return i
    return None



def resolve_relation_keys(rel: Dict[str, Any], side: str) -> List[str]:
    vals = rel.get(f"{side}_key", []) or []
    return [norm_compact(v) for v in vals if norm_compact(v)]



def canonical_relation(table_a: str, table_b: str, key_a: Sequence[str], key_b: Sequence[str]) -> Tuple[Tuple[str, str], Tuple[Tuple[str, ...], Tuple[str, ...]]]:
    ka = tuple(sorted(norm_compact(k) for k in key_a if norm_compact(k)))
    kb = tuple(sorted(norm_compact(k) for k in key_b if norm_compact(k)))
    ta = norm_compact(table_a)
    tb = norm_compact(table_b)
    if (ta, ka) <= (tb, kb):
        return (ta, tb), (ka, kb)
    return (tb, ta), (kb, ka)



def build_gt_relation_set(gt: Dict[str, Any]) -> Set[Tuple[Tuple[str, str], Tuple[Tuple[str, ...], Tuple[str, ...]]]]:
    tables = gt.get("tables", []) or []
    id2tid = {str(t.get("table_id")): str(t.get("table_id")) for t in tables if t.get("table_id") is not None}
    out = set()
    for rel in gt.get("relations", []) or []:
        pa = str(rel.get("parent_table_id", ""))
        ch = str(rel.get("child_table_id", ""))
        if pa not in id2tid or ch not in id2tid:
            continue
        out.add(canonical_relation(
            id2tid[pa],
            id2tid[ch],
            rel.get("parent_key", []) or [],
            rel.get("child_key", []) or [],
        ))
    return out



def build_pred_relation_set_projected_to_gt(
    gt: Dict[str, Any],
    pred: Dict[str, Any],
    pred2gt_table: Dict[int, int],
    col_maps_by_table: Dict[Tuple[int, int], Dict[int, int]],
) -> Set[Tuple[Tuple[str, str], Tuple[Tuple[str, ...], Tuple[str, ...]]]]:
    gt_tables = gt.get("tables", []) or []
    pred_tables = pred.get("tables", []) or []
    gt_table_id_by_idx = {i: str(t.get("table_id")) for i, t in enumerate(gt_tables)}
    out = set()

    def project_keys(pred_table_idx: int, pred_key_names: List[str]) -> List[str]:
        if pred_table_idx not in pred2gt_table:
            return []
        gt_table_idx = pred2gt_table[pred_table_idx]
        pred_cols = pred_tables[pred_table_idx].get("columns", []) or []
        gt_cols = gt_tables[gt_table_idx].get("columns", []) or []
        predcol_to_gtcol = col_maps_by_table.get((gt_table_idx, pred_table_idx), {})

        out_keys: List[str] = []
        for want in pred_key_names:
            want_norm = norm_compact(want)
            mapped = None
            for pj, pcol in enumerate(pred_cols):
                aliases = get_col_aliases(pcol)
                if want_norm in aliases and pj in predcol_to_gtcol:
                    gi = predcol_to_gtcol[pj]
                    mapped = gt_cols[gi].get("physical_name", "")
                    break
            if mapped is None:
                mapped = want_norm
            if norm_compact(mapped):
                out_keys.append(mapped)
        return out_keys

    for rel in pred.get("relations", []) or []:
        pidx = resolve_relation_endpoint_table(rel, pred_tables, "parent")
        cidx = resolve_relation_endpoint_table(rel, pred_tables, "child")
        if pidx is None or cidx is None:
            continue
        if pidx not in pred2gt_table or cidx not in pred2gt_table:
            continue

        gt_pidx = pred2gt_table[pidx]
        gt_cidx = pred2gt_table[cidx]
        pa_tid = gt_table_id_by_idx[gt_pidx]
        ch_tid = gt_table_id_by_idx[gt_cidx]

        pkeys = project_keys(pidx, resolve_relation_keys(rel, "parent"))
        ckeys = project_keys(cidx, resolve_relation_keys(rel, "child"))
        out.add(canonical_relation(pa_tid, ch_tid, pkeys, ckeys))
    return out


# ------------------------- direct profiling/outlier comparison -------------------------


def compute_profile_metrics_direct(
    gt_tables: List[Dict[str, Any]],
    pred_tables: List[Dict[str, Any]],
    table_matches: List[TableMatch],
    col_matches_by_table: Dict[Tuple[int, int], List[ColumnMatch]],
) -> Dict[str, float]:
    density_scores: List[float] = []
    parse_scores: List[float] = []
    consistency_scores: List[float] = []
    numeric_scores: List[float] = []
    date_scores: List[float] = []

    n_numeric_stats_scored = 0
    n_date_cols_scored = 0
    n_profile_columns_scored = 0

    for tm in table_matches:
        gt_t = gt_tables[tm.gt_idx]
        pred_t = pred_tables[tm.pred_idx]

        gt_density = maybe_float((gt_t.get("profile", {}) or {}).get("density"))
        pred_density = maybe_float((pred_t.get("profile", {}) or {}).get("density"))
        score = _score_abs_diff(pred_density, gt_density, 0.05)
        if score is not None:
            density_scores.append(score)

        for cm in col_matches_by_table.get((tm.gt_idx, tm.pred_idx), []):
            gt_c = (gt_t.get("columns", []) or [])[cm.gt_idx]
            pred_c = (pred_t.get("columns", []) or [])[cm.pred_idx]
            gt_prof = gt_c.get("profile", {}) or {}
            pred_prof = pred_c.get("profile", {}) or {}
            if not isinstance(gt_prof, dict):
                gt_prof = {}
            if not isinstance(pred_prof, dict):
                pred_prof = {}

            n_profile_columns_scored += 1

            s = _score_abs_diff(maybe_float(pred_prof.get("parse_success_rate")), maybe_float(gt_prof.get("parse_success_rate")), 0.05)
            if s is not None:
                parse_scores.append(s)

            s = _score_abs_diff(maybe_float(pred_prof.get("type_consistency")), maybe_float(gt_prof.get("type_consistency")), 0.05)
            if s is not None:
                consistency_scores.append(s)

            local_numeric_scores: List[float] = []
            for k, tol in [("min", 0.10), ("max", 0.10), ("mean", 0.10), ("median", 0.10)]:
                s = _score_rel_diff(maybe_float(pred_prof.get(k)), maybe_float(gt_prof.get(k)), tol)
                if s is not None:
                    local_numeric_scores.append(s)
            if local_numeric_scores:
                numeric_scores.append(_avg(local_numeric_scores))
                n_numeric_stats_scored += 1

            local_date_scores: List[float] = []
            s = _score_abs_diff(maybe_float(pred_prof.get("format_consistency")), maybe_float(gt_prof.get("format_consistency")), 0.10)
            if s is not None:
                local_date_scores.append(s)
            gt_hint = gt_prof.get("format_hint")
            pred_hint = pred_prof.get("format_hint")
            if gt_hint is not None:
                local_date_scores.append(1.0 if exactish_equal(gt_hint, pred_hint) else 0.0)
            if local_date_scores:
                date_scores.append(_avg(local_date_scores))
                n_date_cols_scored += 1

    profile_components: List[float] = []
    if density_scores:
        profile_components.append(_avg(density_scores))
    if parse_scores:
        profile_components.append(_avg(parse_scores))
    if consistency_scores:
        profile_components.append(_avg(consistency_scores))
    if numeric_scores:
        profile_components.append(_avg(numeric_scores))
    if date_scores:
        profile_components.append(_avg(date_scores))
    profiling_overall = _avg(profile_components)

    return {
        "profile_table_density_score": _avg(density_scores),
        "profile_parse_success_score": _avg(parse_scores),
        "profile_type_consistency_score": _avg(consistency_scores),
        "profile_numeric_stats_score": _avg(numeric_scores),
        "profile_date_format_score": _avg(date_scores),
        "profiling_overall_score": profiling_overall,
        "profile_n_columns_scored": float(n_profile_columns_scored),
        "profile_n_numeric_cols_scored": float(n_numeric_stats_scored),
        "profile_n_date_cols_scored": float(n_date_cols_scored),
    }



def _extract_top1_outlier_value(prof: Dict[str, Any]) -> Optional[float]:
    keys = [
        "top_1_outlier", "top1_outlier", "top_outlier",
        "top_1_value", "top1_value", "top_outlier_value",
    ]
    out = prof.get("outlier")
    if isinstance(out, dict):
        for k in keys:
            v = maybe_float(out.get(k))
            if v is not None:
                return v
        if isinstance(out.get("outliers"), list) and out.get("outliers"):
            return maybe_float(out["outliers"][0])
    for k in keys:
        v = maybe_float(prof.get(k))
        if v is not None:
            return v
    return None



def compute_outlier_metrics_direct(
    gt_tables: List[Dict[str, Any]],
    pred_tables: List[Dict[str, Any]],
    table_matches: List[TableMatch],
    col_matches_by_table: Dict[Tuple[int, int], List[ColumnMatch]],
) -> Dict[str, float]:
    rate_scores: List[float] = []
    count_scores: List[float] = []
    top1_total = 0
    top1_scored = 0
    top1_correct = 0
    n_numeric_gt_cols = 0
    n_numeric_pred_cols_with_outlier = 0

    numeric_dtypes = {"int", "integer", "float", "double", "number", "numeric"}

    for tm in table_matches:
        gt_t = gt_tables[tm.gt_idx]
        pred_t = pred_tables[tm.pred_idx]
        for cm in col_matches_by_table.get((tm.gt_idx, tm.pred_idx), []):
            gt_c = (gt_t.get("columns", []) or [])[cm.gt_idx]
            pred_c = (pred_t.get("columns", []) or [])[cm.pred_idx]
            dtype = norm_compact(gt_c.get("dtype"))
            if dtype not in numeric_dtypes:
                continue
            n_numeric_gt_cols += 1

            gt_prof = gt_c.get("profile", {}) or {}
            pred_prof = pred_c.get("profile", {}) or {}
            if not isinstance(gt_prof, dict):
                gt_prof = {}
            if not isinstance(pred_prof, dict):
                pred_prof = {}

            gt_out = gt_prof.get("outlier") if isinstance(gt_prof.get("outlier"), dict) else {}
            pred_out = pred_prof.get("outlier") if isinstance(pred_prof.get("outlier"), dict) else {}
            gt_rate = maybe_float(gt_out.get("outlier_rate", gt_prof.get("outlier_rate")))
            pred_rate = maybe_float(pred_out.get("outlier_rate", pred_prof.get("outlier_rate")))
            s = _score_abs_diff(pred_rate, gt_rate, 0.02)
            if s is not None:
                rate_scores.append(s)

            gt_n = maybe_float(gt_out.get("n_outliers", gt_prof.get("n_outliers")))
            pred_n = maybe_float(pred_out.get("n_outliers", pred_prof.get("n_outliers")))
            s = _score_rel_diff(pred_n, gt_n, 0.25)
            if s is not None:
                count_scores.append(s)

            if pred_rate is not None or pred_n is not None or pred_out:
                n_numeric_pred_cols_with_outlier += 1

            gt_top1 = _extract_top1_outlier_value(gt_prof)
            pred_top1 = _extract_top1_outlier_value(pred_prof)
            if gt_top1 is not None:
                top1_total += 1
                if pred_top1 is not None:
                    top1_scored += 1
                    tol = 1e-6 + 1e-6 * max(1.0, abs(gt_top1))
                    if abs(pred_top1 - gt_top1) <= tol:
                        top1_correct += 1

    top1_acc = safe_div(top1_correct, top1_total)
    overall = _avg([x for x in [_avg(rate_scores), _avg(count_scores), top1_acc] if True])
    return {
        "outlier_rate_score": _avg(rate_scores),
        "outlier_count_score": _avg(count_scores),
        "outlier_top1_accuracy": top1_acc,
        "outlier_overall_score": overall,
        "outlier_n_numeric_cols_gt": float(n_numeric_gt_cols),
        "outlier_n_numeric_cols_with_pred": float(n_numeric_pred_cols_with_outlier),
        "outlier_top1_n_scored": float(top1_scored),
        "outlier_top1_n_correct": float(top1_correct),
    }


# ------------------------- source_ref metrics -------------------------


def compute_source_ref_metrics(
    gt_tables: List[Dict[str, Any]],
    pred_tables: List[Dict[str, Any]],
    table_matches: List[TableMatch],
) -> Dict[str, float]:
    field_hits: Dict[str, int] = {k: 0 for k in ["sheet_name", "anchor_text", "block_id", "extraction_kind", "cell_region"]}
    field_den: Dict[str, int] = {k: 0 for k in ["sheet_name", "anchor_text", "block_id", "extraction_kind", "cell_region"]}
    table_all_hit = 0
    soft_scores: List[float] = []

    for tm in table_matches:
        gt_src = gt_tables[tm.gt_idx].get("source_ref") or {}
        pred_src = pred_tables[tm.pred_idx].get("source_ref") or {}
        if not isinstance(gt_src, dict):
            gt_src = {}
        if not isinstance(pred_src, dict):
            pred_src = {}

        all_ok = True
        any_field = False
        for field in field_hits:
            gt_val = gt_src.get(field)
            pred_val = pred_src.get(field)
            if gt_val is None or gt_val == "":
                continue
            any_field = True
            field_den[field] += 1
            sim = source_ref_field_similarity(gt_src, pred_src, field)
            if sim is None:
                all_ok = False
                continue
            if field in {"anchor_text"}:
                ok = sim >= 0.90
            elif field in {"cell_region"}:
                ok = sim >= 0.95
            else:
                ok = exactish_equal(gt_val, pred_val)
            if ok:
                field_hits[field] += 1
            else:
                all_ok = False
        sim = source_ref_similarity(gt_tables[tm.gt_idx], pred_tables[tm.pred_idx])
        if sim is not None:
            soft_scores.append(sim)
        if any_field and all_ok:
            table_all_hit += 1

    out = {
        f"source_ref_{field}_accuracy": safe_div(field_hits[field], field_den[field])
        for field in field_hits
    }
    out["source_ref_all_fields_accuracy"] = safe_div(table_all_hit, len(table_matches))
    out["source_ref_similarity_mean"] = _avg(soft_scores)
    return out


# ------------------------- top-level evaluation -------------------------


def eval_stage0_xlsx_source_ref(
    gt: Dict[str, Any],
    pred: Dict[str, Any],
    summary_mode: str = "lexical",
) -> Dict[str, Any]:
    gt_tables = gt.get("tables", []) or []
    pred_tables = pred.get("tables", []) or []

    table_matches, gt2pred, pred2gt, sim_matrix = match_tables(gt_tables, pred_tables, threshold=0.58)
    n_gt_tables = len(gt_tables)
    n_pred_tables = len(pred_tables)
    n_matched_tables = len(table_matches)

    table_precision = safe_div(n_matched_tables, n_pred_tables)
    table_recall = safe_div(n_matched_tables, n_gt_tables)
    table_f1 = safe_div(2 * table_precision * table_recall, table_precision + table_recall) if (table_precision + table_recall) > 0 else 0.0

    table_logical_hits = 0
    rowcount_hits = 0
    table_sim_sum = 0.0

    total_gt_cols = 0
    total_pred_cols = 0
    total_matched_cols = 0
    dtype_hits = 0
    role_hits = 0
    col_logical_hits = 0
    total_col_pairs = 0
    total_role_pairs = 0
    total_col_logical_pairs = 0

    col_maps_by_table: Dict[Tuple[int, int], Dict[int, int]] = {}
    col_matches_by_table: Dict[Tuple[int, int], List[ColumnMatch]] = {}
    alignment_debug: List[Dict[str, Any]] = []

    matched_gt_table_indices = set()
    matched_pred_table_indices = set()

    for tm in table_matches:
        gt_t = gt_tables[tm.gt_idx]
        pred_t = pred_tables[tm.pred_idx]
        matched_gt_table_indices.add(tm.gt_idx)
        matched_pred_table_indices.add(tm.pred_idx)
        table_sim_sum += tm.sim

        if exactish_equal(gt_t.get("logical_name"), pred_t.get("logical_name")) and gt_t.get("logical_name"):
            table_logical_hits += 1
        if gt_t.get("n_rows") == pred_t.get("n_rows"):
            rowcount_hits += 1

        matches, gt2pred_col, pred2gt_col = match_columns(gt_t, pred_t, threshold=0.55)
        col_maps_by_table[(tm.gt_idx, tm.pred_idx)] = pred2gt_col
        col_matches_by_table[(tm.gt_idx, tm.pred_idx)] = matches

        gt_cols = gt_t.get("columns", []) or []
        pred_cols = pred_t.get("columns", []) or []
        total_gt_cols += len(gt_cols)
        total_pred_cols += len(pred_cols)
        total_matched_cols += len(matches)

        matched_gt_col_indices = set()
        matched_pred_col_indices = set()

        for cm in matches:
            gc = gt_cols[cm.gt_idx]
            pc = pred_cols[cm.pred_idx]
            matched_gt_col_indices.add(cm.gt_idx)
            matched_pred_col_indices.add(cm.pred_idx)
            total_col_pairs += 1
            total_role_pairs += 1
            if gc.get("logical_name"):
                total_col_logical_pairs += 1
            if exactish_equal(gc.get("dtype"), pc.get("dtype")) and gc.get("dtype"):
                dtype_hits += 1
            if exactish_equal(gc.get("role"), pc.get("role")) and gc.get("role"):
                role_hits += 1
            if gc.get("logical_name") and exactish_equal(gc.get("logical_name"), pc.get("logical_name")):
                col_logical_hits += 1

        alignment_debug.append({
            "gt_table_id": gt_t.get("table_id"),
            "gt_physical_name": gt_t.get("physical_name"),
            "pred_table_id": pred_t.get("table_id"),
            "pred_physical_name": pred_t.get("physical_name"),
            "similarity": round(tm.sim, 4),
            "n_gt_columns": len(gt_cols),
            "n_pred_columns": len(pred_cols),
            "n_matched_columns": len(matches),
            "unmatched_gt_columns": [gt_cols[i].get("physical_name") for i in range(len(gt_cols)) if i not in matched_gt_col_indices],
            "unmatched_pred_columns": [pred_cols[i].get("physical_name") for i in range(len(pred_cols)) if i not in matched_pred_col_indices],
        })

    col_precision = safe_div(total_matched_cols, total_pred_cols)
    col_recall = safe_div(total_matched_cols, total_gt_cols)
    col_f1 = safe_div(2 * col_precision * col_recall, col_precision + col_recall) if (col_precision + col_recall) > 0 else 0.0

    pred_rel_set = build_pred_relation_set_projected_to_gt(gt, pred, pred2gt, col_maps_by_table)
    gt_rel_set = build_gt_relation_set(gt)
    rel_inter = gt_rel_set & pred_rel_set
    rel_precision = safe_div(len(rel_inter), len(pred_rel_set))
    rel_recall = safe_div(len(rel_inter), len(gt_rel_set))
    rel_f1 = safe_div(2 * rel_precision * rel_recall, rel_precision + rel_recall) if (rel_precision + rel_recall) > 0 else 0.0

    metrics: Dict[str, Any] = {
        "table_precision": table_precision,
        "table_recall": table_recall,
        "table_f1": table_f1,
        "table_logical_accuracy": safe_div(table_logical_hits, n_matched_tables),
        "table_match_similarity_mean": safe_div(table_sim_sum, n_matched_tables),
        "rowcount_accuracy": safe_div(rowcount_hits, n_matched_tables),
        "n_gt_tables": float(n_gt_tables),
        "n_pred_tables": float(n_pred_tables),
        "n_matched_tables": float(n_matched_tables),
        "column_precision": col_precision,
        "column_recall": col_recall,
        "column_f1": col_f1,
        "dtype_accuracy": safe_div(dtype_hits, total_col_pairs),
        "role_accuracy": safe_div(role_hits, total_role_pairs),
        "column_logical_accuracy": safe_div(col_logical_hits, total_col_logical_pairs),
        "n_gt_columns": float(total_gt_cols),
        "n_pred_columns": float(total_pred_cols),
        "n_matched_columns": float(total_matched_cols),
        "relation_precision": rel_precision,
        "relation_recall": rel_recall,
        "relation_f1": rel_f1,
        "n_gt_relations": float(len(gt_rel_set)),
        "n_pred_relations": float(len(pred_rel_set)),
        "n_matched_relations": float(len(rel_inter)),
        "unmatched_gt_tables": [
            {"table_id": gt_tables[i].get("table_id"), "physical_name": gt_tables[i].get("physical_name")}
            for i in range(len(gt_tables)) if i not in matched_gt_table_indices
        ],
        "unmatched_pred_tables": [
            {"table_id": pred_tables[i].get("table_id"), "physical_name": pred_tables[i].get("physical_name")}
            for i in range(len(pred_tables)) if i not in matched_pred_table_indices
        ],
        "table_alignment": alignment_debug,
    }

    metrics.update(compute_source_ref_metrics(gt_tables, pred_tables, table_matches))
    metrics.update(compute_profile_metrics_direct(gt_tables, pred_tables, table_matches, col_matches_by_table))
    metrics.update(compute_outlier_metrics_direct(gt_tables, pred_tables, table_matches, col_matches_by_table))
    metrics.update(compute_summary_length_metrics(gt, pred))

    if summary_mode in ("lexical", "both"):
        metrics.update(compute_dataset_summary_lexical(gt, pred))
    if summary_mode in ("llm", "both"):
        if HAS_LLM_EVAL and eval_dataset_summary_llm is not None:
            try:
                llm_metrics = eval_dataset_summary_llm(gt, pred)  # type: ignore[arg-type]
                if isinstance(llm_metrics, dict):
                    metrics.update(llm_metrics)
            except Exception as e:
                print(f"[WARN] LLM summary evaluation failed: {e}", file=sys.stderr)
                metrics.setdefault("summary_llm_error", 1.0)
        else:
            print("[WARN] LLM summary evaluation module not available.", file=sys.stderr)
            metrics.setdefault("summary_llm_error", 1.0)

    weighted_keys = [
        ("table_f1", 1.0),
        ("column_f1", 1.0),
        ("dtype_accuracy", 1.0),
        ("role_accuracy", 1.0),
        ("relation_f1", 1.0),
        ("table_logical_accuracy", 0.5),
        ("column_logical_accuracy", 0.5),
        ("rowcount_accuracy", 0.4),
        ("source_ref_similarity_mean", 0.7),
        ("source_ref_all_fields_accuracy", 0.3),
        ("profiling_overall_score", 0.3),
        ("outlier_overall_score", 0.2),
    ]
    num = 0.0
    den = 0.0
    for key, w in weighted_keys:
        # Skip relation scoring in the aggregate when the dataset has no GT relations.
        if key == "relation_f1" and float(metrics.get("n_gt_relations", 0.0)) == 0.0:
            continue
        # Skip outlier scoring in the aggregate when there were no comparable outlier
        # annotations to score. This avoids punishing datasets whose GT intentionally
        # leaves outlier fields null.
        if key == "outlier_overall_score" and float(metrics.get("outlier_top1_n_scored", 0.0)) == 0.0:
            continue
        if key in metrics:
            num += float(metrics[key]) * w
            den += w
    overall_raw = safe_div(num, den)

    # summary scoring:
# - if summary_mode in {"llm", "both"} and summary_llm_score is available,
#   use LLM score for the summary part of total_score
# - otherwise fall back to lexical summary score
# - if summary_mode == "none", do not penalize by summary

    summary_score_source = "none"
    summary_score_used = 1.0

    if summary_mode == "none":
        summary_score_source = "none"
        summary_score_used = 1.0

    elif summary_mode in ("llm", "both"):
        llm_score = metrics.get("summary_llm_score", None)
        if llm_score is not None:
            summary_score_source = "llm"
            summary_score_used = float(llm_score)
        else:
            # fallback to lexical if LLM eval failed / unavailable
            summary_len_score = float(metrics.get("summary_length_score", 1.0))
            summary_f1 = float(metrics.get("summary_f1", 1.0))
            summary_score_source = "lexical_fallback"
            summary_score_used = 0.5 * summary_len_score + 0.5 * summary_f1

    elif summary_mode == "lexical":
        summary_len_score = float(metrics.get("summary_length_score", 1.0))
        summary_f1 = float(metrics.get("summary_f1", 1.0))
        summary_score_source = "lexical"
        summary_score_used = 0.5 * summary_len_score + 0.5 * summary_f1

    summary_score_used = max(0.0, min(1.0, float(summary_score_used)))

    # keep the same penalty strength as before:
    # summary only affects up to 30% penalty
    summary_factor = 1.0 - 0.3 * (1.0 - summary_score_used)
    summary_factor = max(0.0, min(1.0, summary_factor))

    overall_penalized = overall_raw * summary_factor

    metrics["summary_score_source"] = summary_score_source
    metrics["summary_score_used"] = summary_score_used
    metrics["summary_penalty_factor"] = summary_factor

    metrics["overall_struct_score_raw"] = overall_raw
    metrics["overall_struct_score"] = overall_penalized
    metrics["total_score"] = round(overall_penalized * 100.0, 4)
    metrics["max_score"] = 100.0
    metrics["fraction"] = overall_penalized
    return metrics


# ------------------------- CLI -------------------------


def main() -> None:
    parser = argparse.ArgumentParser(description="Stage-0 evaluator for messy workbook/XLSX datasets with source_ref-aware matching.")
    parser.add_argument("--gt", required=True, help="Path to source_ref-based stage0_gt.json")
    parser.add_argument("--pred", required=True, help="Path to predicted JSON")
    parser.add_argument("--summary-mode", choices=["none", "lexical", "llm", "both"], default="lexical")
    parser.add_argument("--pretty", action="store_true")
    args = parser.parse_args()

    gt = load_json(args.gt)
    pred = load_json(args.pred)
    metrics = eval_stage0_xlsx_source_ref(gt, pred, summary_mode=args.summary_mode)

    if args.pretty:
        print(json.dumps(metrics, indent=2, ensure_ascii=False))
    else:
        for k, v in metrics.items():
            if isinstance(v, float):
                print(f"{k}: {v:.4f}")
            else:
                print(f"{k}: {v}")


if __name__ == "__main__":
    main()
