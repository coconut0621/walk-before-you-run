#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Bulk scorer for workbook/XLSX Stage-0 predictions across task folders.

Expected structure:
root/
  task_a/
    stage0_gt.json
    result/
      gemini.json
      claude.json
      gpt.json
  task_b/
    stage0_gt.json
    result/
      claude.json

Behavior:
- scan all first-level subdirectories under --dir
- treat each subdirectory as one task
- inside each task directory:
    - use its own stage0_gt.json
    - score all prediction *.json files inside the task's result/ subfolder
    - write task-local summary into:
        root/xlsx_stage0_scores/<task_name>/
- also write root-level aggregate summaries into:
        root/xlsx_stage0_scores/

Typical usage:
    python bulk_score_xlsx_dir.py
    python bulk_score_xlsx_dir.py --dir . --summary-mode lexical
"""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Any, Dict, List, Tuple

from evaluation_xlsx_source_ref import load_json, eval_stage0_xlsx_source_ref


DEFAULT_EXCLUDE_NAMES = {
    "sample.json",
    "sample_xlsx.json",
    "summary.json",
    "summary.csv",
    "all_tasks_summary.json",
    "all_tasks_summary.csv",
    "all_predictions_summary.csv",
}

SUMMARY_COLUMNS = [
    "file",
    "total_score",
    "max_score",
    "fraction",
    "overall_struct_score_raw",
    "overall_struct_score",
    "table_f1",
    "column_f1",
    "dtype_accuracy",
    "role_accuracy",
    "relation_f1",
    "rowcount_accuracy",
    "table_logical_accuracy",
    "column_logical_accuracy",
    "source_ref_similarity_mean",
    "source_ref_all_fields_accuracy",
    "source_ref_sheet_name_accuracy",
    "source_ref_anchor_text_accuracy",
    "source_ref_block_id_accuracy",
    "source_ref_extraction_kind_accuracy",
    "source_ref_cell_region_accuracy",
    "profiling_overall_score",
    "outlier_overall_score",
    "summary_f1",
    "summary_jaccard",
    "summary_length_score",
    "n_gt_tables",
    "n_pred_tables",
    "n_matched_tables",
    "n_gt_columns",
    "n_pred_columns",
    "n_matched_columns",
    "n_gt_relations",
    "n_pred_relations",
    "n_matched_relations",
]

ALL_PREDICTIONS_COLUMNS = ["task"] + SUMMARY_COLUMNS

TASK_SUMMARY_COLUMNS = [
    "task",
    "n_predictions",
    "mean_total_score",
    "mean_max_score",
    "mean_fraction",
    "best_total_score",
    "best_fraction",
    "best_file",
    "worst_total_score",
    "worst_fraction",
    "worst_file",
]


def round_if_float(x: Any) -> Any:
    if isinstance(x, float):
        return round(x, 6)
    return x


def is_prediction_json(path: Path, gt_path: Path, out_dir: Path) -> bool:
    if path.suffix.lower() != ".json":
        return False
    if path.resolve() == gt_path.resolve():
        return False
    if path.name in DEFAULT_EXCLUDE_NAMES:
        return False
    if path.name.startswith("stage0_gt"):
        return False
    if path.parent.resolve() == out_dir.resolve():
        return False
    if path.parent.name == "per_file":
        return False
    return True


def make_flat_summary_row(path: Path, metrics: Dict[str, Any]) -> Dict[str, Any]:
    row: Dict[str, Any] = {"file": path.name}
    for key in SUMMARY_COLUMNS:
        if key == "file":
            continue
        val = metrics.get(key)
        row[key] = round_if_float(val) if val is not None else ""
    return row


def add_task_to_row(task_name: str, row: Dict[str, Any]) -> Dict[str, Any]:
    out = {"task": task_name}
    for col in SUMMARY_COLUMNS:
        out[col] = row.get(col, "")
    return out


def write_csv(path: Path, fieldnames: List[str], rows: List[Dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def safe_float(x: Any, default: float = 0.0) -> float:
    try:
        if x == "" or x is None:
            return default
        return float(x)
    except Exception:
        return default


def summarize_task(task_name: str, rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    if not rows:
        return {
            "task": task_name,
            "n_predictions": 0,
            "mean_total_score": "",
            "mean_max_score": "",
            "mean_fraction": "",
            "best_total_score": "",
            "best_fraction": "",
            "best_file": "",
            "worst_total_score": "",
            "worst_fraction": "",
            "worst_file": "",
        }

    sorted_rows = sorted(
        rows,
        key=lambda r: (-safe_float(r.get("total_score")), str(r.get("file", "")))
    )
    best = sorted_rows[0]
    worst = sorted(
        rows,
        key=lambda r: (safe_float(r.get("total_score")), str(r.get("file", "")))
    )[0]

    n = len(rows)
    mean_total_score = sum(safe_float(r.get("total_score")) for r in rows) / n
    mean_max_score = sum(safe_float(r.get("max_score")) for r in rows) / n
    mean_fraction = sum(safe_float(r.get("fraction")) for r in rows) / n

    return {
        "task": task_name,
        "n_predictions": n,
        "mean_total_score": round(mean_total_score, 6),
        "mean_max_score": round(mean_max_score, 6),
        "mean_fraction": round(mean_fraction, 6),
        "best_total_score": round_if_float(best.get("total_score", "")),
        "best_fraction": round_if_float(best.get("fraction", "")),
        "best_file": best.get("file", ""),
        "worst_total_score": round_if_float(worst.get("total_score", "")),
        "worst_fraction": round_if_float(worst.get("fraction", "")),
        "worst_file": worst.get("file", ""),
    }


def find_task_dirs(root: Path, out_dir_name: str) -> List[Path]:
    task_dirs: List[Path] = []
    for p in sorted(root.iterdir()):
        if not p.is_dir():
            continue
        if p.name == out_dir_name:
            continue
        if p.name.startswith("."):
            continue
        task_dirs.append(p)
    return task_dirs


def score_one_task(
    task_dir: Path,
    out_root: Path,
    summary_mode: str,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    task_name = task_dir.name
    gt_path = task_dir / "stage0_gt.json"

    if not gt_path.exists():
        raise FileNotFoundError(f"[{task_name}] Missing GT file: {gt_path}")

    task_out_dir = out_root / task_name
    per_file_dir = task_out_dir / "per_file"
    task_out_dir.mkdir(parents=True, exist_ok=True)
    per_file_dir.mkdir(parents=True, exist_ok=True)

    gt = load_json(str(gt_path))

    result_dir = task_dir / "result"
    if not result_dir.is_dir():
        raise FileNotFoundError(
            f"[{task_name}] Missing result/ subfolder: {result_dir}"
        )

    pred_paths = sorted(
        p for p in result_dir.glob("*.json")
        if is_prediction_json(p, gt_path, out_root)
    )

    if not pred_paths:
        raise ValueError(
            f"[{task_name}] No prediction JSON files found in {result_dir}"
        )

    summary_rows: List[Dict[str, Any]] = []
    summary_payload: Dict[str, Any] = {
        "task": task_name,
        "gt": str(gt_path),
        "summary_mode": summary_mode,
        "n_predictions": 0,
        "files": {},
    }

    for pred_path in pred_paths:
        pred = load_json(str(pred_path))
        metrics = eval_stage0_xlsx_source_ref(gt, pred, summary_mode=summary_mode)

        row = make_flat_summary_row(pred_path, metrics)
        summary_rows.append(row)
        summary_payload["files"][pred_path.name] = metrics
        summary_payload["n_predictions"] += 1

        per_file_path = per_file_dir / f"{pred_path.stem}.metrics.json"
        per_file_path.write_text(
            json.dumps(metrics, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    summary_rows.sort(
        key=lambda r: (-safe_float(r.get("total_score")), str(r.get("file", "")))
    )

    summary_csv_path = task_out_dir / "summary.csv"
    summary_json_path = task_out_dir / "summary.json"

    write_csv(summary_csv_path, SUMMARY_COLUMNS, summary_rows)
    summary_json_path.write_text(
        json.dumps(summary_payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    return summary_rows, summary_payload


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Bulk-score workbook/XLSX Stage-0 predictions across task folders."
    )
    parser.add_argument(
        "--dir",
        default=".",
        help="Root directory containing task subfolders. Default: current directory",
    )
    parser.add_argument(
        "--summary-mode",
        choices=["none", "lexical", "llm", "both"],
        default="lexical",
    )
    parser.add_argument(
        "--outdir",
        default="xlsx_stage0_scores",
        help="Output directory under root. Default: xlsx_stage0_scores",
    )
    args = parser.parse_args()

    root = Path(args.dir).resolve()
    out_root = (root / args.outdir).resolve()
    out_root.mkdir(parents=True, exist_ok=True)

    task_dirs = find_task_dirs(root, args.outdir)
    if not task_dirs:
        raise SystemExit(f"No task subdirectories found in {root}")

    all_predictions_rows: List[Dict[str, Any]] = []
    all_tasks_rows: List[Dict[str, Any]] = []
    all_tasks_payload: Dict[str, Any] = {
        "root": str(root),
        "summary_mode": args.summary_mode,
        "n_tasks": 0,
        "tasks": {},
    }

    skipped_tasks: List[Dict[str, str]] = []

    for task_dir in task_dirs:
        task_name = task_dir.name
        try:
            summary_rows, summary_payload = score_one_task(
                task_dir=task_dir,
                out_root=out_root,
                summary_mode=args.summary_mode,
            )

            task_row = summarize_task(task_name, summary_rows)
            all_tasks_rows.append(task_row)

            for row in summary_rows:
                all_predictions_rows.append(add_task_to_row(task_name, row))

            all_tasks_payload["tasks"][task_name] = {
                "task_summary": task_row,
                "detail": summary_payload,
            }
            all_tasks_payload["n_tasks"] += 1

            print(f"[OK] {task_name}: scored {len(summary_rows)} prediction files")

        except Exception as e:
            skipped_tasks.append({"task": task_name, "reason": str(e)})
            print(f"[SKIP] {task_name}: {e}")

    if not all_tasks_rows:
        raise SystemExit("No task folders were successfully scored.")

    all_tasks_rows.sort(key=lambda r: str(r.get("task", "")))
    all_predictions_rows.sort(
        key=lambda r: (str(r.get("task", "")), -safe_float(r.get("total_score")), str(r.get("file", "")))
    )

    if skipped_tasks:
        all_tasks_payload["skipped_tasks"] = skipped_tasks

    all_tasks_summary_csv = out_root / "all_tasks_summary.csv"
    all_tasks_summary_json = out_root / "all_tasks_summary.json"
    all_predictions_summary_csv = out_root / "all_predictions_summary.csv"

    write_csv(all_tasks_summary_csv, TASK_SUMMARY_COLUMNS, all_tasks_rows)
    write_csv(all_predictions_summary_csv, ALL_PREDICTIONS_COLUMNS, all_predictions_rows)
    all_tasks_summary_json.write_text(
        json.dumps(all_tasks_payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print()
    print(f"Scored tasks: {all_tasks_payload['n_tasks']}")
    print(f"Task-level summary CSV: {all_tasks_summary_csv}")
    print(f"Task-level summary JSON: {all_tasks_summary_json}")
    print(f"All predictions summary CSV: {all_predictions_summary_csv}")
    print(f"Per-task outputs under: {out_root}")


if __name__ == "__main__":
    main()