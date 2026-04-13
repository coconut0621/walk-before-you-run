#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Model evaluation pipeline for XLSX Stage-0 tasks.

What this version fixes relative to the older script:
1) Uses the Responses API (recommended for reasoning models).
2) Supports reasoning effort via --reasoning.
3) Uses strict JSON Schema structured outputs instead of old json_object mode.
4) Records per-task and aggregate token usage.
5) Lets you choose whether to read values, formulas, or both from XLSX.
6) Saves raw/model metadata on failures for easier debugging.

Directory layout expected:
  root/
    prompt_xlsx.md
    sample_xlsx.json
    01/
      stage0_gt.json
      some_workbook.xlsx
    02/
      stage0_gt.json
      another_workbook.xlsx

Outputs:
  <task>/result/<name>.json
  <task>/result/<name>.meta.json
  root/run_summary_<name>.json
"""

API_KEY = ""  # leave blank to use OPENAI_API_KEY from the environment

import argparse
import json
import os
import re
import sys
import textwrap
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required. Install with: pip install -U openpyxl")

try:
    from openai import OpenAI
except ImportError:
    sys.exit("openai>=1.x is required. Install with: pip install -U openai")

try:
    import jsonschema
except ImportError:
    jsonschema = None

DEFAULT_MODEL = "gpt-5.4"
DEFAULT_NAME = "gpt"
PROMPT_FILENAME = "prompt_xlsx.md"
SAMPLE_FILENAME = "sample_xlsx.json"
GT_FILENAME = "stage0_gt.json"
RESULT_DIRNAME = "result"

MAX_ROWS_PER_SHEET = 300
MAX_COLS_PER_SHEET = 80
MAX_OUTPUT_TOKENS = 25_000
TEMPERATURE = None
MAX_RETRIES = 4
RETRY_DELAY_S = 6

SKIP_DIRS = {
    "xlsx_stage0_scores",
    "__pycache__",
    ".git",
    ".venv",
    "venv",
    "node_modules",
}

VALID_REASONING = {"none", "minimal", "low", "medium", "high", "xhigh"}
VALID_XLSX_MODES = {"values", "formulas", "both"}
VALID_REASONING_SUMMARY = {"none", "auto", "concise", "detailed"}


def _to_plain(obj: Any) -> Any:
    """Best-effort conversion from SDK/Pydantic objects to plain Python objects."""
    if obj is None:
        return None
    if isinstance(obj, (str, int, float, bool)):
        return obj
    if isinstance(obj, list):
        return [_to_plain(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _to_plain(v) for k, v in obj.items()}
    if hasattr(obj, "model_dump"):
        return obj.model_dump()
    if hasattr(obj, "dict"):
        return obj.dict()
    if hasattr(obj, "__dict__"):
        return {k: _to_plain(v) for k, v in vars(obj).items() if not k.startswith("_")}
    return obj


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


# ══════════════════════════════════════════════════════════════════════════════
# JSON SCHEMA FOR STRICT STRUCTURED OUTPUTS
# ══════════════════════════════════════════════════════════════════════════════

def scalar_schema() -> Dict[str, Any]:
    return {"type": ["string", "number", "integer", "boolean", "null"]}


def nullable_string() -> Dict[str, Any]:
    return {"type": ["string", "null"]}


def nullable_number() -> Dict[str, Any]:
    return {"type": ["number", "null"]}


def nullable_boolean() -> Dict[str, Any]:
    return {"type": ["boolean", "null"]}


def strict_stage0_schema() -> Dict[str, Any]:
    outlier_schema = {
        "type": ["object", "null"],
        "properties": {
            "outlier_rate": nullable_number(),
            "n_outliers": {"type": ["integer", "null"]},
            "top_1_outlier": nullable_number(),
        },
        "required": ["outlier_rate", "n_outliers", "top_1_outlier"],
        "additionalProperties": False,
    }

    column_profile_schema = {
        "type": "object",
        "properties": {
            "parse_success_rate": nullable_number(),
            "type_consistency": nullable_number(),
            "min": scalar_schema(),
            "max": scalar_schema(),
            "mean": nullable_number(),
            "median": nullable_number(),
            "format_consistency": nullable_number(),
            "format_hint": nullable_string(),
            "outlier": outlier_schema,
        },
        "required": [
            "parse_success_rate",
            "type_consistency",
            "min",
            "max",
            "mean",
            "median",
            "format_consistency",
            "format_hint",
            "outlier",
        ],
        "additionalProperties": False,
    }

    dictionary_schema = {
        "type": "object",
        "properties": {
            "plain_description": {"type": "string"},
            "unit": {"type": "string"},
            "value_type": {
                "type": "string",
                "enum": [
                    "id",
                    "continuous",
                    "categorical",
                    "integer",
                    "text",
                    "timestamp",
                    "boolean",
                    "other",
                ],
            },
            "expected_range": {
                "anyOf": [
                    {"type": "null"},
                    {
                        "type": "array",
                        "items": scalar_schema(),
                        "minItems": 2,
                        "maxItems": 2,
                    },
                ]
            },
            "higher_is_better": nullable_boolean(),
            "missing_value_meaning": nullable_string(),
            "allowed_categories": {
                "anyOf": [
                    {"type": "null"},
                    {"type": "array", "items": scalar_schema()},
                ]
            },
        },
        "required": [
            "plain_description",
            "unit",
            "value_type",
            "expected_range",
            "higher_is_better",
            "missing_value_meaning",
            "allowed_categories",
        ],
        "additionalProperties": False,
    }

    column_schema = {
        "type": "object",
        "properties": {
            "physical_name": {"type": "string"},
            "logical_name": {"type": "string"},
            "dtype": {
                "type": "string",
                "enum": ["integer", "float", "boolean", "string", "date"],
            },
            "role": {
                "type": "string",
                "enum": [
                    "primary_key",
                    "foreign_key",
                    "numeric_feature",
                    "categorical_feature",
                    "timestamp",
                    "id",
                    "target",
                    "free_text",
                    "other",
                ],
            },
            "is_unique": {"type": "boolean"},
            "is_nullable": {"type": "boolean"},
            "dictionary": dictionary_schema,
            "profile": column_profile_schema,
        },
        "required": [
            "physical_name",
            "logical_name",
            "dtype",
            "role",
            "is_unique",
            "is_nullable",
            "dictionary",
            "profile",
        ],
        "additionalProperties": False,
    }

    source_ref_schema = {
        "type": "object",
        "properties": {
            "sheet_name": nullable_string(),
            "anchor_text": nullable_string(),
            "block_id": nullable_string(),
            "extraction_kind": nullable_string(),
            "cell_region": nullable_string(),
        },
        "required": [
            "sheet_name",
            "anchor_text",
            "block_id",
            "extraction_kind",
            "cell_region",
        ],
        "additionalProperties": False,
    }

    table_schema = {
        "type": "object",
        "properties": {
            "table_id": {"type": "string"},
            "physical_name": {"type": "string"},
            "logical_name": {"type": "string"},
            "n_rows": {"type": "integer"},
            "profile": {
                "type": "object",
                "properties": {
                    "density": nullable_number(),
                },
                "required": ["density"],
                "additionalProperties": False,
            },
            "columns": {"type": "array", "items": column_schema},
            "source_ref": source_ref_schema,
        },
        "required": [
            "table_id",
            "physical_name",
            "logical_name",
            "n_rows",
            "profile",
            "columns",
            "source_ref",
        ],
        "additionalProperties": False,
    }

    relation_schema = {
        "type": "object",
        "properties": {
            "relation_id": {"type": "string"},
            "parent_table_id": {"type": "string"},
            "child_table_id": {"type": "string"},
            "parent_key": {"type": "array", "items": {"type": "string"}},
            "child_key": {"type": "array", "items": {"type": "string"}},
        },
        "required": [
            "relation_id",
            "parent_table_id",
            "child_table_id",
            "parent_key",
            "child_key",
        ],
        "additionalProperties": False,
    }

    dataset_summary_schema = {
        "type": "object",
        "properties": {
            "plain_description": {"type": "string"},
            "n_tables": {"type": "integer"},
            "notes": nullable_string(),
        },
        "required": ["plain_description", "n_tables", "notes"],
        "additionalProperties": False,
    }

    return {
        "type": "object",
        "properties": {
            "tables": {"type": "array", "items": table_schema},
            "relations": {"type": "array", "items": relation_schema},
            "dataset_summary": dataset_summary_schema,
        },
        "required": ["tables", "relations", "dataset_summary"],
        "additionalProperties": False,
    }


# ══════════════════════════════════════════════════════════════════════════════
# XLSX → TEXT
# ══════════════════════════════════════════════════════════════════════════════

def _safe_repr(value: Any) -> str:
    text = repr(value)
    if len(text) > 240:
        text = text[:237] + "..."
    return text


def _load_workbook_pair(xlsx_path: Path) -> Tuple[Any, Any]:
    wb_formulas = openpyxl.load_workbook(str(xlsx_path), data_only=False)
    wb_values = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    return wb_formulas, wb_values


def _cell_repr(cell_formula: Any, cell_value: Any, mode: str) -> Optional[str]:
    formula_val = cell_formula.value
    value_val = cell_value.value

    if mode == "values":
        if value_val is None:
            return None
        return _safe_repr(value_val)

    if mode == "formulas":
        if formula_val is None:
            return None
        return _safe_repr(formula_val)

    # both
    if formula_val is None and value_val is None:
        return None

    is_formula = getattr(cell_formula, "data_type", None) == "f" or (
        isinstance(formula_val, str) and formula_val.startswith("=")
    )

    if is_formula:
        return _safe_repr({"formula": formula_val, "value": value_val})

    if value_val is not None:
        return _safe_repr(value_val)

    return _safe_repr(formula_val)


def xlsx_to_text(
    xlsx_path: Path,
    *,
    max_rows: int,
    max_cols: int,
    mode: str,
) -> str:
    """Serialize an XLSX workbook to an LLM-readable text block."""
    try:
        wb_formulas, wb_values = _load_workbook_pair(xlsx_path)
    except Exception as exc:
        return f"[ERROR reading {xlsx_path.name}: {exc}]"

    lines: List[str] = []
    lines.append(f"FILE: {xlsx_path.name}")
    lines.append(f"READ_MODE: {mode}")
    lines.append(f"SHEETS: {', '.join(wb_formulas.sheetnames)}")
    lines.append("")

    for sheet_name in wb_formulas.sheetnames:
        ws_f = wb_formulas[sheet_name]
        ws_v = wb_values[sheet_name]

        actual_rows = min(ws_f.max_row or 0, max_rows)
        actual_cols = min(ws_f.max_column or 0, max_cols)
        merged_ranges = list(getattr(ws_f, "merged_cells", []).ranges)

        lines.append(
            f"=== SHEET: {sheet_name!r} "
            f"(max_row={ws_f.max_row}, max_col={ws_f.max_column}, "
            f"merged_ranges={len(merged_ranges)}) ==="
        )
        if merged_ranges:
            preview = ", ".join(str(rng) for rng in merged_ranges[:12])
            if len(merged_ranges) > 12:
                preview += ", ..."
            lines.append(f"  MERGED_RANGES: {preview}")

        for r in range(1, actual_rows + 1):
            row_parts: List[str] = []
            for c in range(1, actual_cols + 1):
                cell_f = ws_f.cell(row=r, column=c)
                cell_v = ws_v.cell(row=r, column=c)
                rendered = _cell_repr(cell_f, cell_v, mode)
                if rendered is None:
                    continue
                coord = f"{get_column_letter(c)}{r}"
                row_parts.append(f"{coord}={rendered}")
            if row_parts:
                lines.append("  " + ", ".join(row_parts))

        truncated_r = (ws_f.max_row or 0) > max_rows
        truncated_c = (ws_f.max_column or 0) > max_cols
        if truncated_r or truncated_c:
            lines.append(
                f"  [truncated: showing {actual_rows}/{ws_f.max_row} rows, "
                f"{actual_cols}/{ws_f.max_column} cols]"
            )
        lines.append("")

    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# PROMPT BUILDER
# ══════════════════════════════════════════════════════════════════════════════

def build_input_messages(
    task_dir: Path,
    *,
    system_prompt: str,
    sample_json_text: str,
    max_rows: int,
    max_cols: int,
    xlsx_mode: str,
) -> List[Dict[str, str]]:
    xlsx_files = sorted(task_dir.glob("*.xlsx"))
    if not xlsx_files:
        raise FileNotFoundError(f"No xlsx files found in {task_dir}")

    user_parts: List[str] = []
    user_parts.append("## Workbook content\n")
    for xlsx_path in xlsx_files:
        user_parts.append(
            xlsx_to_text(
                xlsx_path,
                max_rows=max_rows,
                max_cols=max_cols,
                mode=xlsx_mode,
            )
        )

    user_parts.append("## Output format reference (sample_xlsx.json)\n")
    user_parts.append("Use only as a format template. Do NOT copy its values.\n")
    user_parts.append(sample_json_text)

    user_parts.append(
        "\n## Task\n"
        "Analyze the workbook content above and produce the stage-0 JSON exactly "
        "as described in the system instructions. Output only the final JSON object."
    )

    return [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": "\n".join(user_parts)},
    ]


# ══════════════════════════════════════════════════════════════════════════════
# RESPONSES API CALL
# ══════════════════════════════════════════════════════════════════════════════

def should_retry(exc: Exception) -> bool:
    name = exc.__class__.__name__.lower()
    status = getattr(exc, "status_code", None)
    if status in {408, 409, 429, 500, 502, 503, 504}:
        return True
    retry_markers = [
        "ratelimit",
        "timeout",
        "connection",
        "servererror",
        "internalservererror",
        "apierror",
    ]
    return any(marker in name for marker in retry_markers)


def call_openai_responses(
    client: OpenAI,
    *,
    model: str,
    input_messages: List[Dict[str, str]],
    schema: Dict[str, Any],
    reasoning: Optional[str],
    reasoning_summary: str,
    temperature: Optional[float],
    max_output_tokens: int,
    max_retries: int,
    retry_delay: float,
) -> Tuple[str, Dict[str, Any]]:
    reasoning_payload: Optional[Dict[str, str]] = None
    if reasoning:
        reasoning_payload = {"effort": reasoning}
        if reasoning_summary != "none":
            reasoning_payload["summary"] = reasoning_summary

    last_exc: Optional[Exception] = None
    for attempt in range(1, max_retries + 1):
        try:
            request_kwargs = {
                "model": model,
                "input": input_messages,
                "reasoning": reasoning_payload,
                "max_output_tokens": max_output_tokens,
                "text": {
                    "format": {
                        "type": "json_schema",
                        "name": "stage0_xlsx_schema",
                        "schema": schema,
                        "strict": True,
                    }
                },
            }

            # GPT-5 / reasoning families reject explicit temperature.
            if temperature is not None and not model.startswith(("gpt-5", "o1", "o3", "o4")):
                request_kwargs["temperature"] = temperature

            response = client.responses.create(**request_kwargs)

            raw_text = getattr(response, "output_text", None)
            if not raw_text:
                response_plain = _to_plain(response)
                raw_text = extract_text_from_response_payload(response_plain)

            meta = build_response_meta(response, reasoning_payload)
            return raw_text, meta

        except Exception as exc:
            last_exc = exc
            if attempt == max_retries or not should_retry(exc):
                raise
            wait = retry_delay * attempt
            print(f"\n  [retry] {exc.__class__.__name__}: waiting {wait}s before retry {attempt}/{max_retries}...", flush=True)
            time.sleep(wait)

    assert last_exc is not None
    raise last_exc


def extract_text_from_response_payload(payload: Dict[str, Any]) -> str:
    outputs = payload.get("output", []) or []
    text_chunks: List[str] = []
    for item in outputs:
        if item.get("type") != "message":
            continue
        for content_item in item.get("content", []) or []:
            if content_item.get("type") == "output_text":
                text_chunks.append(content_item.get("text", ""))
    return "\n".join(chunk for chunk in text_chunks if chunk)


def build_response_meta(response: Any, reasoning_payload: Optional[Dict[str, str]]) -> Dict[str, Any]:
    plain = _to_plain(response)
    usage = plain.get("usage") or {}
    input_details = usage.get("input_tokens_details") or {}
    output_details = usage.get("output_tokens_details") or {}
    return {
        "response_id": plain.get("id"),
        "model": plain.get("model"),
        "status": plain.get("status"),
        "reasoning_request": reasoning_payload,
        "reasoning_response": plain.get("reasoning"),
        "usage": usage,
        "input_tokens": usage.get("input_tokens", 0),
        "cached_input_tokens": input_details.get("cached_tokens", 0),
        "output_tokens": usage.get("output_tokens", 0),
        "reasoning_tokens": output_details.get("reasoning_tokens", 0),
        "total_tokens": usage.get("total_tokens", 0),
        "service_tier": plain.get("service_tier"),
    }


# ══════════════════════════════════════════════════════════════════════════════
# JSON EXTRACTION + VALIDATION
# ══════════════════════════════════════════════════════════════════════════════

def extract_json(raw: str) -> Dict[str, Any]:
    text = raw.strip()
    text = re.sub(r"^```[^\n]*\n", "", text)
    text = re.sub(r"\n?```\s*$", "", text)
    text = text.strip()
    return json.loads(text)


def validate_result_schema(result: Dict[str, Any], schema: Dict[str, Any]) -> None:
    if jsonschema is None:
        return
    jsonschema.validate(instance=result, schema=schema)


# ══════════════════════════════════════════════════════════════════════════════
# TASK DISCOVERY
# ══════════════════════════════════════════════════════════════════════════════

def find_task_dirs(root: Path) -> List[Path]:
    dirs: List[Path] = []
    for p in sorted(root.iterdir()):
        if not p.is_dir():
            continue
        if p.name in SKIP_DIRS or p.name.startswith("."):
            continue
        if not (p / GT_FILENAME).exists():
            continue
        dirs.append(p)
    return dirs


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run an OpenAI model over XLSX Stage-0 tasks.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent(
            """
            Examples:
              python run_model_pipeline.py
              python run_model_pipeline.py --model gpt-5.4 --reasoning low --name gpt54_low
              python run_model_pipeline.py --task 05 --task 10 --xlsx-mode both
              python run_model_pipeline.py --skip-existing
            """
        ),
    )
    parser.add_argument("--dir", default=".", help="Root directory containing task subfolders.")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"OpenAI model name. Default: {DEFAULT_MODEL}")
    parser.add_argument("--name", default=DEFAULT_NAME, help=f"Output filename stem. Default: {DEFAULT_NAME}")
    parser.add_argument("--api-key", default="", help="Optional API key. Overrides API_KEY/env if provided.")
    parser.add_argument("--task", metavar="TASK", action="append", dest="tasks", help="Run only this task folder. Can be repeated.")
    parser.add_argument("--skip-existing", action="store_true", help="Skip tasks whose result JSON already exists.")
    parser.add_argument("--reasoning", default="none", choices=sorted(VALID_REASONING), help="Reasoning effort for reasoning-capable models.")
    parser.add_argument("--reasoning-summary", default="none", choices=sorted(VALID_REASONING_SUMMARY), help="Optional reasoning summary to request from the model.")
    parser.add_argument("--xlsx-mode", default="both", choices=sorted(VALID_XLSX_MODES), help="Read workbook cells as values, formulas, or both.")
    parser.add_argument("--max-rows", type=int, default=MAX_ROWS_PER_SHEET, help=f"Max rows per sheet to serialize. Default: {MAX_ROWS_PER_SHEET}")
    parser.add_argument("--max-cols", type=int, default=MAX_COLS_PER_SHEET, help=f"Max cols per sheet to serialize. Default: {MAX_COLS_PER_SHEET}")
    parser.add_argument("--max-output-tokens", type=int, default=MAX_OUTPUT_TOKENS, help=f"Max output tokens. Default: {MAX_OUTPUT_TOKENS}")
    parser.add_argument("--temperature", type=float, default=TEMPERATURE, help="Sampling temperature for models that support it. Omit for GPT-5/o-series reasoning models.")
    parser.add_argument("--max-retries", type=int, default=MAX_RETRIES, help=f"Max transient retries. Default: {MAX_RETRIES}")
    parser.add_argument("--retry-delay", type=float, default=RETRY_DELAY_S, help=f"Base retry delay in seconds. Default: {RETRY_DELAY_S}")
    args = parser.parse_args()

    root = Path(args.dir).resolve()
    prompt_path = root / PROMPT_FILENAME
    sample_path = root / SAMPLE_FILENAME

    for p in (prompt_path, sample_path):
        if not p.exists():
            sys.exit(f"Required file not found: {p}")

    system_prompt = _read_text(prompt_path)
    sample_json_text = _read_text(sample_path)
    try:
        json.loads(sample_json_text)
    except json.JSONDecodeError as exc:
        sys.exit(f"sample_xlsx.json is not valid JSON: {exc}")

    api_key = args.api_key or API_KEY or os.environ.get("OPENAI_API_KEY", "")
    if not api_key:
        sys.exit(
            "No API key found. Either set API_KEY at the top of the script, "
            "pass --api-key, or export OPENAI_API_KEY."
        )

    client = OpenAI(api_key=api_key)
    schema = strict_stage0_schema()

    all_task_dirs = find_task_dirs(root)
    if not all_task_dirs:
        sys.exit(f"No task folders with {GT_FILENAME} found in {root}")

    if args.tasks:
        selected = set(args.tasks)
        task_dirs = [d for d in all_task_dirs if d.name in selected]
        missing = selected - {d.name for d in task_dirs}
        if missing:
            print(f"[WARNING] Task(s) not found: {', '.join(sorted(missing))}")
        if not task_dirs:
            sys.exit("No matching tasks found.")
    else:
        task_dirs = all_task_dirs

    print(f"Model          : {args.model}")
    print(f"Reasoning      : {args.reasoning}")
    print(f"Reasoning sum. : {args.reasoning_summary}")
    print(f"XLSX mode      : {args.xlsx_mode}")
    if args.temperature is None:
        print("Temperature    : <omitted>")
    else:
        print(f"Temperature    : {args.temperature}")
    print(f"Output         : <task>/result/{args.name}.json")
    print(f"Tasks          : {len(task_dirs)}")
    print()

    ok_count = 0
    skip_count = 0
    error_count = 0
    run_usage = {
        "input_tokens": 0,
        "cached_input_tokens": 0,
        "output_tokens": 0,
        "reasoning_tokens": 0,
        "total_tokens": 0,
    }
    task_summaries: List[Dict[str, Any]] = []

    for task_dir in task_dirs:
        task_name = task_dir.name
        result_dir = task_dir / RESULT_DIRNAME
        result_dir.mkdir(exist_ok=True)
        out_path = result_dir / f"{args.name}.json"
        #meta_path = result_dir / f"{args.name}.meta.json"
        raw_path = result_dir / f"{args.name}.raw.txt"
        err_path = result_dir / f"{args.name}.error.txt"

        if args.skip_existing and out_path.exists():
            print(f"[SKIP ] {task_name}: {out_path.name} already exists")
            skip_count += 1
            continue

        print(f"[RUN  ] {task_name}", end="  ", flush=True)

        try:
            input_messages = build_input_messages(
                task_dir,
                system_prompt=system_prompt,
                sample_json_text=sample_json_text,
                max_rows=args.max_rows,
                max_cols=args.max_cols,
                xlsx_mode=args.xlsx_mode,
            )
        except Exception as exc:
            print(f"ERROR building prompt: {exc}")
            err_path.write_text(f"build_input_messages error:\n{exc}\n", encoding="utf-8")
            error_count += 1
            task_summaries.append({"task": task_name, "status": "build_error", "error": str(exc)})
            continue

        total_chars = sum(len(m["content"]) for m in input_messages)
        print(f"(~{total_chars:,} chars)", end="  ", flush=True)

        try:
            raw_text, meta = call_openai_responses(
                client,
                model=args.model,
                input_messages=input_messages,
                schema=schema,
                reasoning=args.reasoning,
                reasoning_summary=args.reasoning_summary,
                temperature=args.temperature,
                max_output_tokens=args.max_output_tokens,
                max_retries=args.max_retries,
                retry_delay=args.retry_delay,
            )
        except Exception as exc:
            print(f"API ERROR: {exc}")
            err_path.write_text(str(exc), encoding="utf-8")
            error_count += 1
            task_summaries.append({"task": task_name, "status": "api_error", "error": str(exc)})
            continue

        try:
            result = extract_json(raw_text)
            validate_result_schema(result, schema)
        except Exception as exc:
            print(f"JSON/SCHEMA ERROR: {exc}")
            raw_path.write_text(raw_text, encoding="utf-8")
            err_path.write_text(str(exc), encoding="utf-8")
            #meta_path.write_text(json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8")
            error_count += 1
            task_summaries.append({"task": task_name, "status": "parse_or_schema_error", "error": str(exc), "meta": meta})
            continue

        out_path.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
        #meta_path.write_text(json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8")

        for k in run_usage:
            run_usage[k] += int(meta.get(k, 0) or 0)

        n_tables = len(result.get("tables", []))
        print(
            f"→ saved ({n_tables} tables, input={meta['input_tokens']}, "
            f"output={meta['output_tokens']}, reasoning={meta['reasoning_tokens']})"
        )
        ok_count += 1
        task_summaries.append({
            "task": task_name,
            "status": "ok",
            "tables": n_tables,
            "meta": meta,
            "result_path": str(out_path.relative_to(root)),
        })

    summary = {
        "model": args.model,
        "reasoning": args.reasoning,
        "reasoning_summary": args.reasoning_summary,
        "xlsx_mode": args.xlsx_mode,
        "max_rows": args.max_rows,
        "max_cols": args.max_cols,
        "max_output_tokens": args.max_output_tokens,
        "temperature": args.temperature,
        "ok": ok_count,
        "skipped": skip_count,
        "errors": error_count,
        "usage": run_usage,
        "tasks": task_summaries,
    }
    summary_path = root / f"run_summary_{args.name}.json"
    summary_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")

    print()
    print(f"Done. OK={ok_count} skipped={skip_count} errors={error_count}")
    print(
        "Usage: "
        f"input={run_usage['input_tokens']}, "
        f"cached_input={run_usage['cached_input_tokens']}, "
        f"output={run_usage['output_tokens']}, "
        f"reasoning={run_usage['reasoning_tokens']}, "
        f"total={run_usage['total_tokens']}"
    )
    print(f"Run summary saved to: {summary_path}")
    if ok_count:
        print(f"\nTo score results, run:\n  python bulk_score_xlsx_dir.py --dir {args.dir}")


if __name__ == "__main__":
    main()
