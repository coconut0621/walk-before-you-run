# Walk Before You Run: The Importance of Data Exploration for Data Analysis Agents.

This repository contains the code, prompts, schemas, datasets, and evaluation pipeline for the paper:

**Walk Before You Run: The Importance of Data Exploration for Data Analysis Agents**

## Overview

We study **Data Exploration**, the pre-analysis step in which a tool must infer the logical structure of an uploaded workbook before answering downstream questions.

This repository includes:

- the Data Exploration prompting contract for workbook understanding
- the JSON schema template for TOBENAMED outputs
- an OpenAI-based pipeline for running workbook-to-JSON inference
- an automatic evaluator for TOBENAMED predictions
- a bulk scoring script for benchmarking predictions across task folders
- task folders with ground-truth metadata, workbook inputs, and test results

## Repository structure

```text
.
├── prompt_xlsx.md
├── sample_xlsx.json
├── run_model_pipeline.py
├── evaluation_xlsx_source_ref.py
├── eval_dataset_summary_llm.py
├── bulk_score_xlsx_dir.py
├── tasks/
│   ├── 01/
│   │   ├── stage0_gt.json
│   │   └── *.xlsx
│   │   │   ├── result
│   ├── 02/
│   │   ├── stage0_gt.json
│   │   └── *.xlsx
│   │   │   ├── result
│   └── ...


Each task folder should contain:

* one or more workbook files (`*.xlsx`)
* one `stage0_gt.json` file, which is the ground truth for the Data Exploration step

## Environment

We recommend Python 3.10+.

Install the required packages:

```bash
pip install openai openpyxl jsonschema
```

Notes:

* `openai` is required for running model inference and optional LLM-based summary scoring.
* `openpyxl` is required for reading Excel workbooks.
* `jsonschema` is used for validating structured JSON outputs.

## API key setup

`run_model_pipeline.py` uses the OpenAI API.

By default, the script reads the API key from the environment variable `OPENAI_API_KEY`.
This is the recommended setup.

### On Linux / WSL

```bash
export OPENAI_API_KEY="your_api_key_here"
```

### On Windows PowerShell

```powershell
$env:OPENAI_API_KEY="your_api_key_here"
```

You may also pass the key directly with:

```bash
python run_model_pipeline.py --api-key YOUR_API_KEY
```

but using environment variables is safer.

## Running the Stage-0 inference pipeline

The main testing script is:

```bash
python run_model_pipeline.py
```

Typical example:

```bash
python run_model_pipeline.py --model gpt-5.4 --reasoning medium --xlsx-mode both --name gpt
```

This script expects the root directory to contain:

* `prompt_xlsx.md`
* `sample_xlsx.json`
* task subfolders such as `01/`, `02/`, etc.

For each task, the script reads the workbook(s), builds the prompt, calls the model, and saves the predicted Stage-0 JSON to:

```text
<task>/result/<name>.json
```

Example:

```text
01/result/gpt.json
```

## Scoring predictions

After predictions are generated, run the bulk scorer:

```bash
python bulk_score_xlsx_dir.py --dir .
```

This script:

* scans all task folders
* uses each task’s own `stage0_gt.json`
* scores all prediction JSON files in each task’s `result/` folder
* writes aggregated outputs under:

```text
xlsx_stage0_scores/
```

You can also choose how dataset summaries are scored:

```bash
python bulk_score_xlsx_dir.py --dir . --summary-mode lexical
python bulk_score_xlsx_dir.py --dir . --summary-mode llm
python bulk_score_xlsx_dir.py --dir . --summary-mode both
```

## Main scripts

### `run_model_pipeline.py`

Runs an OpenAI model on workbook-based Stage-0 tasks using structured JSON outputs.

Current implementation is built for OpenAI models via the Responses API.

### `evaluation_xlsx_source_ref.py`

Evaluates Stage-0 predictions against ground truth with source-ref-aware matching, column/relation scoring, profiling checks, outlier checks, and summary scoring.

### `eval_dataset_summary_llm.py`

Optional LLM-as-judge module for evaluating `dataset_summary`.

### `bulk_score_xlsx_dir.py`

Bulk scoring script for running the evaluator across all task folders and collecting summary outputs.

## Using other tools or models

The provided inference pipeline (`run_model_pipeline.py`) is currently implemented for OpenAI models.

However, the evaluator is **model-agnostic** as long as a tool produces prediction files in the same Stage-0 JSON format.

To evaluate another tool without this pipeline, you can:

1. run that tool separately
2. save its prediction as a JSON file under each task’s `result/` folder
3. ensure the output matches the required schema
4. run `bulk_score_xlsx_dir.py`

For example:

```text
01/result/claude.json
01/result/gemini.json
```

The evaluator will score them in the same way.

## Notes

* `sample_xlsx.json` is a format template only, not a content template.
* `prompt_xlsx.md` defines the Stage-0 workbook understanding task.
* The evaluator assumes that predictions follow the same high-level JSON structure as `sample_xlsx.json`.

## Citation

If you use this repository, please cite the associated paper once bibliographic details are available.

