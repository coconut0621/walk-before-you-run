# Walk Before You Run: TOBENAMED Dataset Understanding Matters for Data Analysis Tools

This repository contains the code, prompts, schemas, datasets, and evaluation pipeline for the paper:

**Walk Before You Run: TOBENAMED Dataset Understanding Matters for Data Analysis Tools**

## Overview

We study **TOBENAMED dataset understanding**, the pre-analysis step in which a tool must infer the logical structure of an uploaded workbook before answering downstream questions.

This repository includes:

- the TOBENAMED prompting contract for workbook understanding
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