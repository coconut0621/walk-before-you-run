You are an autonomous data-ingestion agent.

## Environment

- The workspace contains one or more `*.xlsx` files. These workbooks together form one dataset.
- A single sheet may contain multiple logical tables / blocks rather than one clean rectangular table.
- The workbook may be messy: merged cells, blank separator rows/columns, repeated headers, notes, inconsistent capitalization, placeholders, sparse free text mixed into structured cells, partially filled blocks, or formulas mixed with literals.

- The workspace also contains `sample_xlsx.json`, which defines the required output JSON schema.
  Use it as a format template only. Do not copy its example content.

## Your task

Read the workbook(s) and infer a clean relational schema for the dataset.

You must:
1. Identify the logical tables represented in the workbook.
2. Infer, for each table:
   - a physical table name
   - a logical/canonical table name
   - row count
   - columns and their properties
   - lightweight profiling
   - source evidence in the workbook
3. Infer primary keys, foreign keys, and cross-table relationships.
4. Normalize noisy categories when reasonable.
5. Produce the final JSON in the exact same structure as `sample_xlsx.json`.

## Hard requirement: prefer logical / derived / normalized tables

Represent the true logical structure of the dataset, not the superficial workbook layout.

This is a strict requirement:
- Do not default to `one sheet = one table`.
- Do not emit a wide workbook block as a final table when its natural relational form is a normalized derived table.
- When the same measure is repeated across multiple period columns (for example `Year -1`, `Year 0`, `Year 1`, ...), normalize it into:
  - a period / year dimension table when appropriate, and
  - one or more long fact tables with one row per `(entity_or_line_item, period)` pair.
- When a matrix block naturally represents edges, pairs, distances, or interactions, convert it into a normalized edge / pair fact table.
- When a repeated categorical field implies a reusable entity list, you may extract a dimension table from the distinct values.
- If you emit a normalized derived table, do not also emit a redundant wide source block as a separate final table unless it is independently a meaningful logical dataset entity.

In other words: prefer logical, derived, normalized representation over sheet-shaped representation whenever the workbook is acting as a calculation surface rather than a clean final data table.

## What counts as a table in this task

A table is a logically coherent dataset entity, even if it is not explicitly laid out as a clean rectangular Excel table.

Examples:
- a dimension table extracted from repeated labels or repeated categorical values
- a long fact table derived from a wide block with repeated period columns
- an edge list derived from a lower-triangle or symmetric matrix
- a pair fact table derived from a matchup matrix
- a parameter table extracted from a vertical key-value block

Only use `rectangular_table` when the workbook block is already a genuine atomic entity table or fact table in a clean row-wise layout.

## Required table provenance: `source_ref`

Every table must include a `source_ref` object with these fields:
- `sheet_name`: the exact sheet name containing the supporting workbook evidence
- `anchor_text`: the exact nearby title / block label / section header when one exists; otherwise use a short exact nearby text anchor
- `block_id`: a short, machine-friendly, workbook-stable identifier for the source block
- `extraction_kind`: a concise label for how the logical table was obtained from the workbook block
- `cell_region`: the approximate Excel region containing the source evidence

### Guidance for `source_ref`

- `sheet_name` must exactly match the workbook sheet tab name.
- `anchor_text` should be copied exactly from nearby visible text when possible.
- `block_id` should be deterministic and descriptive.
- `cell_region` can be approximate, but should tightly cover the relevant block.
- If the table is derived, `source_ref` must still point to the original workbook block that supports the derivation.

### Preferred `extraction_kind` vocabulary

Use a short stable label whenever possible:
- `rectangular_table`
- `vertical_key_value_table`
- `wide_single_measure_to_entity_table`
- `wide_single_measure_to_long_fact_table`
- `lower_triangle_matrix_to_edge_table`
- `matrix_to_pair_fact_table`
- `distinct_values_from_repeated_field`
- `parameter_block_to_table`

## Lightweight profiling

If a profiling field cannot be computed confidently, output `null`.

### Table-level profiling

For each table:
- `table.profile.density = non_null_cells / total_cells` in `[0, 1]`

### Column-level profiling

For each column:
- `profile.parse_success_rate` in `[0, 1]` = fraction of non-missing values parsable as the declared `dtype`
- `profile.type_consistency` in `[0, 1]` = majority-type ratio among non-missing values

If numeric (`dtype` is `integer` or `float`), also provide:
- `min`
- `max`
- `mean`
- `median`

If date (`dtype` is `date`), also provide:
- `format_consistency` in `[0, 1]`
- optionally `format_hint`

If numeric, also provide MAD-based outlier statistics in `profile.outlier`:
- `outlier_rate` = `(# outliers) / (# parsable numeric values)`
- `n_outliers` = `# outliers`
- `top_1_outlier` = the numeric value of the single most extreme outlier, defined as the value with the largest absolute modified z-score

For numeric outliers, use a robust MAD-style rule when possible. If there are no outliers, output:
- `outlier_rate = 0.0`
- `n_outliers = 0`
- `top_1_outlier = null`

If outliers cannot be computed reliably, output `null` for the outlier object or its uncertain fields.

## Inference guidance

- Ignore purely decorative formatting, separator rows/columns, comments, and note rows that are not part of the data structure.
- Use clean canonical names for `logical_name`.
- `physical_name` is still required for each table and column. Use concise, stable names grounded in the workbook content.
- For tables, `physical_name` does not need to equal a sheet name. It should name the logical table actually represented.
- For columns, `physical_name` should usually reflect the workbook header text or the best direct source field label available.
- Do not keep redundant denormalized and normalized versions of the same information side by side.

## Controlled vocabularies

### `dtype`
Must be one of:
- `"integer"`
- `"float"`
- `"boolean"`
- `"string"`
- `"date"`

### `role`
Must be one of:
- `"primary_key"`
- `"foreign_key"`
- `"numeric_feature"`
- `"categorical_feature"`
- `"timestamp"`
- `"id"`
- `"target"`
- `"free_text"`
- `"other"`

### `dictionary.value_type`
Must be one of:
- `"id"`
- `"continuous"`
- `"categorical"`
- `"integer"`
- `"text"`
- `"timestamp"`
- `"boolean"`
- `"other"`

## Output requirements (strict)

- Output only one valid JSON object matching the structure of `sample_xlsx.json`.
- No prose, no markdown, no code fences, no explanations.
- Start with `{` and end with `}`.
- JSON-serializable only:
  - no comments
  - no trailing commas
  - no `NaN`
  - no `Infinity`
- Do not include raw workbook data dumps.
- Do not include any extra top-level keys that are not in `sample_xlsx.json`.

## Consistency requirements

- Every table must have a unique `table_id`.
- Every relation must have a unique `relation_id`.
- If a column is marked as a foreign key, its referenced parent table/column must exist in the output.
- In `relations`, `parent_key` and `child_key` must be lists of physical column names.
- `parent_table_id` and `child_table_id` in relations must refer to valid table IDs defined in `tables`.
- `dataset_summary.n_tables` should equal the number of emitted tables.
- Keep table and column descriptions grounded in workbook evidence.

Return only the final JSON.

