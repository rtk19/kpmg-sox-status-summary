# Codex Project Notes

Project name: `kpmg-sox-status-summary`

Last reviewed: 2026-05-27

## What This App Does

This is a Hebrew RTL web tool for KPMG SOX status summaries. Users upload one or more `.xlsx` SOX control matrix files, and the app generates a new `.xlsx` workbook that:

- preserves the original visible worksheets,
- skips existing generated summary/status sheets,
- detects each sheet's SOX table structure,
- appends a localized summary sheet named `סיכום` or `Summary`,
- downloads the generated workbook automatically.

The UI is rendered from `templates/index.html`, styled by `static/styles.css`, and driven by `static/app.js`. The workbook processing logic is in `sox_processor.py`.

## Structure

```text
app.py
sox_processor.py
requirements.txt
templates/index.html
static/app.js
static/styles.css
static/kpmg-logo.png
static/fonts/
```

## Run

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python app.py
```

Open `http://127.0.0.1:5000/`.

## Processing Flow

Public processor entry points:

- `process_sox_file(file, manual_overrides=None)`
- `process_sox_files(files, manual_overrides=None)`
- `classify_error(err)`

High-level algorithm:

1. Validate at least one file.
2. Load each workbook with `openpyxl`.
3. Analyze each source workbook to detect language.
4. Copy all visible non-generated worksheets into a combined workbook.
5. If multiple files have duplicate sheet names, prefix with the file stem.
6. Analyze the combined workbook.
7. Remove existing generated `status_N` sheets if any.
8. Extract all dynamic status values from detected status columns.
9. Create one summary sheet.
10. Return workbook bytes and the generated filename.

Generated filenames:

- Single file: `<original_stem>_with_status.xlsx`
- Multiple files: `sox_summary_of_controls_combined_with_status.xlsx`

## Sheet Detection Rules

Visible sheets only are processed. Hidden sheets are ignored.

Generated/input summary sheets are skipped while copying when their names match:

- `status_N`
- `סיכום`
- `summary`

Header row detection scans the first 30 rows and scores rows by:

- number of non-empty cells,
- header-like text cells,
- status header matches,
- key column matches,
- id/control number matches,
- non-empty count in the following row.

Table end detection starts below the header and stops at the first row where columns A, B, and C are all empty.

Status columns are required. Recognized status/conclusion headers include:

- Hebrew: `מסקנה`, `בקרה`, `סטטוס`, `מסקנה סבב א`, `מסקנה סבב א'`, and any header beginning with `מסקנה`.
- English: `status`, `conclusion`, `control`, `results`, and any header beginning with `conclusion`.

Key column is optional. Recognized key headers are prioritized:

1. `מפתח`, `key control`
2. `בקרת מפתח`, `key`
3. Other headers containing `מפתח` but not `לא מפתח`

If multiple candidates exist at the same priority, processing stops with a Hebrew clarification error.

ID column is optional. It prefers exact matches like `מס`, `מס'`, `מספר`, `מס בקרה`, `control number`, `control no`, `id`, `no`, `no.`.

Process column is optional. Recognized as `process`, `תהליך`, then secondary names like `process name`, `שם התהליך`. Headers containing `sub`, `sub-`, `sub_`, `משנה`, or `תת` are rejected so sub-process columns are not used.

Round numbering follows status column order only: first status column is round 1, second is round 2, etc.

## Summary Sheet Behavior

Language detection checks status headers. If uploaded source files have mixed languages, summary language defaults to English. Otherwise it uses the detected language, defaulting to Hebrew.

The summary sheet is named:

- Hebrew: `סיכום`
- English: `Summary`

Columns are dynamic:

- Always starts with sheet name.
- Adds process column if any processed sheet has a process column.
- Adds round column only if any sheet has more than one status column.
- Adds total controls, key controls, one column per dynamic status value, total statuses, IT controls, annual, execution percent.

Rows are created per `(sheet x process x status round)` when process values exist, or per `(sheet x status round)` otherwise. Rounds with no status values for that process/sheet are skipped.

The summary uses Excel formulas (`SUMPRODUCT`, `COUNTIF`, etc.) so Excel recalculates when opened.

Excluded status values are not added as dynamic status columns:

- empty string
- `לא מפתח`
- `not key`
- `IT`
- `it`
- `annual`
- `שנתי`

Execution percent depends on a dynamic status value exactly equal to `עבר` or `pass` after lowercase/trim. If no such status exists, execution percent is `0`.

## UI Notes

The UI should feel the same as the original:

- Hebrew-first and RTL.
- Heebo font served locally from `static/fonts`.
- KPMG logo served from `static/kpmg-logo.png`.
- Same upload, loading, success, help dialog, tabs, auto-download, redownload, and reset behavior.

## Current Gaps / Watch Points

- There are no formal fixture tests yet. Processor changes should be verified with representative `.xlsx` files.
- Summary formulas may not be recalculated by `openpyxl`; Excel/Sheets recalculates them when opened.
- Sheet copying preserves the important workbook features used here, but not every possible Excel feature.
- `manual_overrides` exists for key-column overrides but no UI currently exposes it.
- Very large workbooks may affect server memory/performance.
