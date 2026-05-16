# Mutual Fund Parser

## Project Context

This project exists to parse mutual fund portfolio disclosure files into a clean, standardized dataset.

Asset Management Companies publish these disclosures in Excel formats that drift over time and vary across AMCs. In practice, headers change, sections move, columns split/merge, and noisy rows appear in otherwise structured tables.

The parser in this repo is metadata-driven. Your job is to improve robustness so parsing continues to work as these real-world format changes happen.

Your objective in this assignment: improve parser and metadata robustness under uncertainty, with limited hidden evaluator runs.

## Rules

1. Work on any non-`main` branch (name it however you want).
2. Metadata/parser fixes are evaluated in this assignment with a limited run budget.
3. Open a PR to `main`; evaluation runs only when you comment `/eval` on the PR, and is executed by a private controller workflow.

## Why This Exists

This setup intentionally tests whether you can use LLM-assisted coding with engineering discipline, not just speed.

- You should make deliberate, hypothesis-driven edits rather than broad random changes.
- You should constrain the agent's scope, validate locally, and only then push a patch and burn a run.
- You should avoid letting the coding agent run wild and burn evaluation runs.

Each hidden evaluation run is treated as a proxy for LLM token and compute spend. In real production systems, trial-and-error loops are expensive. This assignment is designed to reward careful iteration, regression awareness, and controlled decision-making.

Treat each `/eval` request as a paid move. Validate locally first.

## Evaluation Run
- Open a PR to `main` and push updates to your PR branch.
- Comment `/eval` on the PR to request an immediate evaluation run.
- Wait about 30-120 seconds for the evaluator run to be picked up and post `evaluate` status/comment output.
- Evaluation is not triggered automatically by push; only `/eval` requests trigger runs.
- Requests are deduplicated per PR head commit SHA. Repeating `/eval` on the same commit does not create a new run.
- To trigger a new run, push a new commit to the PR branch, then comment `/eval` again.
- Expand the PR comment for run details and evaluator log tail.
- Evaluator inputs and data are private and not present in this repo.
- See `RUBRIC.md` for scoring details and baseline score interpretation.

## Assignment Scope

- Visible inputs: `1_input/DSP/`
- Primary goal: fix extraction quality via robust metadata updates.
- Focus: stable anchors, consistent header mapping, no filename-specific hacks.

## Hierarchy Extraction Rules

The parser assigns `instrument_type`, `category`, and `subcategory` using a stateful
hierarchy tracker defined in `2_metadata/DSP.json` under each table's
`vertical_hierarchy`.

### How rows are interpreted

1. The parser first decides whether a row is a hierarchy marker or a data row.
2. If the row matches a hierarchy marker, the active hierarchy state is updated.
3. If the row is a data row, the current active hierarchy state is attached to that record.
4. `category_overrides` are then applied (for standard data rows), and can overwrite hierarchy fields or exclude the row.

### Marker matching behavior

- Marker text is normalized (spacing/prefix cleanup) before matching.
- If `hierarchy_marker_column` is configured, marker checks use that column.
- Otherwise, marker checks use the first non-empty normalized cell.
- If needed, the parser also checks the `instrument_name` column for marker text.
- Matching order is:
1. End markers (`end_keywords`) first
2. Child levels before parent levels
3. Parent (`instrument_type`) last

This prevents parent labels from incorrectly overriding an active child section.

### Table boundaries

- `table_end_keywords` close the current table.
- Hierarchy `end_keywords` close the matching active hierarchy level.
- Rows outside an active hierarchy block are skipped when the table defines a hierarchy.

### Net Receivables / Payables handling

- Rows containing configured `net_receivables_text` are treated as special records.
- These rows may use `net_receivables_instrument_type` / `net_receivables_category` overrides when configured.
- Otherwise, they inherit current hierarchy context.

### Example (conceptual)

Assume the parser sees these rows in order:

1. `DEBT INSTRUMENTS` (hierarchy marker)
2. `BOND & NCD's` (hierarchy marker)
3. `Listed / awaiting listing on the stock exchanges` (hierarchy marker)
4. `HDFC Bank Ltd ...` (data row)

For row 4, output hierarchy fields should be:

- `instrument_type = DEBT INSTRUMENTS`
- `category = BOND & NCD's`
- `subcategory = Listed / awaiting listing on the stock exchanges`

If a later row contains an end marker like `Total`, that closes the current active level.
Subsequent data rows inherit whatever hierarchy path remains active after that close.

## What You Can Change

- `3_parser/` parser logic (entrypoint: `3_parser/interview_parse.py`)
- `2_metadata/` metadata rules
- your own local scripts/checks/utilities

Possible local helpers:
- regression scripts
- output diff scripts
- record-count sanity checks
- table/header match diagnostics

## Strategy Tips

- Make small, hypothesis-driven edits.
- Verify locally before every push.
- Track exactly what changed per push.
- Protect prior behavior while fixing current failures.

## Local Quick Start

Install dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Run visible parses:

```bash
python3 3_parser/interview_parse.py --amc DSP
```

Default output directory:
- `4_output_interview/<AMC>/`

Optional CLI help:

```bash
python3 3_parser/interview_parse.py --help
```

## Repo Contents

- `1_input/` visible fixtures
- `2_metadata/` starter metadata
- `3_parser/` parser runtime
- `Header_Verification_Historical/config.json` shared parser config