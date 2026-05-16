# Rubric

This assignment reports a **progress score** in addition to the strict pass/fail evaluator result.

- Pass/fail still requires exact output match.
- Rubric score is for progress tracking and partial credit visibility.

## Scoring Model (Per Parsed File)

Total: **100 points**

- **Record Count (40 points)**
  - Measures how close your output row count is to expected.
  - Row metrics use a canonical row view: `All Data` sheet when present; otherwise all expected sheets.
  - `count_ratio = max(0, 1 - abs(candidate_rows - expected_rows) / expected_rows)`
  - `count_points = 40 * count_ratio`

- **Record Accuracy (40 points)**
  - Measures how many expected rows are actually correct.
  - Uses **exact full-row matching** against expected output.
  - A row is counted as correct only if **every expected column value** matches.
  - If even one column is wrong (or missing), that row is counted as incorrect.
  - Computed on the same canonical row view used for record count.
  - `accuracy_ratio = correct_rows / expected_rows`
  - `accuracy_points = 40 * accuracy_ratio`

- **Column Presence (10 points)**
  - Measures whether expected columns exist in your parsed output.
  - Computed across all expected sheets.
  - `column_presence_ratio = expected_columns_present / total_expected_columns`
  - `column_presence_points = 10 * column_presence_ratio`

- **Header Literal + Order (10 points)**
  - Measures exact header-string match in the expected column positions.
  - Computed across all expected sheets.
  - `header_ratio = headers_matching_exact_string_and_position / total_expected_columns`
  - `header_points = 10 * header_ratio`

File score:
- `file_score = count_points + accuracy_points + column_presence_points + header_points`

AMC score:
- Recomputed from aggregate totals across all files in scope.

## Baseline

Using the starter assignment inputs and starter parser/metadata (no candidate edits), the DSP baseline is:

- **DSP AMC total: 66.84 / 100**
  - Count: 38.00
  - Accuracy: 12.37
  - Column Presence: 8.24
  - Header Literal + Order: 8.24
