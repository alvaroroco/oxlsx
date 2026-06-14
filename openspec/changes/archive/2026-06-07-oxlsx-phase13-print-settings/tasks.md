# Tasks: worksheet-print-settings

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 260-380 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single slice |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: pending
400-line budget risk: Low

## Phase 1: Reserved print-defined names bridge

- [x] 1.1 Add `src/defined_names.rs` parser/writer for `_xlnm.Print_Area` and `_xlnm.Print_Titles` only.
- [x] 1.2 Add print-title segment splitter and unknown `<definedName>` passthrough capture.
- [x] 1.3 Add unit tests for reserved parse, passthrough, and quoted-sheet commas.

## Phase 2: Worksheet-local API + workbook wiring

- [x] 2.1 Add worksheet-local `print_area`, `print_title_rows`, `print_title_cols` getters/setters in `src/reader.rs`.
- [x] 2.2 Parse reserved print defined names from `xl/workbook.xml` and project by `localSheetId`.
- [x] 2.3 Store unrelated defined names as passthrough fragments on `Workbook`.
- [x] 2.4 Emit `<definedNames>` in `write_workbook_xml` from worksheet state + passthrough extras.

## Phase 3: Python property wiring

- [x] 3.1 Add `PyWorksheet.print_area`, `print_title_rows`, and `print_title_cols` properties in `src/python.rs`.

## Phase 4: Spec test mapping (1:1 to 12 scenarios)

- [x] 4.1 Rust tests (6): `print_area_*` and `print_title_rows_and_cols_*` scenarios in `tests/integration_test.rs`.
- [x] 4.2 Python tests (6): `test_print_area_*` and `test_print_title_rows_and_cols_*` in `tests/test_python.py`.
