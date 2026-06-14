# Tasks: Worksheet Freeze Panes

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 180-280 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | ask-always |
| Chain strategy | pending |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: pending
400-line budget risk: Low

## Phase 1: Foundation / API surface

- [x] 1.1 Add `freeze_panes: Option<String>` to `Worksheet` and expose `freeze_panes()` + `set_freeze_panes(Option<&str>)` with `A1` clear semantics.
- [x] 1.2 Add `PyWorksheet.freeze_panes` property wiring in `src/python.rs`.
- [x] 1.3 Keep coordinate handling local and minimal for split derivation.

## Phase 2: Core read/write implementation

- [x] 2.1 Parse sheet `<pane state="frozen" topLeftCell="...">` into worksheet state.
- [x] 2.2 Emit minimal Excel-compatible `<sheetViews>/<sheetView>/<pane>` + required `<selection>` before `<sheetData>`.
- [x] 2.3 Keep scope worksheet-local only (no defined names, print settings, tables).

## Phase 3: Integration / parity tests

- [x] 3.1 Add Rust integration test `freeze_panes_round_trip_existing_workbook`.
- [x] 3.2 Add Rust integration test `freeze_panes_write_new_workbook`.
- [x] 3.3 Add Rust integration test `freeze_panes_clear_removes_pane`.
- [x] 3.4 Add Python parity tests `test_freeze_panes_round_trip`, `test_freeze_panes_write_new_workbook`, and `test_freeze_panes_clear_removes_pane`.

## Phase 4: Known risk / deferred

- [x] 4.1 Document and keep risk explicit: unrelated sheet-view child data may be dropped in this slice.
- [ ] 4.2 Follow-up slice: preserve arbitrary unrelated `<sheetView>` state in RMW flows.
