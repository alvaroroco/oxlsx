# Tasks: worksheet-bulk-write (Real Phase 10)

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 450–550 (Rust core ~200, PyO3 ~120, tests ~170, docs ~10) |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | Slice 1 (core) → Slice 2 (append+move) → Slice 3 (Python) |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: pending
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| Slice 1 | Core shift engine + insert/delete (4 ops) + merges/dims + table guard | PR 1 | Base = main or feature/bulk-write; tests+docs; scenarios 2,3,4,5,7,8,10 |
| Slice 2 | append + move_range + doc fix | PR 2 | Base = Slice 1 branch or main; tests+docs; scenarios 1,6,9 |
| Slice 3 | PyO3 bindings (6 ops) + dict-append + cross-engine | PR 3 | Base = Slice 2 branch or main; pytest; scenarios 11,12,13 |

## Slice 1: Core Shift Engine + Insert/Delete (Scenarios 2,3,4,5,7,8,10)

### Phase 1.1: Foundation (coords + error)

- [x] 1.1.1 Add `parse_a1(ref: &str) -> Option<(u32, u32)>` helper to `src/coords.rs` (parse "B3" → (2, 3))
- [x] 1.1.2 Add `parse_range(ref: &str) -> Option<(u32, u32, u32, u32)>` helper to `src/coords.rs` (parse "A1:C5" → (1,1,3,5))
- [x] 1.1.3 Add test `coords_parse_a1_and_range_correctly` in `src/coords.rs` (verify "AA10" → (27,10), "A1:B3" → (1,1,2,3))
- [x] 1.1.4 Add `OxlsxError::TableIntersection(String)` and `OxlsxError::NotImplemented(String)` to `src/error.rs`
- [x] 1.1.5 Add test `error_renders_table_intersection_message` in `src/error.rs` (verify display string)

### Phase 1.2: Shift Engine Core (private)

- [x] 1.2.1 RED: Write `#[test] insert_rows_shifts_cells_down` in `src/reader.rs` (scenario 2: insert at idx=2 shifts rows 2..4 → 4..6)
- [x] 1.2.2 Add private `shift_rows(&mut self, start_idx: u32, delta: i32) -> Result<(), OxlsxError>` to `impl Worksheet` in `src/reader.rs` (drain+rebuild `cells` HashMap via `parse_a1`/`rc_to_a1`; drop if deleted; shift `row_heights`)
- [x] 1.2.3 GREEN: Implement `pub fn insert_rows(&mut self, idx: u32, amount: u32) -> Result<(), OxlsxError>` (call `shift_rows(idx, amount as i32)` after table guard) to pass test
- [x] 1.2.4 RED: Write `#[test] delete_rows_shifts_cells_up` in `src/reader.rs` (scenario 3: delete idx=2 amount=1 shifts rows 3..5 → 2..4)
- [x] 1.2.5 GREEN: Implement `pub fn delete_rows(&mut self, idx: u32, amount: u32) -> Result<(), OxlsxError>` (call `shift_rows(idx, -(amount as i32))`) to pass test
- [x] 1.2.6 RED: Write `#[test] delete_rows_does_not_translate_formulas` in `src/reader.rs` (scenario 10: formula `=A3+1` at B3 shifts to B2 with text unchanged)
- [x] 1.2.7 Verify formula verbatim preservation in `shift_rows` (formulas ride on `Value::Formula(String)`, no mutation needed)

### Phase 1.3: Column shift (mirror row logic)

- [x] 1.3.1 RED: Write `#[test] insert_cols_shifts_cells_right` in `src/reader.rs` (scenario 4: insert at idx=2 shifts cols B..D → C..E)
- [x] 1.3.2 Add private `shift_cols(&mut self, start_idx: u32, delta: i32) -> Result<(), OxlsxError>` to `impl Worksheet` (same logic, shift col index; shift `column_widths`)
- [x] 1.3.3 GREEN: Implement `pub fn insert_cols(&mut self, idx: u32, amount: u32) -> Result<(), OxlsxError>` (call `shift_cols(idx, amount as i32)`) to pass test
- [x] 1.3.4 RED: Write `#[test] delete_cols_shifts_cells_left` in `src/reader.rs` (scenario 5: delete idx=2 amount=1 shifts cols C..D → B..C)
- [x] 1.3.5 GREEN: Implement `pub fn delete_cols(&mut self, idx: u32, amount: u32) -> Result<(), OxlsxError>` (call `shift_cols(idx, -(amount as i32))`) to pass test

### Phase 1.4: Merged cells & dimensions shift

- [x] 1.4.1 RED: Write `#[test] delete_rows_shifts_merged_cells_and_dimensions` in `src/reader.rs` (scenario 7: delete idx=2, a merge "A4:B5" → "A3:B4", partial-straddle dropped, row_heights shifted)
- [x] 1.4.2 Add merge-shift logic to `shift_rows`: parse each `merged_cells` range via `parse_range`, shift endpoints if fully past band, drop if partial-straddle
- [x] 1.4.3 GREEN: Verify test passes with merge+dim shift
- [x] 1.4.4 Add merge-shift logic to `shift_cols` (mirror row logic for cols)

### Phase 1.5: Table intersection guard

- [x] 1.5.1 RED: Write `#[test] bulk_op_intersecting_table_returns_err` in `src/reader.rs` (scenario 8: worksheet with table A1:C5, insert_rows(2,1) → `Err(TableIntersection)`, workbook unchanged)
- [x] 1.5.2 Add private `check_table_intersection(&self, axis: Axis, start_idx: u32, amount: u32) -> Result<(), OxlsxError>` to `impl Worksheet` (parse each `table.ref_range` via `parse_range`, test for band overlap)
- [x] 1.5.3 GREEN: Call `check_table_intersection` at the top of `insert_rows`/`delete_rows`/`insert_cols`/`delete_cols`; test passes

### Phase 1.6: Clippy + fmt

- [x] 1.6.1 Run `cargo fmt` and `cargo clippy -- -D warnings`; fix any issues
- [x] 1.6.2 Run `cargo test` and verify all Slice 1 scenarios (2,3,4,5,7,8,10) pass

## Slice 2: Append + move_range + Doc Fix (Scenarios 1,6,9 + doc fix)

### Phase 2.1: Append (Rust core)

- [x] 2.1.1 RED: Write `#[test] append_adds_row_at_bottom` in `src/reader.rs` (scenario 1: append [v1,v2,v3] at max_row+1)
- [x] 2.1.2 GREEN: Implement `pub fn append(&mut self, values: impl IntoIterator<Item = Value>)` in `src/reader.rs` (compute `max_row() + 1`, iterate values from col A, call `set_cell`)
- [x] 2.1.3 Verify test passes

### Phase 2.2: move_range

- [x] 2.2.1 RED: Write `#[test] move_range_relocates_block` in `src/reader.rs` (scenario 6: move "A1:B2" by rows=2,cols=1 with translate=False, formula text verbatim)
- [x] 2.2.2 GREEN: Implement `pub fn move_range(&mut self, range: &str, rows: i32, cols: i32, translate: bool) -> Result<(), OxlsxError>` (parse range, check table intersection for source+dest, extract cells, recompute refs, clear source, insert at dest; error if translate=True)
- [x] 2.2.3 RED: Write `#[test] move_range_translate_true_is_unsupported` in `src/reader.rs` (scenario 9: move_range with translate=True → `Err(NotImplemented)`)
- [x] 2.2.4 GREEN: Add `if translate { return Err(NotImplemented(...)) }` at the top of `move_range`; test passes

### Phase 2.3: Doc fix

- [x] 2.3.1 Edit `SPEC.md` line 103: change "✅ Phase 10 — Bulk Write API" to "🔲 Phase 10 — Bulk Write API (in progress)"
- [x] 2.3.2 Edit `README.md` line 224: uncheck the bulk-write item (change `[x]` to `[ ]`)
- [x] 2.3.3 Run `cargo test` and verify all Slice 2 scenarios (1,6,9) pass

## Slice 3: PyO3 Bindings + Cross-engine Tests (Scenarios 11,12,13)

### Phase 3.1: PyO3 method bindings

- [x] 3.1.1 Add `insert_rows(idx: u32, amount: Option<u32>)` to `#[pymethods] impl PyWorksheet` in `src/python.rs` (lock workbook, call `ws.insert_rows(idx, amount.unwrap_or(1))`, map error)
- [x] 3.1.2 Add `delete_rows(idx: u32, amount: Option<u32>)` to `#[pymethods] impl PyWorksheet` (mirror insert_rows)
- [x] 3.1.3 Add `insert_cols(idx: u32, amount: Option<u32>)` to `#[pymethods] impl PyWorksheet`
- [x] 3.1.4 Add `delete_cols(idx: u32, amount: Option<u32>)` to `#[pymethods] impl PyWorksheet`
- [x] 3.1.5 Add `move_range(range: &str, rows: Option<i32>, cols: Option<i32>, translate: Option<bool>)` to `#[pymethods] impl PyWorksheet` (defaults: rows=0,cols=0,translate=false)

### Phase 3.2: append with dict/list padding

- [x] 3.2.1 Add `append(iterable: &Bound<'_, PyAny>)` to `#[pymethods] impl PyWorksheet` in `src/python.rs` (check if dict: extract keys (str or int), resolve via `column_index_from_letters`/direct int, build padded `Vec<Value>` with `Value::Empty`, else iterate and map via `py_to_value`, call `ws.append`)
- [x] 3.2.2 RED: Write pytest `test_append_list_and_dict` (scenario 11: append list, then dict with letter/int/mixed keys, verify columns match openpyxl)
- [x] 3.2.3 GREEN: Debug padding logic until test passes

### Phase 3.3: Cross-engine round-trip tests

- [x] 3.3.1 RED: Write pytest `test_delete_and_insert_rows` (scenario 12: identical fixture in oxlsx+openpyxl, apply same insert/delete sequence, read back, assert grids equal)
- [x] 3.3.2 GREEN: Debug until parity holds
- [x] 3.3.3 RED: Write pytest `test_move_range` (scenario 13: move_range with translate=False parity against openpyxl)
- [x] 3.3.4 GREEN: Debug until parity holds

### Phase 3.4: Final verification

- [x] 3.4.1 Run `cargo clippy -- -D warnings` and `cargo fmt --check`
- [x] 3.4.2 Run `cargo test` (all Rust scenarios 1-10)
- [x] 3.4.3 Run `pytest` (all Python scenarios 11-13)
- [x] 3.4.4 Verify all 14 spec scenarios pass
