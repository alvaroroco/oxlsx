# Tasks: Lossless RMW preservation of unmodeled table features

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 450-550 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 (Rust infrastructure + tests) → PR 2 (Python bindings + tests) |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main (recommended for independent review) |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: stacked-to-main
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Rust infrastructure: retention store, capture, dirty tracking, verbatim emission, id allocation, remove_table + Rust tests 1-5 | PR 1 | Base branch main; complete Rust behavior with tests/docs; ~350-400 LOC |
| 2 | Python bindings: expose remove_table, dirty tracking via property setters + Python tests 6-7 | PR 2 | Base branch PR #1 or main after merge; Python parity layer; ~100-150 LOC |

## Phase 1: Infrastructure & Data Model

- [x] 1.1 Add `RawTable` struct to `src/table.rs` with `pub(crate) struct RawTable { pub bytes: Vec<u8>, pub dirty: bool }`
- [x] 1.2 Add `table_retention: Vec<Option<RawTable>>` field to `Worksheet` in `src/reader.rs` (initialize empty in constructors)
- [x] 1.3 Update `Worksheet::new()` and `Worksheet::default()` to initialize `table_retention` as empty `Vec`

## Phase 2: Reader — Capture Raw Bytes

- [x] 2.1 In `src/reader.rs` at the table-read loop (~lines 648-653), read the full part bytes into `Vec<u8>` (use `std::io::Read::read_to_end`), then parse via `Cursor` wrapper
- [x] 2.2 After parsing each table, push `Some(RawTable { bytes, dirty: false })` to `ws.table_retention` aligned with the `ws.tables` push
- [x] 2.3 **TEST**: Write `table_capture_stores_raw_bytes` — open a table-bearing fixture, assert `ws.table_retention[0].is_some()` and `bytes.len() > 0`

## Phase 3: Dirty Tracking & Removal

- [x] 3.1 Update `Worksheet::table_mut` to mark dirty: locate table index, if `table_retention[idx].is_some()`, set `.dirty = true`, then return `&mut tables[idx]`
- [x] 3.2 Add `pub fn remove_table(&mut self, display_name: &str) -> Option<Table>` to `Worksheet` — find index, remove from both `tables` and `table_retention`, return the `Table`
- [x] 3.3 Update `Worksheet::add_table` to push `None` to `table_retention` (new tables have no raw bytes)
- [x] 3.4 **TEST**: Write `dirty_tracking_getters_do_not_mark_dirty` — open fixture, call `ws.tables()`, assert `table_retention[0].dirty == false`; call `ws.table_mut(...)`, assert `dirty == true`

## Phase 4: Writer — ID Allocation Refactor

- [x] 4.1 In `src/writer.rs`, rename `assign_table_numbers` to `assign_table_ids`
- [x] 4.2 Refactor `assign_table_ids` to return `Vec<Vec<(u32 /* assigned id */, bool /* is_verbatim */)>>` — scan all sheets, collect verbatim internal ids from `table_retention` (parse `<table id=...>`), build a `HashSet<u32>`, then allocate `next_free` ids for dirty/new tables avoiding the HashSet
- [x] 4.3 **TEST**: Write unit test `assign_table_ids_avoids_verbatim_collision` — mock a workbook with one verbatim table `id=5` and one new table, assert the new table gets `id != 5` and both are unique

## Phase 5: Writer — Verbatim vs Model Emission

- [x] 5.1 In the writer's table-part emission loop (currently `writer.rs:358-372`), introduce a global `part_number` counter starting at 1
- [x] 5.2 For each table on each sheet, check `ws.table_retention[table_idx]`: if `Some(raw)` and `!raw.dirty`, write `raw.bytes` to `xl/tables/table{part_number}.xml` WITHOUT modifying the bytes; else serialize from model with the pre-allocated id via `serialize_table_xml`
- [x] 5.3 Increment `part_number` after each table; use `part_number` for content-type override, `<tableParts>` rId, and sheet-rel target (all existing wiring)
- [x] 5.4 Update `write_content_types_from_source` to strip ALL old `/xl/tables/tableN.xml` overrides and re-add `1..=part_number` (already done for Phase 13, confirm still correct)

## Phase 6: Rust Integration Tests (TDD — write tests FIRST per scenario)

- [x] 6.1 **FIXTURE**: Author `samples/tables_totals.xlsx` via `.venv` openpyxl: one table with `totalsRowCount=1`, `totalsRowFunction="sum"`, a `calculatedColumnFormula`, and a `sortState` (store in repo)
- [x] 6.2 **TEST 1**: `table_unmodified_rmw_preserves_raw_table_xml` — open `tables_totals.xlsx`, edit cell B5, save, reopen, assert the saved `xl/tables/table1.xml` contains `totalsRowCount`, `totalsRowFunction`, `calculatedColumnFormula`, and `sortState` attributes/children (byte or substring match)
- [x] 6.3 **TEST 2**: `table_mutated_rmw_serializes_from_model` — open same fixture, call `table_mut`, set a new `style`, save, assert `totalsRowCount` is GONE from the output XML but modeled fields (`displayName`, `ref`, `columns`) are correct
- [x] 6.4 **TEST 3**: `table_added_then_saved_uses_model_path` — create new workbook, add a table via `add_table`, save, assert part exists and is Excel-valid (no repair on reopen)
- [x] 6.5 **TEST 4**: `table_removed_drops_part_rel_and_tablepart` — open fixture with 2 tables, remove one via `remove_table`, save, assert the removed table's part is absent from the ZIP, its `<tablePart>` is gone from sheet XML, and the remaining table is intact
- [x] 6.6 **TEST 5**: `mixed_verbatim_and_model_tables_unique_ids` — open fixture with one verbatim table, add a new table, save, assert both have distinct workbook-wide `<table id>`s and the case-insensitive `displayName` uniqueness guard still holds

## Phase 7: Python Bindings & Tests

- [~] 7.1 ~~Add `remove_table(name: str)` method to `PyWorksheet`~~ — DEVIATION: openpyxl has NO `ws.remove_table` (verified in venv: `hasattr(ws, "remove_table") == False`). Per the MANDATORY openpyxl-parity rule, the removal path is exposed only through `del ws.tables["X"]` (task 7.2). The Rust `Worksheet::remove_table` (Unit 1) is the backing call.
- [x] 7.2 Implement `__delitem__` on `PyTableList` to support `del ws.tables["Table1"]`, routing to `Worksheet::remove_table` (raises `KeyError` when absent, mirroring openpyxl's `dict` subclass)
- [x] 7.3 Ensure Python property setters on `PyTable` route through the inner `Attached` mutation path — CONFIRMED: `set_tableStyleInfo` already calls `with_table_mut` → `Worksheet::table_mut` (marks dirty). Read-only getters use `with_table` → `tables()` (no dirty). No change needed.
- [x] 7.4 **TEST 6**: `test_table_totals_row_survives_cell_edit_rmw` (THE mandatory regression) — open `samples/tables_totals.xlsx` via Python oxlsx, edit unrelated cell `H10`, save, reopen with **openpyxl** (venv 3.1.2), assert no repair AND `totalsRowCount == 1`, `Revenue.totalsRowFunction == "sum"`, `Units.calculatedColumnFormula` present, `sortState` present
- [x] 7.5 **TEST 7**: `test_table_mutation_documents_feature_loss` — open same file, mutate via `ws.tables["SalesTable"].tableStyleInfo`, save, reopen with openpyxl: valid, modeled fields (`displayName`/`ref`/columns/style) correct, unmodeled features dropped on the model path (Inventory stays verbatim). Plus `test_table_delete_via_python_removes_table` + `test_table_delete_missing_raises_keyerror` cover the `__delitem__` path.

## Phase 8: Cleanup & Documentation

- [x] 8.1 Update the module-level doc in `src/table.rs` to reflect that unmodeled features NOW survive RMW when the table is not mutated (done in Unit 1; verified: module doc + `Table`/`RawTable` docblocks describe verbatim-when-unmodified)
- [x] 8.2 Run `cargo clippy -- -D warnings` and `cargo fmt --check` — fix any new lints (clean on both default and `--features python`)
- [x] 8.3 Run full test suite (`cargo test` + `pytest`) — ensure no regressions (cargo: 92 lib + 87 integ PASS; pytest: 121 passed, 0 failed)
