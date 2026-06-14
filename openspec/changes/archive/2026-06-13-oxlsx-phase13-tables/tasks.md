# Tasks: Worksheet Tables (`ws.add_table` / `ws.tables`)

## Status: ALL UNITS COMPLETE (1 READ, 2 WRITE+RMW, 3 Python) — 11/11 spec tests green. Ready for sdd-verify.
Final gates: cargo lib 83 + integration 75; clippy (python) 0; fmt 0; pytest 114.

## Chain: feature-branch-chain (`feature/phase13-tables`) — PR1 (Unit 1) → PR2 (Unit 2) → PR3 (Unit 3). Delivery: auto-chain.

## DESIGN AMENDMENT (locked, implemented)
- Unit 2 (Rust): `wb.save()` workbook-wide duplicate guard → `Err(OxlsxError::DuplicateTableName)`. DONE.
- Unit 3 (Python): early `ValueError` in `ws.add_table` before mutation + `map_oxlsx_error` maps DuplicateTableName→ValueError. DONE.

---

## Unit 1: Rust READ — COMPLETE (Scenario 1)
Tasks 1.1–1.13 done. `relation_targets_for_type` helper added (multi-table). Fixture `samples/tables.xlsx`.

## Unit 2: Rust WRITE + RMW Orphan Fix + Rust Duplicate Guard — COMPLETE (Scenarios 2-7 + 2.18)
Tasks 2.1–2.23 done. `serialize_table_xml`, `derive_columns_from_header`, `Table::new`, `Worksheet::add_table`, `upsert_table_rels`/`relation_id_for_target`/`next_rel_id`, workbook-wide id allocation, CT overrides (both paths), `<tableParts>` last child, `xmlns:r`, `is_writer_owned_entry` claims `xl/tables/*`, `SheetRelRefs` struct. openpyxl-validated (no repair). Discovered pre-existing comment/vml rId bug (obs #274, NOT fixed, out of scope).

## Unit 3: Python Parity — COMPLETE (Scenarios 8, 9a, 9b, 10)

### Phase 1: PyO3 Model
- [x] 3.1 `#[pyclass(name="Table")] PyTable`; getters `displayName`, `ref` (`#[pyo3(name="ref")]`), `name`, `tableColumns`, `tableStyleInfo` (+ setter); ctor `Table(displayName, ref)` via `r#ref`
- [x] 3.2 `#[pyclass(name="TableStyleInfo")] PyTableStyleInfo` — `TableStyleInfo(name, *, showFirstColumn=False, showLastColumn=False, showRowStripes=False, showColumnStripes=False)` + camelCase getters
- [x] 3.3 `#[pyclass(name="TableColumn")] PyTableColumn` (id, name)
- [x] 3.4 Registered PyTable/PyTableStyleInfo/PyTableColumn in module init

### Phase 2: Dict-like `ws.tables`
- [x] 3.5 `PyTableList` — `__getitem__(name)`, `__contains__`, `__len__`, `__iter__` (keys), `keys()` + `PyTableNamesIter`
- [x] 3.6 `PyWorksheet.tables` getter
- [x] 3.7 Pytest Scenario 8 `test_tables_round_trip` (2×2 engine matrix)

### Phase 3: `add_table` with Early Duplicate Check
- [x] 3.8 `PyWorksheet.add_table` — early workbook-wide scan → `ValueError` before Rust mutation; `map_oxlsx_error` maps `DuplicateTableName`→`ValueError`
- [x] 3.9 Pytest Scenario 9a `test_tables_add_new`
- [x] 3.10 Pytest Scenario 9b `test_tables_duplicate_display_name_raises` (parametric: both openpyxl & oxlsx raise ValueError)
- [x] 3.11 Pytest Scenario 10 `test_tables_style_info`

**Unit 3 NOTE**: PyTable is a snapshot clone. Column-derivation timing differs (oxlsx eager / openpyxl lazy) but converges after save+reload — tests assert on reloaded tables. Tests cross-validate oxlsx↔openpyxl in both directions.

---

## Scenario Mapping (1:1 + amended) — ALL DONE
| # | Scenario | Unit | Status |
|---|----------|------|--------|
|1|tables_read_from_existing_workbook|1|✅|
|2|tables_add_and_save_round_trip|2|✅|
|3|tables_add_emits_table_part_and_sheet_rel|2|✅|
|4|tables_columns_derived_from_header_row|2|✅|
|5|tables_style_info_round_trip|2|✅|
|6|tables_rmw_preserves_existing_tables|2|✅|
|7|tables_multiple_tables_unique_ids_and_rids|2|✅|
|2.18|tables_save_rejects_duplicate_display_name (Rust)|2|✅|
|8|test_tables_round_trip|3|✅|
|9a|test_tables_add_new|3|✅|
|9b|test_tables_duplicate_display_name_raises|3|✅|
|10|test_tables_style_info|3|✅|

**Total**: 11/11 tests done. Change implementation COMPLETE — proceed to sdd-verify → sdd-archive.
