# worksheet-tables Specification (Phase 13 — minimal Excel-valid cut)

## Purpose
Expose Excel Tables with openpyxl parity: read existing tables (`ws.tables`) and create minimal Excel-VALID tables (`ws.add_table`). Reuse `src/sheet_rels.rs` and the Phase 12b comments new-part pattern. Fix the confirmed latent RMW bug where `write_sheet_xml` drops `<tableParts>`, orphaning existing tables on save.

## In Scope
Table model (`id`, `name`, `display_name`, `ref`, `header_row_count`, `columns: Vec<TableColumn{id,name}>`, `style: Option<TableStyleInfo>`); `TableStyleInfo{name, show_first_column, show_last_column, show_row_stripes, show_column_stripes}`; read + create + round-trip; workbook-wide id/rId management; RMW preservation; Python parity (`Table`, `TableStyleInfo`, `ws.add_table`, `ws.tables`).

## Out of Scope (deferred)
dxf style overrides, totals row, calculated columns, autoFilter sort/filter criteria, sortState. **OUT: charts, pivots.**

## XML Contract (locked, verified vs openpyxl 3.1.5)
- Part `xl/tables/tableN.xml`: `<table id="N" name displayName ref headerRowCount="1" xmlns="...spreadsheetml/2006/main"><autoFilter ref="{ref}"/><tableColumns count="K"><tableColumn id="1" name="..."/>...</tableColumns>[<tableStyleInfo name showFirstColumn showLastColumn showRowStripes showColumnStripes/>]</table>`. Booleans emitted as `1`/omitted. `tableStyleInfo` emitted only when style present.
- `[Content_Types].xml`: `<Override PartName="/xl/tables/tableN.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`. No new Default extension.
- Sheet XML: `<tableParts count="M"><tablePart r:id="rIdX"/>...</tableParts>` as the LAST child of `<worksheet>` (after pageMargins/all existing children).
- Sheet rel (via `sheet_rels.rs`): `Type=".../relationships/table"`, `Target="/xl/tables/tableN.xml"` (leading `/`; `resolve_sheet_target` already handles it). Add `TABLE_REL_TYPE` + `upsert_table_rels` mirroring `upsert_comment_rels`.
- id/rId: table `id` unique workbook-wide (scan all sheets + parsed ids); rIds allocated next-free per sheet, must coexist with comments/hyperlink rels.

## Requirements

### Requirement: Read existing tables
The system MUST parse `xl/tables/tableN.xml` via the `.../table` sheet rel and expose them through `ws.tables`, a dict-like collection keyed by `display_name`.

#### Scenario: tables_read_from_existing_workbook
- GIVEN an xlsx whose sheet references a table part
- WHEN the workbook is opened
- THEN `ws.tables["Table1"]` exposes the correct `display_name`, `ref`, and `columns` (id + name each)

### Requirement: Create tables with derived columns and duplicate rejection
`ws.add_table(table)` MUST add a table, auto-deriving `columns` from the header row of `ref` (ids `1..N`, names from header cells) when columns are not supplied. It MUST reject a `display_name` already present anywhere in the workbook, returning `Err(OxlsxError)` (Python `ValueError`) — matching openpyxl. It MUST NOT mutate state on rejection.

#### Scenario: tables_add_and_save_round_trip
- GIVEN a worksheet with a header row at the top of `ref`
- WHEN `add_table(Table{display_name:"Table1", ref:"A1:C3"})` is called, saved, and reopened
- THEN the table is present with the same `display_name`, `ref`, and `columns`

#### Scenario: tables_columns_derived_from_header_row
- GIVEN `ref="A1:C3"` with header cells `h1,h2,h3` and no explicit columns
- WHEN the table is added
- THEN columns are `{id:1,name:"h1"},{2,"h2"},{3,"h3"}`

#### Scenario: add_table rejects duplicate display_name
- GIVEN a workbook already containing table `"Table1"`
- WHEN `add_table` is called with `display_name:"Table1"` (any sheet)
- THEN it returns `Err`/raises `ValueError` and no new part is emitted

### Requirement: Excel-valid serialization
On save the system MUST emit the table part, the content-type override, `<tableParts>` in the sheet XML, and the sheet rel per the XML Contract, producing a file openpyxl/Excel opens without repair.

#### Scenario: tables_add_emits_table_part_and_sheet_rel
- GIVEN a worksheet with one added table
- WHEN saved
- THEN the ZIP contains `xl/tables/table1.xml`, the CT override, `<tableParts>` in the sheet XML, and a `.../table` rel in `xl/worksheets/_rels/sheet1.xml.rels`

### Requirement: TableStyleInfo round-trip
`TableStyleInfo` (`name` + 4 booleans) MUST round-trip; emitted only when set.

#### Scenario: tables_style_info_round_trip
- GIVEN a table with `TableStyleInfo{name:"TableStyleMedium9", show_row_stripes:true}`
- WHEN saved and reopened
- THEN the style name and boolean flags are preserved

### Requirement: Workbook-wide unique ids and non-colliding rIds
Multiple tables MUST receive unique workbook-wide `id`s; per-sheet table rIds MUST NOT collide with existing comment/hyperlink rels.

#### Scenario: tables_multiple_tables_unique_ids_and_rids
- GIVEN two tables added (same or different sheets), at least one sheet also having comment rels
- WHEN saved
- THEN each table has a distinct `id`, each `tablePart` rId is unique within its sheet, and pre-existing rels are intact

### Requirement: RMW preserves existing tables and unknown parts
Opening a table-bearing file and saving MUST re-emit `<tableParts>` and retain the table parts + rels (regression for the latent drop bug). Unknown parts and unrelated sheet rels (comments/hyperlinks) MUST be preserved. Additionally, for any table that was read and NOT mutated through the API, the saved `xl/tables/tableN.xml` MUST be byte-for-byte the original retained bytes, so unmodeled features (`totalsRow*`, `calculatedColumnFormula`, `dxf`, `autoFilter` criteria, `sortState`, etc.) survive with ZERO loss. The common case — open a file, edit an unrelated cell, save with all tables untouched — MUST be fully lossless.
(Previously: only re-emitted `<tableParts>` and retained the part/rel; the part itself was always re-serialized from the thin model, silently dropping unmodeled features.)

#### Scenario: tables_rmw_preserves_existing_tables
- GIVEN a file with an existing table (and comment/hyperlink rels)
- WHEN an unrelated cell is modified and the file is saved
- THEN the existing table is still referenced via `<tableParts>` (not orphaned), its part/rel remain, and unrelated rels + unknown parts are preserved

#### Scenario: table_unmodified_rmw_preserves_raw_table_xml
- GIVEN a fixture whose `tableN.xml` carries unmodeled features (`totalsRowCount`, a `totalsRowFunction`, a `<calculatedColumnFormula>`, and a `<sortState>`)
- WHEN an UNRELATED cell is modified and the workbook is saved (the table is not touched)
- THEN the saved `xl/tables/tableN.xml` is byte-equivalent to the original (the totals row, calculated-column formula, and sort state all survive) and the `<tableParts>`/rel/content-type wiring is intact

### Requirement: Verbatim retention controls the table serialization path
The system MUST retain the original `tableN.xml` bytes for every table read from a source workbook, stored off the public `Table` struct. On save, for each table the system MUST choose exactly one path: if retained bytes exist AND the table is not dirty, re-emit the retained bytes verbatim into `xl/tables/tableN.xml`; otherwise serialize from the model via the existing XML Contract. A table created via `add_table` has no retained bytes and MUST use the model path. The chosen path MUST NOT alter `<tableParts>` ordering, the content-type override, or sheet-rel allocation — verbatim bytes are placed at the writer-assigned part filename and wired by rId exactly as model-serialized parts are.

#### Scenario: table_mutated_rmw_serializes_from_model
- GIVEN the same totals-row fixture opened for RMW
- WHEN the table is mutated through the API (e.g. a new `TableStyleInfo` is set) and the workbook is saved
- THEN that table's `xl/tables/tableN.xml` is serialized from the model (documented loss: totalsRow/calculatedColumnFormula/sortState are dropped), the modeled fields (`display_name`, `ref`, `columns`, `style`) round-trip, and the file opens in openpyxl without repair

#### Scenario: table_added_then_saved_uses_model_path
- GIVEN a workbook with a brand-new table added via `add_table` (no retained bytes)
- WHEN the workbook is saved
- THEN the table is serialized from the model per the XML Contract and opens without repair

#### Scenario: table_removed_drops_part_rel_and_tablepart
- GIVEN a workbook with two retained tables, one of which is removed before save
- WHEN the workbook is saved
- THEN the removed table's retained bytes, `xl/tables/tableN.xml` part, its sheet rel, its `<tablePart>` entry, and its content-type override are all gone, while the remaining table's part/rel/`<tablePart>`/override and all unrelated parts are preserved

### Requirement: Table dirty-tracking determines model vs verbatim
The system MUST track, per retained table, whether it has been mutated through the API ("dirty"). Any operation that exposes mutable access to a table's modeled fields — `table_mut(display_name)` returning `&mut Table`, and through it any change to `name`, `display_name`, `ref_range`, `header_row_count`, `columns`, or `style` — MUST mark that table dirty (conservatively: acquiring `&mut` access marks it dirty). Read-only access — `tables()` and any by-name getter — MUST NOT mark a table dirty. A dirty table MUST take the model serialization path; a clean retained table MUST take the verbatim path.

#### Scenario: dirty_tracking_getters_do_not_mark_dirty
- GIVEN a retained table read from a source workbook
- WHEN only read-only accessors (`tables()`, by-name lookup) are used and an unrelated cell is edited
- THEN the table remains clean and is re-emitted verbatim on save

#### Scenario: mixed_verbatim_and_model_tables_unique_ids
- GIVEN a workbook with one retained (clean, verbatim) table and one newly added (model) table
- WHEN the workbook is saved
- THEN both tables receive unique workbook-wide `id`s (model-path ids allocated to avoid the retained verbatim id) and unique per-sheet rIds, the case-insensitive `display_name` uniqueness guard still holds across both, and both reopen intact

### Requirement: Python parity for lossless RMW
The Python bindings MUST exhibit the same verbatim-when-unmodified behavior: editing an unrelated cell preserves a table's unmodeled features, while mutating the table through the Python API takes the model path.

#### Scenario: test_table_totals_row_survives_cell_edit_rmw
- GIVEN an openpyxl-authored file with a totals-row table (a `totalsRowFunction` and a `<calculatedColumnFormula>`)
- WHEN it is opened with the oxlsx Python module, an UNRELATED cell is edited, and the file is saved
- THEN reopening with openpyxl shows the file opens with no repair AND the totals row and the calculated-column formula SURVIVED

#### Scenario: test_table_mutation_documents_feature_loss
- GIVEN the same file opened via the Python module
- WHEN the table is mutated through the Python API and the file is saved
- THEN reopening with openpyxl shows a valid file whose modeled fields (`displayName`, `ref`, columns) are correct (locking the documented (ii) feature-loss contract)

### Requirement: Python parity
PyO3 MUST expose `Table`, `TableStyleInfo`, `ws.add_table`, and `ws.tables` matching openpyxl import paths/usage.

#### Scenario: test_tables_round_trip
- GIVEN a saved xlsx with a table
- WHEN opened via the Python module
- THEN `ws.tables["Table1"]` returns the table with matching `ref`/columns

#### Scenario: test_tables_add_new
- GIVEN a worksheet with a header row
- WHEN `ws.add_table(Table(displayName="Table1", ref="A1:C3"))` is called and saved
- THEN reopening shows the table; a duplicate `displayName` raises `ValueError`

#### Scenario: test_tables_style_info
- GIVEN `TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)` set on a table
- WHEN saved and reopened
- THEN the style name and flags are preserved

## Test Mapping (1:1)
| # | Scenario | Test |
|---|----------|------|
|1|read existing|`tables_read_from_existing_workbook`|
|2|add round-trip|`tables_add_and_save_round_trip`|
|3|emits part+rel|`tables_add_emits_table_part_and_sheet_rel`|
|4|derived columns|`tables_columns_derived_from_header_row`|
|5|style round-trip|`tables_style_info_round_trip`|
|6|RMW preserve|`tables_rmw_preserves_existing_tables`|
|7|unique ids/rIds|`tables_multiple_tables_unique_ids_and_rids`|
|8|unmodified RMW preserves raw XML|`table_unmodified_rmw_preserves_raw_table_xml`|
|9|mutated RMW serializes from model|`table_mutated_rmw_serializes_from_model`|
|10|added table uses model path|`table_added_then_saved_uses_model_path`|
|11|removed table drops part/rel|`table_removed_drops_part_rel_and_tablepart`|
|12|mixed verbatim/model unique IDs|`mixed_verbatim_and_model_tables_unique_ids`|
|13|py round-trip|`test_tables_round_trip`|
|14|py add+dup|`test_tables_add_new`|
|15|py style|`test_tables_style_info`|
|16|py totals row survives RMW|`test_table_totals_row_survives_cell_edit_rmw`|
|17|py mutation documents feature loss|`test_table_mutation_documents_feature_loss`|
