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
Opening a table-bearing file and saving MUST re-emit `<tableParts>` and retain the table parts + rels (regression for the latent drop bug). Unknown parts and unrelated sheet rels (comments/hyperlinks) MUST be preserved.

#### Scenario: tables_rmw_preserves_existing_tables
- GIVEN a file with an existing table (and comment/hyperlink rels)
- WHEN an unrelated cell is modified and the file is saved
- THEN the existing table is still referenced via `<tableParts>` (not orphaned), its part/rel remain, and unrelated rels + unknown parts are preserved

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
|8|py round-trip|`test_tables_round_trip`|
|9|py add+dup|`test_tables_add_new`|
|10|py style|`test_tables_style_info`|
