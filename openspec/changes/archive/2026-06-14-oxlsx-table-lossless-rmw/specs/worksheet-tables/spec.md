# Delta for worksheet-tables — Lossless RMW of unmodeled table features

## Scope of this delta

**In scope**: preserve table features oxlsx does NOT model (`totalsRowCount`/`totalsRowShown`, per-column `totalsRowLabel`/`totalsRowFunction`/`calculatedColumnFormula`, `dxf` overrides, `autoFilter` criteria, `sortState`, any other `tableN.xml` attributes/children) across read-modify-write, via **verbatim retention** of the original `tableN.xml` bytes.

**Out of scope (unchanged from base)**: NO first-class API for totalsRow/sortState/autoFilter-criteria/dxf/calculatedColumnFormula; NO extras-bag parsing of unmodeled nodes; **charts, pivots OUT.**

**Design notes (normative intent):**
- Approach **(A) Verbatim-when-unmodified**: at read, retain each table's raw `tableN.xml` bytes in a parallel store on `Worksheet` (NOT as public `Table` fields, so external `Table { .. }` construction is unaffected). At save, re-emit verbatim when the table is clean; serialize from the model when dirty.
- **Linkage is by part filename + sheet rId**, never the internal `<table id>`. Verbatim bytes keep their original `id`/`displayName`; the writer MUST NOT rewrite ids inside retained bytes. Model-path table ids are allocated to avoid colliding with retained verbatim ids, preserving workbook-wide uniqueness.
- Verified vs openpyxl 3.1.2, the shapes to survive include `totalsRowCount` on `<table>`, `totalsRowLabel`/`totalsRowFunction` on `<tableColumn>`, `<calculatedColumnFormula>` child, and `<sortState><sortCondition/></sortState>`.

## MODIFIED Requirements

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

## ADDED Requirements

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

## Test Mapping (1:1)
| # | Scenario | Test | Kind |
|---|----------|------|------|
|1|table_unmodified_rmw_preserves_raw_table_xml|`table_unmodified_rmw_preserves_raw_table_xml`|Rust|
|2|table_mutated_rmw_serializes_from_model|`table_mutated_rmw_serializes_from_model`|Rust|
|3|table_added_then_saved_uses_model_path|`table_added_then_saved_uses_model_path`|Rust|
|4|table_removed_drops_part_rel_and_tablepart|`table_removed_drops_part_rel_and_tablepart`|Rust|
|5|mixed_verbatim_and_model_tables_unique_ids|`mixed_verbatim_and_model_tables_unique_ids`|Rust|
|6|test_table_totals_row_survives_cell_edit_rmw|`test_table_totals_row_survives_cell_edit_rmw`|Python|
|7|test_table_mutation_documents_feature_loss|`test_table_mutation_documents_feature_loss`|Python|

Supporting (non-mandatory) coverage: `dirty_tracking_getters_do_not_mark_dirty` (Rust) guards the read-only-no-dirty rule.
