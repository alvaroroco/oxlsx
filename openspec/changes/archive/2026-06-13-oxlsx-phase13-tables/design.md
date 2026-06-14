# Technical Design: Worksheet Tables (`ws.add_table` / `ws.tables`)

## Architecture Overview
Tables will be modeled as fully-owned structures attached to the `Worksheet`. Following the Phase 12b comments pattern, tables are parsed via `sheet_rels.rs` during workbook loading and fully modeled in memory. On save, they are serialized to new `xl/tables/tableN.xml` parts, integrated into the `[Content_Types].xml`, wired to the sheet via `sheet_rels.rs`, and linked inside the sheet XML via `<tableParts>`. To fix the latent RMW orphan bug, `writer.rs` claims ownership of all table parts, ensuring existing tables are parsed, retained, and freshly re-emitted with a valid `<tableParts>` closing tag.

## Data Model (`src/table.rs`)
```rust
pub struct Table {
    pub id: u32,
    pub name: String,
    pub display_name: String,
    pub ref_range: String, // "A1:D5"
    pub header_row_count: u32,
    pub columns: Vec<TableColumn>,
    pub style: Option<TableStyleInfo>,
}

pub struct TableColumn {
    pub id: u32,
    pub name: String,
}

pub struct TableStyleInfo {
    pub name: String,
    pub show_first_column: bool,
    pub show_last_column: bool,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
}
```
*Constraints*: PyO3-ready, no public lifetimes, no `get_` prefixes.

## Implementation Plan

### 1. Read Flow & Worksheet State (`src/reader.rs`)
- Add `pub(crate) tables: Vec<Table>` to `Worksheet`. Expose via `pub fn tables(&self) -> &[Table]`.
- During `Workbook` parsing, after `parse_sheet_rels`:
  - Find entries matching `TABLE_REL_TYPE` (`".../relationships/table"`).
  - Resolve targets (e.g., `../tables/table1.xml`), open ZIP entries, parse using `table::parse_table_xml`.
  - Add to `ws.tables`.

### 2. Write Flow & The RMW Orphan Fix (`src/writer.rs`)
- **Ownership**: Update `is_writer_owned_entry` to return `true` for `xl/tables/table*.xml`. This prevents `copy_unknown_entries` from blindly copying orphaned tables.
- **ID Allocation (Workbook-wide)**: Before the sheet loop, collect all existing `table.id`s across all sheets. Assign `max_id + 1` to any table missing an ID.
- **rId Allocation (Sheet-local)**: Reuse `sheet_rels.rs`. Add `upsert_table_rels` (mirroring `upsert_comment_rels`) to allocate non-colliding `rId`s while preserving unrelated hyperlink/comment rels.
- **Part Emission**: Loop through `ws.tables`, write `xl/tables/tableN.xml` (incrementing N workbook-wide).
- **Content Types**: Add override `<Override PartName="/xl/tables/tableN.xml" ContentType="...spreadsheetml.table+xml"/>`.
- **Sheet XML (`<tableParts>`)**: At the END of `write_sheet_xml`, strictly after `<pageMargins>` and `<pageSetup>`, just before `</worksheet>`, emit `<tableParts count="N"><tablePart r:id="rIdX"/>...`.

### 3. Add Table & Column Derivation (`src/table.rs` & `Worksheet`)
- `Worksheet::add_table(&mut self, mut table: Table) -> Result<(), OxlsxError>`.
- **Derivation**: If `table.columns` is empty, parse the header row from `ref_range` (Row 1). Iterate columns. Read `ws.cell(...)` to get the header string. Assign `id: 1..N`. Fallback for empty cells: `"Column{i}"`.

### 4. Duplicate Display Name Check & Python Parity (`src/python.rs`)
- **Python Check**: `PyWorksheet::add_table` inspects `self.workbook().worksheets()` to collect all `display_name`s workbook-wide. If a duplicate is found, it raises `ValueError("Table with name X already exists")` immediately, matching openpyxl exactly.
- **PyO3 Model**: `PyTable`, `PyTableStyleInfo`.
- **PyTableList**: Implements `__getitem__`, `__iter__`, `__len__`, `keys()` returning `display_name`s to emulate openpyxl's dict-like `ws.tables`.

## 3-Unit Chain Plan
1. **Unit 1 — Rust READ**: `table.rs` model & parser, `Worksheet::tables`, `reader.rs` wiring. (Covers Scenario 1)
2. **Unit 2 — Rust WRITE + RMW fix**: `table.rs` serializer, column derivation, workbook-wide ID allocator, `writer.rs` tableParts/CT/emission, `sheet_rels.rs` additions. (Covers Scenarios 2, 3, 4, 5, 6, 7)
3. **Unit 3 — Python Parity**: `PyTable`, `PyTableStyleInfo`, `PyTableList`, `add_table` with workbook-wide duplicate checking, pytest suite. (Covers Scenarios 8, 9, 10)

## Out of Scope (Deferred)
dxf overrides (`headerRowDxfId`, etc.), `totalsRow`, calculated column formulas, autoFilter sort/filter criteria (only `<autoFilter ref>` is emitted), `sortState`. Charts & pivots remain entirely out of scope.

## Risks & Edge Cases
- **Excel Strictness**: Missing `id`, `name`, or `displayName` breaks the file. Mitigated by enforcing them in the `Table` struct and ensuring sequential ID allocation.
- **RMW Lossy Parse**: Existing tables with `totalsRow` will be read into our `Table` model and serialized back *without* `totalsRow`. This is acceptable for a "first cut" as Excel gracefully handles the missing tags without repair flags.
- **rId Collisions**: Safely mitigated by `id_for_type_or_next` inside `upsert_table_rels`.
