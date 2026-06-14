# Technical Design: worksheet-hyperlinks

## Architecture Overview

We re-implement `worksheet-hyperlinks` natively on the current base, introducing the `Hyperlink` model on `Cell`. We extend `SheetRelEntry` with a strictly backward-compatible `target_mode` field to support `TargetMode="External"` for hyperlinks without altering existing relationship XML. The reader extracts raw hyperlinks during SAX parsing and resolves their external targets once `original_sheet_rels` are loaded. The writer emits the `<hyperlinks>` schema block correctly ordered after `<mergeCells>`, allocating external relationship IDs sequentially after comments and tables to prevent collision.

## Data Model

**1. `src/hyperlink.rs` (New)**
```rust
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Hyperlink {
    pub(crate) target: Option<String>,
    pub(crate) location: Option<String>,
    pub(crate) tooltip: Option<String>,
    pub(crate) display: Option<String>,
}

impl Hyperlink {
    pub fn new_external(target: &str) -> Self { ... }
    pub fn target(&self) -> Option<&str> { ... }
    // Getters for location, tooltip, display
}
```
*Attached to `Cell` via `pub(crate) write_hyperlink: Option<Hyperlink>`.*

**2. `src/sheet_rels.rs` (Extension)**
```rust
pub struct SheetRelEntry {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: Option<String>, // NEW: Defaults to None for byte-identity
}
```

## Reader & Writer Flow

### Reader (`src/reader.rs`)
1. **SAX Parse**: In `parse_sheet`, match `<hyperlink>`. Extract `ref`, `r:id`, `location`, `tooltip`, `display`. Return these alongside the `Worksheet` as a temporary `RawHyperlink` struct.
2. **Resolution**: In `Workbook::open`, after populating `original_sheet_rels` for the sheet, iterate the extracted raw hyperlinks. If `r:id` is present, look up the `Target` in the sheet's rels.
3. **Attach**: Construct the final `Hyperlink` model and attach it to the cell via `ws.ensure_cell_mut(ref).write_hyperlink = Some(...)`.

### Writer (`src/writer.rs`)
1. **Extraction**: `write_xlsx_to_path` gathers all external hyperlink targets from `ws.cells()`.
2. **Upsert**: Call `sheet_rels::upsert_hyperlink_rels(&sheet_rels_entries, &targets)` which:
   - Drops any existing `HYPERLINK_REL_TYPE` entries (prevents staleness).
   - Allocates new rIds via `next_rel_id(merged)` (guaranteed to be *after* comments and tables).
   - Returns the updated rels and a mapping of `ref -> rId`.
3. **Emission**: `write_sheet_xml` emits `<hyperlinks>` **immediately after** `<mergeCells>` (and before `<pageMargins>` / `<legacyDrawing>`). It iterates the sheet's hyperlinks, injecting the allocated `r:id` for external links, and writes `location`, `tooltip`, `display` verbatim.
4. **Guard**: Include `has_hyperlinks(ws)` in the `if` block that triggers writing the `_rels/sheetN.xml.rels` file.

## Python Bindings (`src/python.rs`)
```rust
#[pyclass(name = "Hyperlink")]
#[derive(Clone)]
pub struct PyHyperlink {
    #[pyo3(get, set)] pub target: Option<String>,
    #[pyo3(get, set)] pub location: Option<String>,
    // ... tooltip, display
}
```
- **Getter**: `PyCell.hyperlink()` maps the cell's `Hyperlink` to a new `PyHyperlink`.
- **Setter**: `PyCell.set_hyperlink(target: &str)` updates the underlying worksheet cell (matching the ergonomics of `cell.hyperlink = "http..."`).

## Testing Strategy & Spec Mapping

Tests map 1:1 to the 8 scenarios defined in the spec:
1. `hyperlink_read_from_existing_workbook`: Load fixture, assert `cell.hyperlink().unwrap().target()`.
2. `hyperlink_write_new_workbook`: `set_cell_hyperlink`, RMW, assert target + `TargetMode`.
3. `hyperlink_rmw_preserves_existing`: Assert `tooltip`/`display` survive RMW untouched.
4. `hyperlink_coexists_with_comments_and_tables`: Multi-rel fixture, assert `next_rel_id` prevents collisions.
5. **Byte-Identity Guard** (`existing_rels_remain_byte_identical_without_hyperlinks`): RMW `tables_totals.xlsx`; assert `TargetMode` omission keeps legacy rels byte-for-byte identical.
6. `internal_location_hyperlink_preserved_on_rmw`: Assert `location` survives RMW.
7. `test_hyperlink_round_trip`: Pytest setting external string.
8. `test_hyperlink_read`: Pytest asserting returned `.target`.

*Fixture requirement*: Create `samples/hyperlinks.xlsx` via openpyxl containing both external (with tooltip) and internal (location-only) hyperlinks to act as the source of truth for RMW tests.

## Risks & Tradeoffs
- **Python Value Semantics**: `cell.hyperlink` returns a detached `PyHyperlink`. Modifying `cell.hyperlink.target = "foo"` in Python won't update the Rust cell. Users must use assignment `cell.hyperlink = "foo"` (which triggers the setter). This is a known tradeoff with PyO3 lightweight property wrappers vs openpyxl's deep proxy objects, matching `PyFont`.
- **XML Ordering**: If `<hyperlinks>` is misplaced, Excel requires repair. Strict ordering post-`mergeCells` prevents this.

## Delivery Plan (3-Slice Chain)
*At ~450 lines, this work is chunked into chained PRs.*
- **Slice 1 (Model & Rels)**: `src/hyperlink.rs` model, `SheetRelEntry.target_mode` extension, `upsert_hyperlink_rels`. (Scenario 5).
- **Slice 2 (Reader & Writer)**: `RawHyperlink` extraction, `write_sheet_xml` emission, RMW preservation, fixtures. (Scenarios 1-4, 6).
- **Slice 3 (Python)**: `PyHyperlink`, PyO3 getter/setters, pytest coverage. (Scenarios 7-8).