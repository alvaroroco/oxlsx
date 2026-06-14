# worksheet-hyperlinks Specification

## Purpose

Read, write, and losslessly round-trip cell hyperlinks on the current main base, reusing `src/sheet_rels.rs`. Cut 1 targets external-URL hyperlinks via the `<hyperlinks>` worksheet element + external relationships; internal/location links and tooltip/display are preserved on RMW.

## In Scope

- Read external-URL hyperlinks (`<hyperlinks>` + `r:id` → external rel target).
- Write external hyperlinks (`<hyperlinks>` element + external rel via `upsert_hyperlink_rels`).
- `Hyperlink` model: `target` + preserved `tooltip` / `display` / `location`.
- Rust `cell.hyperlink()`, `ws.set_cell_hyperlink(...)`; Python `cell.hyperlink` object (`.target` + preserved fields).
- Lossless RMW; rId coexistence with comments/tables; byte-identical unrelated rels.

## Out of Scope (deferred)

Ergonomic authoring of internal `location` links (preserved only); auto-applying the Excel "Hyperlink" cell style; hyperlink removal API; charts/pivots. Resurrecting the old `relationships.rs` is **forbidden**.

## Implementation Notes

- **`target_mode` default**: `SheetRelEntry` gains `target_mode: Option<String>` defaulting to `None`; `write_sheet_rels` omits `TargetMode` when `None` so comment/table/printerSettings rels serialize byte-identically. External hyperlink rels set `Some("External")`.
- **rId order**: `upsert_hyperlink_rels` allocates rIds **after** comments + tables (`next_rel_id`); all sheet rIds stay unique.
- **Element position**: `<hyperlinks>` is emitted **after `<mergeCells>`** and before `<pageMargins>` (verified against openpyxl 3.1.2 CT_Worksheet order; the stale branch placed it wrong).
- **sheet_rels reuse**: add `HYPERLINK_REL_TYPE` (`.../relationships/hyperlink`) + `upsert_hyperlink_rels` mirroring `upsert_comment_rels`/`upsert_table_rels`; no second rels module.

## Requirements

### Requirement: Read external hyperlinks

The reader MUST parse `<hyperlinks>` and resolve each external `r:id` to its rel target, exposing it via `cell.hyperlink()`.

#### Scenario: Read external hyperlink from existing workbook

- GIVEN a fixture whose sheet has `<hyperlink ref="A1" r:id="rId1"/>` and a rel `Type=.../hyperlink Target="https://example.com" TargetMode="External"`
- WHEN the workbook is opened and `cell("A1").hyperlink()` is read
- THEN it returns `Some` with `target == "https://example.com"`

### Requirement: Write external hyperlinks

`set_cell_hyperlink` MUST attach an external hyperlink that serializes to a `<hyperlinks>` element plus an external rel. An empty target MUST NOT emit a hyperlink.

#### Scenario: Write external hyperlink to a new workbook

- GIVEN a new workbook with an external hyperlink set on a cell
- WHEN it is saved and reopened
- THEN the target round-trips, `<hyperlinks>` is present after `<mergeCells>`, and the rel has `TargetMode="External"`

### Requirement: Lossless RMW preservation

RMW MUST preserve hyperlinks and their `tooltip`/`display`/`location`, including internal `location`-only links, plus all unrelated parts and rels, when unrelated content is edited.

#### Scenario: RMW preserves an existing external hyperlink

- GIVEN a fixture with an external hyperlink carrying `tooltip` and `display`
- WHEN an unrelated cell is edited and the workbook is saved
- THEN the hyperlink, its tooltip/display, and all unrelated rels are preserved

#### Scenario: Internal location hyperlink preserved on RMW

- GIVEN a fixture with an internal `<hyperlink ref="F1" location="Sheet!A1"/>` (no rId)
- WHEN an unrelated edit is saved
- THEN the internal hyperlink survives verbatim (no rId fabricated)

### Requirement: rId coexistence with comments and tables

When a sheet has a hyperlink AND a comment AND a table, all parts MUST round-trip with unique, non-colliding rIds and remain a valid workbook.

#### Scenario: Hyperlink coexists with comments and tables

- GIVEN a sheet containing a hyperlink, a comment, and a table
- WHEN the workbook is saved and reopened
- THEN all three round-trip, every sheet rId is unique, and the file opens without repair

### Requirement: Unrelated rels remain byte-identical

Adding `target_mode` MUST NOT change serialization of comment/table/printerSettings rels.

#### Scenario: Existing rels byte-identical without hyperlinks

- GIVEN a fixture with comment/table/printerSettings rels and no hyperlinks
- WHEN it is RMW-saved
- THEN the sheet rels serialize byte-identically (preserving the no-repair byte-equality discriminators)

### Requirement: Python parity

Python `cell.hyperlink` MUST expose a `Hyperlink` object with `.target` (and preserved fields), settable by target.

#### Scenario: Python hyperlink round-trip

- GIVEN a cell with `hyperlink` target set via the Python API
- WHEN saved and reopened with openpyxl
- THEN the target survives and openpyxl reports no repair

#### Scenario: Python read hyperlink target

- GIVEN an existing file with an external hyperlink
- WHEN `cell.hyperlink.target` is read in Python
- THEN it equals the stored URL

## Test Mapping (1:1)

| # | Scenario | Test |
|---|----------|------|
| 1 | Read external | `hyperlink_read_from_existing_workbook` |
| 2 | Write external | `hyperlink_write_new_workbook` |
| 3 | RMW preserves | `hyperlink_rmw_preserves_existing` |
| 4 | Coexistence | `hyperlink_coexists_with_comments_and_tables` |
| 5 | Byte-identical rels | `existing_rels_remain_byte_identical_without_hyperlinks` |
| 6 | Internal preserved | `internal_location_hyperlink_preserved_on_rmw` |
| 7 | Py round-trip | `test_hyperlink_round_trip` |
| 8 | Py read | `test_hyperlink_read` |
