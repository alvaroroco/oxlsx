# Worksheet Comments

## Purpose

Plain text worksheet comment read/write behavior with author metadata and minimum OOXML plumbing for Excel-compatible save/reload.

## Requirements

### Requirement: Expose plain text cell comments

The system exposes `cell.comment` with plain text comment text and author metadata when a worksheet comment exists. Plain text comments only are in scope; rich text formatting is not required.

**Scenario: comment_round_trip**

- GIVEN a workbook containing a plain text comment with an author on a cell
- WHEN the workbook is opened, saved, and reopened
- THEN the same cell exposes the same comment text and author metadata

**Scenario: test_comment_round_trip**

- GIVEN a workbook containing a plain text comment with an author on a cell
- WHEN Python opens the workbook and reads `cell.comment`
- THEN the comment text and author match the workbook contents

### Requirement: Assign plain text comments

The system allows assigning a plain text comment with author metadata to a worksheet cell through `ws.set_cell_comment(...)` or an equivalent API aligned with openpyxl ergonomics.

**Scenario: comment_write_new_cell**

- GIVEN a workbook with a cell that has no comment
- WHEN Rust assigns a plain text comment and author, saves, and reloads the workbook
- THEN the target cell persists the same comment text and author metadata

**Scenario: test_comment_write_new_cell**

- GIVEN a workbook with a cell that has no comment
- WHEN Python assigns a plain text comment and author, saves, and reloads the workbook
- THEN the target cell exposes the persisted comment text and author metadata

### Requirement: Serialize comment package plumbing safely

The system serializes per-sheet comment XML, required sheet relationship entries, and minimum legacy drawing/VML plumbing sufficient for Excel to open the saved file. In read-modify-write flows, unknown existing sheet relationships are preserved unchanged.

**Scenario: sheet_comment_relationships_preserved**

- GIVEN a workbook whose sheet relationships include comments and unrelated relationship entries
- WHEN the workbook is opened and saved without changing those unrelated relationships
- THEN comment targets still round-trip and unknown existing sheet relationships remain preserved

**Scenario: Excel-compatible comment package**

- GIVEN a workbook where a plain text comment is added or updated
- WHEN the workbook is saved
- THEN the package contains the minimum comment XML and legacy drawing/VML plumbing needed for Excel to open the file
