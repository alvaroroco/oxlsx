# worksheet-comments Specification

## Purpose

Define plain text worksheet comment read/write behavior with author metadata and minimum OOXML plumbing for Excel-compatible save/reload.

## Requirements

### Requirement: Expose plain text cell comments

The system MUST expose `cell.comment` with plain text comment text and author metadata when a worksheet comment exists. Plain text comments only are in scope; rich text formatting MUST NOT be required in this phase.

#### Scenario: comment_round_trip

- GIVEN a workbook containing a plain text comment with an author on a cell
- WHEN the workbook is opened, saved, and reopened
- THEN the same cell MUST expose the same comment text and author metadata
- AND the scenario MUST map 1:1 to Rust `#[test] comment_round_trip`

#### Scenario: test_comment_round_trip

- GIVEN a workbook containing a plain text comment with an author on a cell
- WHEN Python opens the workbook and reads `cell.comment`
- THEN the comment text and author MUST match the workbook contents
- AND the scenario MUST map 1:1 to Python `def test_comment_round_trip`

### Requirement: Assign plain text comments

The system MUST allow assigning a plain text comment with author metadata to a worksheet cell through `ws.set_cell_comment(...)` or an equivalent API aligned with openpyxl ergonomics.

#### Scenario: comment_write_new_cell

- GIVEN a workbook with a cell that has no comment
- WHEN Rust assigns a plain text comment and author, saves, and reloads the workbook
- THEN the target cell MUST persist the same comment text and author metadata
- AND the scenario MUST map 1:1 to Rust `#[test] comment_write_new_cell`

#### Scenario: test_comment_write_new_cell

- GIVEN a workbook with a cell that has no comment
- WHEN Python assigns a plain text comment and author, saves, and reloads the workbook
- THEN the target cell MUST expose the persisted comment text and author metadata
- AND the scenario MUST map 1:1 to Python `def test_comment_write_new_cell`

### Requirement: Serialize comment package plumbing safely

The system MUST serialize per-sheet comment XML, required sheet relationship entries, and minimum legacy drawing/VML plumbing sufficient for Excel to open the saved file. In read-modify-write flows, unknown existing sheet relationships MUST be preserved unchanged.

#### Scenario: sheet_comment_relationships_preserved

- GIVEN a workbook whose sheet relationships include comments and unrelated relationship entries
- WHEN the workbook is opened and saved without changing those unrelated relationships
- THEN comment targets MUST still round-trip and unknown existing sheet relationships MUST remain preserved
- AND the scenario MUST map 1:1 to Rust `#[test] sheet_comment_relationships_preserved`

#### Scenario: Excel-compatible comment package

- GIVEN a workbook where a plain text comment is added or updated
- WHEN the workbook is saved
- THEN the package MUST contain the minimum comment XML and legacy drawing/VML plumbing needed for Excel to open the file
- AND rich text comment fidelity, formatting, and dimensions MAY remain unsupported in this phase
