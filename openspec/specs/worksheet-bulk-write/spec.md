# worksheet-bulk-write Specification

## Purpose

Programmatic row/column manipulation on `Worksheet`, mirroring openpyxl 3.1.2:
`append`, `insert_rows`, `delete_rows`, `insert_cols`, `delete_cols`, `move_range`.
This is the real Phase 10 — these methods do not yet exist in `src/`, while `SPEC.md`/`README.md`
incorrectly mark Phase 10 "✅ Done".

## Scope

### In Scope
- The six methods above on the Rust `Worksheet` + PyO3 parity.
- A deterministic shift engine that rebuilds `cells: HashMap<String, Cell>` keyed by A1 via `coords.rs`.
- Shifting `merged_cells`, `row_heights`, `column_widths` along with cells.
- A table-intersection guard that errors instead of corrupting.

### Out of Scope
- `move_range(translate=True)` — formula translation is unsupported (returns an error).
- Formula reference rewriting on insert/delete (openpyxl does NOT translate — parity).
- Auto-shifting `print_area`, `print_title_rows`, `print_title_cols`, `defined_names` (left as-is — openpyxl parity).
- Proper table-ref shifting (deferred; the guard rejects such ops for now).

## Design Notes (behavioral, non-binding on internals)

- **HashMap rebuild**: a shift parses each cell's A1 key via `coords::column_index` + row digits; cells in/after the affected band get a recomputed key via `coords::rc_to_a1`; the map is rebuilt deterministically. `value`, `style_id`, `comment`, and `hyperlink` ride on `Cell`, so they move with it. Bounds (`max_row`/`max_column`/...) recompute from the rebuilt scan.
- **merges/dims superset divergence**: openpyxl 3.1.2 leaves `merged_cells` and row/column dimensions UNCHANGED on insert/delete (verified). oxlsx INTENTIONALLY DIVERGES by shifting them, to avoid stale/corrupt merges and misplaced dimensions. This is a documented, deliberate superset of openpyxl behavior.
- **Table guard**: shifting cells under a retained table would either corrupt the table or force the verbatim path to ship a stale `ref`. Until proper ref-shifting exists, any op whose affected band intersects a `table.ref_range` returns `Err(OxlsxError)` with the model unchanged.

## Requirements

### Requirement: Append a row at the bottom

`ws.append(iterable)` MUST write a new row at `max_row + 1`. The Rust core MUST accept a positional iterator placing values from column A rightward. The Python layer MUST accept a list/tuple/generator (positional from column A) AND a dict whose keys are column letters (e.g. `"A"`) OR 1-based integer indices (e.g. `2`), with mixed keys allowed.

#### Scenario: append_adds_row_at_bottom

- GIVEN a worksheet whose `max_row` is N
- WHEN `append([v1, v2, v3])` is called
- THEN the values land in `A{N+1}`, `B{N+1}`, `C{N+1}`
- AND `max_row` becomes `N+1`

#### Scenario: test_append_list_and_dict

- GIVEN an empty worksheet (Python binding)
- WHEN appending a list, then a dict with letter keys, int keys, and mixed keys
- THEN each value lands in the column its key/position resolves to (letters and 1-based ints agree)
- AND the resulting grid matches openpyxl's `append` output for the same inputs

### Requirement: Insert blank rows and columns

`ws.insert_rows(idx, amount=1)` and `ws.insert_cols(idx, amount=1)` MUST use 1-based `idx`, default `amount` to 1, and shift existing cells down/right by `amount`, leaving the inserted band blank.

#### Scenario: insert_rows_shifts_cells_down

- GIVEN cells in rows 1..4
- WHEN `insert_rows(2, 2)` is called
- THEN former rows 2..4 move to rows 4..6
- AND rows 2 and 3 are blank, row 1 is unchanged

#### Scenario: insert_cols_shifts_cells_right

- GIVEN cells in columns A..D of row 1
- WHEN `insert_cols(2, 1)` is called
- THEN former columns B..D move to C..E
- AND column B is blank, column A is unchanged

### Requirement: Delete rows and columns

`ws.delete_rows(idx, amount=1)` and `ws.delete_cols(idx, amount=1)` MUST use 1-based `idx`, default `amount` to 1, remove the cells in the band, and shift remaining cells up/left by `amount`.

#### Scenario: delete_rows_shifts_cells_up

- GIVEN cells in rows 1..5
- WHEN `delete_rows(2, 1)` is called
- THEN former rows 3..5 move to rows 2..4
- AND the original row-2 values are gone, row 1 unchanged

#### Scenario: delete_cols_shifts_cells_left

- GIVEN cells in columns A..D of row 1
- WHEN `delete_cols(2, 1)` is called
- THEN former columns C..D move to B..C
- AND the original column-B values are gone, column A unchanged

### Requirement: Preserve formula text verbatim

Insert, delete, and `move_range` operations MUST move formula cell values WITHOUT translating their references (openpyxl parity). The formula string is carried unchanged to the cell's new position.

#### Scenario: delete_rows_does_not_translate_formulas

- GIVEN a formula cell `B3` whose text is `=A3+1`
- WHEN `delete_rows(2, 1)` shifts `B3` up to `B2`
- THEN the cell at `B2` still holds the verbatim text `=A3+1` (no rewrite to `=A2+1`)

### Requirement: Move a cell range

`ws.move_range(cell_range, rows=0, cols=0, translate=False)` MUST relocate the block by `(rows, cols)`, clear source cells not overwritten by the destination, and carry formulas verbatim. `translate=True` is unsupported and MUST return an error.

#### Scenario: move_range_relocates_block

- GIVEN a populated block `A1:B2` containing values and a formula
- WHEN `move_range("A1:B2", rows=2, cols=1)` is called with default `translate=False`
- THEN the block content appears at `C3:D4`
- AND vacated source cells are empty AND formula text is unchanged

#### Scenario: move_range_translate_true_is_unsupported

- GIVEN any worksheet
- WHEN `move_range(range, rows, cols, translate=True)` is called
- THEN it returns `Err(OxlsxError)` (NotImplemented) and the model is unchanged

#### Scenario: test_move_range

- GIVEN a worksheet via the Python binding
- WHEN `move_range(..., translate=False)` is applied
- THEN the resulting grid matches openpyxl's `move_range` output for the same inputs

### Requirement: Shift merged cells and dimensions with bulk ops

Insert/delete operations MUST shift `merged_cells`, `row_heights`, and `column_widths` consistently with the cell shift. Merged ranges fully past the band move by `amount`; a merged range only partially straddling a deleted band MUST be dropped (no corruption). This intentionally diverges from openpyxl 3.1.2, which leaves them unchanged.

#### Scenario: delete_rows_shifts_merged_cells_and_dimensions

- GIVEN a merged range below the deletion point and a `row_heights` entry below it
- WHEN `delete_rows(idx, amount)` is called
- THEN the merged range and the row-height key both shift up by `amount`
- AND any merge that only partially overlaps the deleted band is removed

### Requirement: Guard tables against corrupting shifts

Any `insert_rows`/`insert_cols`/`delete_rows`/`delete_cols`/`move_range` whose affected band intersects an existing `table.ref_range` MUST return `Err(OxlsxError)` and leave the workbook unchanged. No silent corruption and no verbatim-path stale `ref`.

#### Scenario: bulk_op_intersecting_table_returns_err

- GIVEN a worksheet with a table over `A1:C5`
- WHEN an insert/delete/move whose band intersects `A1:C5` is requested
- THEN the call returns `Err(OxlsxError)`
- AND the worksheet (cells, merges, dimensions, table) is byte-for-byte unchanged

### Requirement: Python parity for insert/delete round-trip

The Python bindings MUST expose `insert_rows`/`delete_rows` (and the column variants) with openpyxl-identical signatures and produce grids equivalent to openpyxl after the same sequence of operations.

#### Scenario: test_delete_and_insert_rows

- GIVEN identical starting data in oxlsx and openpyxl
- WHEN the same insert and delete sequence runs in both
- THEN reading the oxlsx result back yields the same cell grid as openpyxl

### Requirement: Correct stale Phase 10 documentation

`SPEC.md` and `README.md` MUST stop claiming "Phase 10 — Bulk write API ✅ Done" until these methods exist and pass tests. The status MUST reflect in-progress/planned, and the README checklist item MUST be unchecked.

#### Scenario: phase10_docs_reflect_reality

- GIVEN the current docs marking Phase 10 done
- WHEN this change ships
- THEN `SPEC.md` line for Phase 10 is `🔲`/in-progress and `README.md` shows the bulk-write item unchecked

## Test Mapping (1:1)

| # | Test | Kind | Requirement |
|---|------|------|-------------|
| 1 | `append_adds_row_at_bottom` | Rust | Append a row at the bottom |
| 2 | `insert_rows_shifts_cells_down` | Rust | Insert blank rows and columns |
| 3 | `delete_rows_shifts_cells_up` | Rust | Delete rows and columns |
| 4 | `insert_cols_shifts_cells_right` | Rust | Insert blank rows and columns |
| 5 | `delete_cols_shifts_cells_left` | Rust | Delete rows and columns |
| 6 | `move_range_relocates_block` | Rust | Move a cell range |
| 7 | `delete_rows_shifts_merged_cells_and_dimensions` | Rust | Shift merged cells and dimensions |
| 8 | `bulk_op_intersecting_table_returns_err` | Rust | Guard tables against corrupting shifts |
| 9 | `move_range_translate_true_is_unsupported` | Rust | Move a cell range |
| 10 | `delete_rows_does_not_translate_formulas` | Rust | Preserve formula text verbatim |
| 11 | `test_append_list_and_dict` | Python | Append a row at the bottom |
| 12 | `test_delete_and_insert_rows` | Python | Python parity for insert/delete round-trip |
| 13 | `test_move_range` | Python | Move a cell range |
