# Design: worksheet-bulk-write

## Technical Approach

The core challenge is that oxlsx stores cells in a sparse `HashMap<String, Cell>` keyed by A1 reference. Structural mutations (insert/delete rows/cols) require re-evaluating every cell's coordinate, discarding cells in deleted bands, and re-keying shifted cells. We will build a unified, deterministic `shift_cells` engine that performs a single $O(N)$ pass over the map.

This engine will also explicitly shift `merged_cells`, `row_heights`, and `column_widths` to exceed openpyxl's safety (preventing stale dimension maps).

To avoid silent data corruption (e.g., leaving a table pointing to the wrong cells), we implement a strict **Table Intersection Guard**: any operation whose affected band intersects a table's bounding box will proactively return an error.

The Python layer will adapt openpyxl's flexible dictionary/iterable `append` semantics into a contiguous, 0-indexed positional array of `Value`s (padded with `Value::Empty`) before passing it to the Rust core.

## Architecture Decisions

### Decision: Shift Engine Data Structure

**Choice**: Rebuild the `HashMap<String, Cell>` entirely during a shift.
**Alternatives considered**:
1. Mutate keys in-place (impossible in safe Rust HashMaps without remove/insert, prone to key collisions during the shift).
2. Change the internal model to `BTreeMap<u32, HashMap<u32, Cell>>` (row -> col -> Cell).
**Rationale**: Rebuilding into a fresh `HashMap` via `.drain()` ensures zero key collisions (e.g., when row 2 moves to row 3, it won't overwrite an existing row 3 before row 3 moves to row 4). It retains the SAX-friendly fast-lookup nature of the current `A1` string keys while maintaining $O(N)$ shift performance.

### Decision: Merged Cells & Dimensions Shift

**Choice**: Shift `merged_cells`, `row_heights`, and `column_widths`; drop merges that intersect a deleted band.
**Alternatives considered**: Match openpyxl 3.1.2 exactly (do not shift them at all).
**Rationale**: Openpyxl's behavior leaves the worksheet in a visually and structurally corrupted state. Shifting them is a predictable superset of openpyxl behavior that honors the SDD requirement to prevent data loss.

### Decision: Table Safety

**Choice**: Return `OxlsxError::TableIntersection` if a shift or move band overlaps `table.ref_range`.
**Alternatives considered**:
1. Silently shift the cells but leave the table ref alone (openpyxl parity, but corrupts the workbook on save).
2. Fully implement table reference shifting.
**Rationale**: Fully shifting table references and marking them dirty requires complex cross-module coordinate math and breaks the RMW verbatim path. Rejecting the operation upfront is safe, honors the verbatim pipeline, and cleanly defers proper table-shifting to a future phase.

## Data Flow

```text
[Python Client]
      │ `ws.append({'C': 1})`
      ▼
[PyO3 Binding (python.rs)] ──► Maps dict keys to 0-based indices
      │                        Pads missing with `Value::Empty`
      │ `[Empty, Empty, Number(1)]`
      ▼
[Rust Worksheet (reader.rs)]
      │ `append(impl IntoIterator<Item = Value>)`
      │ Computes `max_row() + 1`
      ▼
[Shift Engine / Mutator] ──► Writes to `HashMap<String, Cell>`
                             Updates `merged_cells`, bounds, etc.
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/reader.rs` | Modify | Add `append`, `insert_rows/cols`, `delete_rows/cols`, `move_range`. Add private `shift_engine` and table intersection guard. |
| `src/coords.rs` | Modify | Add `parse_a1(ref) -> (u32, u32)` and `parse_range(ref) -> (u32, u32, u32, u32)` helpers. |
| `src/python.rs` | Modify | Add `PyWorksheet` bindings for the 6 new methods. Add `append` dict-to-array padding logic using `py_to_value`. |
| `src/error.rs` | Modify | Add `OxlsxError::TableIntersection(String)` and `OxlsxError::NotImplemented(String)`. |
| `SPEC.md`, `README.md` | Modify | Uncheck Phase 10; update status to in-progress. |

## Interfaces / Contracts

```rust
// src/error.rs
pub enum OxlsxError {
    // ...
    #[error("operation intersects table '{0}', which is currently unsupported")]
    TableIntersection(String),
    #[error("not implemented: {0}")]
    NotImplemented(String),
}

// src/reader.rs
impl Worksheet {
    pub fn append(&mut self, values: impl IntoIterator<Item = Value>) {}
    pub fn insert_rows(&mut self, idx: u32, amount: u32) -> Result<(), OxlsxError> {}
    pub fn delete_rows(&mut self, idx: u32, amount: u32) -> Result<(), OxlsxError> {}
    pub fn insert_cols(&mut self, idx: u32, amount: u32) -> Result<(), OxlsxError> {}
    pub fn delete_cols(&mut self, idx: u32, amount: u32) -> Result<(), OxlsxError> {}
    pub fn move_range(&mut self, range: &str, rows: i32, cols: i32, translate: bool) -> Result<(), OxlsxError> {}
}
```

### Shift Algorithm (Row Example)
```rust
let mut new_cells = HashMap::new();
for (ref_str, mut cell) in self.cells.drain() {
    let (col, row) = coords::parse_a1(&ref_str)?;
    if row < start_idx {
        new_cells.insert(ref_str, cell);
    } else if amount < 0 && row < start_idx + (-amount) as u32 {
        // Drop cell (deleted band)
    } else {
        let new_row = (row as i32 + amount) as u32;
        let new_ref = coords::rc_to_a1(new_row, col);
        cell.reference = new_ref.clone();
        new_cells.insert(new_ref, cell);
    }
}
self.cells = new_cells;
```

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit (Rust) | Slice 1 scenarios (insert/delete, merges, table guard) | `#[test]` in `reader.rs`. Verify hashmap sizes, specific cell values via `.cell()`, and that `TableIntersection` is thrown correctly. |
| Unit (Rust) | Slice 2 scenarios (append, move_range) | Verify positional append lands at `max_row + 1`. Verify `move_range` clears source and returns error for `translate=True`. |
| Integration (Python) | Slice 3 scenarios (parity) | pytest: verify `append` dict/list padding logic. Use openpyxl fixture to perform operations in both engines and assert grid equality. |

## Migration / Rollout

No data migration required. The feature is purely additive. It changes the behavior of `merged_cells` and dimensions compared to openpyxl, but this is a documented strict improvement that requires no user intervention.

## Open Questions
- None. Table ref-shifting is explicitly deferred.

## Chained Delivery Plan

1. **Slice 1**: Core shift engine, `coords` range parsers, `insert/delete_rows/cols`, merges/dims shifting, table intersection guard. Rust tests (Scenarios 2, 3, 4, 5, 7, 8, 10).
2. **Slice 2**: `append` (Rust), `move_range`, `NotImplemented` error. Rust tests (Scenarios 1, 6, 9) + `SPEC.md`/`README.md` doc fixes.
3. **Slice 3**: `PyO3` bindings across all 6 methods. Python `append` iterable/dict mapping logic. Pytest cross-engine round-trips (Scenarios 11, 12, 13).
