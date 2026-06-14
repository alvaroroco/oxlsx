# Design: Lossless RMW of unmodeled table features

## Technical Approach

To achieve verbatim-when-unmodified preservation without mutating the public `Table` struct, we decouple the table's in-memory data model from its retention state. We introduce a parallel `table_retention: Vec<Option<RawTable>>` on `Worksheet`. At read time, the exact `tableN.xml` bytes are retained here. Any API mutation (via `table_mut()`) conservatively marks the table dirty. 

At save time, we decouple the internal `<table id>` from the part filename sequence (`tableN.xml`). Verbatim tables keep their original internal `id` (written as-is via the retained bytes), and model tables receive newly allocated `id`s that avoid collisions. The part sequence, content-type overrides, and sheet-rel rIds are generated strictly sequentially across the workbook, completely independent of the internal table IDs.

## Architecture Decisions

### Decision: Retention State Location
**Choice**: A parallel `table_retention: Vec<Option<RawTable>>` on `Worksheet`.
**Alternatives considered**: Private fields on `Table` (rejected: breaks public struct-literal construction `Table { .. }` used by PyO3/Rust consumers); `HashMap<String, RawTable>` keyed by `displayName` (rejected: brittle if a user mutates the `displayName` via `table_mut`).
**Rationale**: A parallel `Vec` aligns 1:1 with `ws.tables`. Operations like `add_table` and `remove_table` mutate both vectors at the same index, keeping them perfectly synchronized regardless of internal `Table` mutations.

### Decision: Decoupling Part Number from Internal `<table id>`
**Choice**: The writer assigns sequential part numbers (`1..=total_tables`) for filenames/rIds/content-types, but allocates internal `table.id`s dynamically to avoid colliding with verbatim tables.
**Alternatives considered**: Parse and rewrite the internal `id="..."` attribute inside the raw verbatim bytes (rejected: violates the "verbatim" requirement, risks schema corruption or XML parsing failures on edge cases).
**Rationale**: Excel linkages rely on the sheet-rel `Target` (the part filename) and `rId` in `<tableParts>`. The internal `<table id>` only needs to be workbook-unique.

### Decision: Dirty-Tracking Strategy
**Choice**: `Worksheet::table_mut` marks the table dirty immediately upon access.
**Alternatives considered**: Granular dirty tracking via getters/setters on every `Table` field (rejected: over-engineered, requires making `Table` fields private which breaks existing API).
**Rationale**: Acquiring a mutable reference (`&mut`) implies mutation intent. This is simple, safe, and guarantees we never emit stale verbatim bytes if the model was touched.

## Data Flow

1. **Read Path (`reader.rs`)**:
   - Locate `.../table` target. Read the part fully into `bytes: Vec<u8>`.
   - Parse `bytes` via a `Cursor` into a `Table`.
   - `ws.tables.push(table)` and `ws.table_retention.push(Some(RawTable { bytes, dirty: false }))`.
2. **Mutation Path (`reader.rs` / `python.rs`)**:
   - Read-only getters (`tables()`) access `ws.tables` safely.
   - `table_mut("Name")` locates index `i`, sets `ws.table_retention[i].dirty = true`, and returns `&mut ws.tables[i]`.
   - Python property setters internally call `table_mut`, inheriting dirty tracking.
   - `ws.add_table(t)` pushes `t` to `tables` and `None` to `table_retention`.
3. **Save Path (`writer.rs`)**:
   - `assign_table_ids` collects all verbatim IDs into a `HashSet`, then allocates `next_free` IDs to all model/dirty tables.
   - The writer loops `sheets` and `tables`, incrementing a global `part_number`.
   - For verbatim: write `raw.bytes` to `xl/tables/table{part_number}.xml`.
   - For model: assign the pre-allocated `id`, run `serialize_table_xml`, write output to `xl/tables/table{part_number}.xml`.

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/reader.rs` | Modify | Add `table_retention` to `Worksheet`. Update `parse_table_xml` call to capture raw bytes. Add `remove_table` method. Hook `table_mut` to flip `dirty = true`. |
| `src/table.rs` | Modify | Add `pub(crate) struct RawTable { pub bytes: Vec<u8>, pub dirty: bool }`. |
| `src/writer.rs` | Modify | Rename `assign_table_numbers` to `assign_table_ids` and implement verbatim-aware collision avoidance. Emit loop at `358-372` branches on verbatim vs model. Update `target_refs` to use sequential part numbers. |
| `src/python.rs` | Modify | Add `remove_table` / `__delitem__` support to `PyTableList` to map to scenario 4. |

## Interfaces / Contracts

```rust
// In src/table.rs
#[derive(Debug, Clone)]
pub(crate) struct RawTable {
    pub bytes: Vec<u8>,
    pub dirty: bool,
}

// In src/reader.rs (Worksheet)
pub(crate) tables: Vec<Table>,
pub(crate) table_retention: Vec<Option<RawTable>>,

// Mutation contract
pub fn table_mut(&mut self, display_name: &str) -> Option<&mut Table> {
    let idx = self.tables.iter().position(|t| t.display_name.eq_ignore_ascii_case(display_name))?;
    if let Some(Some(raw)) = self.table_retention.get_mut(idx) {
        raw.dirty = true;
    }
    Some(&mut self.tables[idx])
}

pub fn remove_table(&mut self, display_name: &str) -> Option<Table> {
    let idx = self.tables.iter().position(|t| t.display_name.eq_ignore_ascii_case(display_name))?;
    self.table_retention.remove(idx);
    Some(self.tables.remove(idx))
}
```

## Testing Strategy

| Layer | What to Test (Mapped to Scenarios) | Approach |
|-------|------------------------------------|----------|
| Unit | `table_unmodified_rmw_preserves_raw_table_xml` | Open a mock bytes fixture with totalsRow, edit cell, check bytes. |
| Unit | `table_mutated_rmw_serializes_from_model` | Mutate style via `table_mut`, verify totalsRow is gone. |
| Unit | `table_added_then_saved_uses_model_path` | `add_table`, verify output validity. |
| Unit | `table_removed_drops_part_rel_and_tablepart` | `remove_table`, verify `xl/tables/` omission. |
| Unit | `mixed_verbatim_and_model_tables_unique_ids` | Mix retained and added tables, assert `<table id>` disjointness. |
| Unit | `dirty_tracking_getters_do_not_mark_dirty` | Assert `table_retention[i].dirty == false` after `tables()`. |
| Python E2E | `test_table_totals_row_survives_cell_edit_rmw` | Round-trip real openpyxl totalsRow file in `.venv`, verify `openpyxl.load_workbook` repair-free load + totals features present. |
| Python E2E | `test_table_mutation_documents_feature_loss` | Mutate ref_range in Python, assert loss in openpyxl. |

## Open Questions
- None.

## Sizing / Rollout
~150-250 lines total. A **single unified PR** is strongly recommended, as breaking this down into verbatim/retention (Unit 1) and dirty-tracking (Unit 2) creates a midpoint commit where users can't actually mutate tables cleanly without a temporary break in behavior. 
