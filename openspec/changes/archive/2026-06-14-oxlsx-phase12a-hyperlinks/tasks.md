# Tasks: Worksheet Hyperlinks (Phase 12a re-implementation)

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 420–460 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | 3-slice feature-branch chain |
| Delivery strategy | ask-on-risk |
| Chain strategy | feature-branch-chain |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: feature-branch-chain
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Model + rels foundation + byte-identity guard | Slice 1 | Targets `feature/phase12a-hyperlinks-reimpl`; scenarios 5; ~140 lines |
| 2 | Reader/writer round-trip + RMW + fixtures | Slice 2 | Targets PR #1 branch; scenarios 1-4, 6; ~180 lines |
| 3 | Python bindings + pytest | Slice 3 | Targets PR #2 branch; scenarios 7-8; ~100 lines |

---

## Slice 1: Model & Rels Foundation (~140 lines)

**Target branch**: `feature/phase12a-hyperlinks-reimpl` (off `main`)  
**Spec scenario coverage**: 5 (byte-identical rels guard)

### Phase 1.1: Foundation

- [x] 1.1.1 Create `src/hyperlink.rs` with `Hyperlink` struct: `target`, `location`, `tooltip`, `display` as `Option<String>` fields
- [x] 1.1.2 Add `pub fn new_external(target: &str) -> Self` constructor
- [x] 1.1.3 Add accessor `pub fn target(&self) -> Option<&str>`; defer `location`/`tooltip`/`display` accessors (preserved but not headline)
- [x] 1.1.4 Export `Hyperlink` from `src/lib.rs`

### Phase 1.2: SheetRelEntry Extension (critical for byte-identity)

- [x] 1.2.1 Add `pub target_mode: Option<String>` field to `SheetRelEntry` in `src/sheet_rels.rs`
- [x] 1.2.2 Update **all** `SheetRelEntry` construction sites (parser + upsert functions + tests) to explicitly set `target_mode: None` by default
- [x] 1.2.3 Modify `write_sheet_rels` to emit `TargetMode` attribute **only** when `target_mode.is_some()` — use `if let Some(mode) = &entry.target_mode { relationship.push_attribute(...) }`

### Phase 1.3: Hyperlink Rels Infrastructure

- [x] 1.3.1 Add `HYPERLINK_REL_TYPE` const to `src/sheet_rels.rs` (`http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`)
- [x] 1.3.2 Write `upsert_hyperlink_rels(existing: &[SheetRelEntry], targets: &[&str]) -> Vec<SheetRelEntry>`: filter out old hyperlink rels, allocate new rIds via `next_rel_id(merged)`, push entries with `target_mode: Some("External".to_string())`

### Phase 1.4: Testing (TDD — Scenario 5)

- [x] 1.4.1 **Write failing test** `existing_rels_remain_byte_identical_without_hyperlinks` (Rust integration): load `samples/tables_totals.xlsx`, RMW save to temp file, read both sheet rels as bytes, `assert_eq!` byte-for-byte (guards no regression on TargetMode default)
- [x] 1.4.2 Verify `cargo test existing_rels_remain_byte_identical_without_hyperlinks` **passes** after 1.2.3
- [x] 1.4.3 Unit test `upsert_hyperlink_rels` allocates distinct rIds after comments/tables: construct existing rels with comments (rId1) + table (rId3), upsert 2 hyperlinks → assert rId4, rId5
- [x] 1.4.4 `cargo clippy -- -D warnings && cargo fmt --check`

**Slice 1 deliverable**: Compiles; byte-identity guard passes; ready for reader/writer.

---

## Slice 2: Reader & Writer Round-Trip (~180 lines)

**Target branch**: off PR #1 branch (`feature/phase12a-hyperlinks-reimpl`)  
**Spec scenario coverage**: 1, 2, 3, 4, 6

### Phase 2.1: Cell Model Extension

- [x] 2.1.1 Add `pub(crate) write_hyperlink: Option<Hyperlink>` to `Cell` in `src/reader.rs`
- [x] 2.1.2 Add `pub fn hyperlink(&self) -> Option<&Hyperlink>` accessor to `Cell`
- [x] 2.1.3 Add `pub fn set_cell_hyperlink(&mut self, reference: &str, hyperlink: impl Into<Hyperlink>)` to `Worksheet` (mirrors `set_cell_font`)

### Phase 2.2: Reader — Parse Hyperlinks

- [x] 2.2.1 Define `RawHyperlink { ref, r_id, location, tooltip, display }` struct (module-private) in `src/reader.rs`
- [x] 2.2.2 In `parse_sheet`, add SAX match for `b"hyperlink"` element: extract `ref`, `r:id`, `location`, `tooltip`, `display` via `attr_value`. Return `Vec<RawHyperlink>` alongside `Worksheet`.
- [x] 2.2.3 In `Workbook::open`, after loading `original_sheet_rels` for each sheet, iterate its `RawHyperlink`s: if `r_id` is Some, resolve via `relation_target_for_id` (you may need a helper), construct `Hyperlink`, attach via `ws.set_cell_hyperlink`

### Phase 2.3: Writer — Emit Hyperlinks

- [x] 2.3.1 In `write_sheet_xml`, after the `mergeCells` block (L1203-1210), insert the `<hyperlinks>` emission: iterate cells with `write_hyperlink`, collect `(ref, Hyperlink)` pairs
- [x] 2.3.2 For external hyperlinks (target.is_some()), call `upsert_hyperlink_rels` in `write_xlsx_to_path` and get back the `r:id` map
- [x] 2.3.3 Emit `<hyperlinks xmlns:r="...">` start tag; for each hyperlink emit `<hyperlink ref="..." r:id="..." tooltip="..." display="..." location="..."/>` (only include attrs that are Some); write end tag
- [x] 2.3.4 In `write_xlsx_to_path`, extend the rels-write guard (`if has_comments(ws) || !sheet_part_numbers.is_empty()`) to also check `has_hyperlinks(ws)` (define helper)

### Phase 2.4: Fixture Authoring (openpyxl 3.1.2)

- [x] 2.4.1 **Create `samples/hyperlinks.xlsx`** via `.venv/bin/python`: sheet with A1 external hyperlink (`https://example.com`, tooltip="Example"), B1 internal location hyperlink (`location="Sheet!A1"`, no rId), C1:D1 merged cells (locks element order), save
- [x] 2.4.2 Verify fixture manually opens in Excel without repair

### Phase 2.5: Testing (TDD — Scenarios 1-4, 6)

- [x] 2.5.1 **Scenario 1** `hyperlink_read_from_existing_workbook`: load `samples/hyperlinks.xlsx`, assert `ws.cell("A1").unwrap().hyperlink().unwrap().target() == Some("https://example.com")`
- [x] 2.5.2 **Scenario 2** `hyperlink_write_new_workbook`: create workbook, `set_cell_hyperlink("A1", "http://rust-lang.org")`, save to temp, reopen, assert target round-trips + `<hyperlinks>` after `<mergeCells>` + rel `TargetMode="External"`
- [x] 2.5.3 **Scenario 3** `hyperlink_rmw_preserves_existing`: load fixture, edit B2 (unrelated), save, reopen, assert A1 hyperlink + tooltip preserved
- [x] 2.5.4 **Scenario 4** `hyperlink_coexists_with_comments_and_tables`: create sheet with comment (A1), table (A2:B3), hyperlink (A4), assert all round-trip, rIds unique
- [x] 2.5.5 **Scenario 6** `internal_location_hyperlink_preserved_on_rmw`: load fixture, assert B1 `location` preserved (no rId fabricated)
- [x] 2.5.6 `cargo clippy -- -D warnings && cargo fmt --check`

**Slice 2 deliverable**: Hyperlinks read/write/RMW; openpyxl regression passes; all Rust scenarios green.

---

## Slice 3: Python Bindings (~100 lines)

**Target branch**: off PR #2 branch  
**Spec scenario coverage**: 7, 8

### Phase 3.1: PyHyperlink Wrapper

- [x] 3.1.1 Define `#[pyclass(name = "Hyperlink")] pub struct PyHyperlink { #[pyo3(get, set)] pub target: Option<String>, ... }` in `src/python.rs`
- [x] 3.1.2 Add `impl From<&Hyperlink> for PyHyperlink` to clone fields
- [x] 3.1.3 Export `PyHyperlink` to Python module

### Phase 3.2: PyCell Hyperlink Access

- [x] 3.2.1 Add `#[getter] fn hyperlink(&self) -> Option<PyHyperlink>` to `PyCell` (maps `self.inner.hyperlink()`)
- [x] 3.2.2 Add `fn set_hyperlink(&self, target: &str) -> PyResult<()>` to `PyWorksheet` (locks workbook, calls `ws.set_cell_hyperlink`)

### Phase 3.3: Testing (TDD — Scenarios 7, 8)

- [x] 3.3.1 **Write pytest** `test_hyperlink_round_trip` (Scenario 7): set `ws["A1"].value = "link"`, `ws.set_hyperlink("A1", "http://example.com")`, save, reopen with openpyxl, assert target + no repair warning
- [x] 3.3.2 **Write pytest** `test_hyperlink_read` (Scenario 8): load `samples/hyperlinks.xlsx`, assert `ws["A1"].hyperlink.target == "https://example.com"`
- [x] 3.3.3 `cargo test && pytest tests/`

**Slice 3 deliverable**: Python API exposes hyperlinks; pytest green; all 8 scenarios complete.

---

## Dependencies & Blocking

- **Slice 1 → Slice 2**: Slice 2 builds on the model + rels; blocking.
- **Slice 2 → Slice 3**: Python bindings require the Rust round-trip; blocking.
- Within each slice, tasks are mostly sequential (TDD: write test → implement → verify).

## Rollback Plan

- Slice 1: revert the branch; no schema change (additive field with safe default).
- Slice 2: revert PR #2; RMW unaffected (Slice 1 still compiles).
- Slice 3: revert PR #3; Rust API remains functional.
