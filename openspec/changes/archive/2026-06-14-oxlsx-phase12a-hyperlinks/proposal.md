# Proposal: Worksheet Hyperlinks (Phase 12a, re-implemented)

## Intent

Expose `cell.hyperlink` (openpyxl parity) on the **current** main base, reusing `src/sheet_rels.rs`. A prior branch implemented this on a stale phase-9 base with its own obsolete `src/relationships.rs`; that branch is reference-only. This re-implements clean and makes main feature-complete for v0.1.0.

## Scope

### In Scope
- Read external-URL hyperlinks: parse `<hyperlinks>`, resolve `r:id`→external target via sheet rels.
- Write external hyperlinks: emit `<hyperlinks>` + external rels (`TargetMode="External"`) via new `upsert_hyperlink_rels`.
- `Hyperlink` model (`src/hyperlink.rs`) carrying `target` + preserved `tooltip`/`display`/`location` (lossless RMW).
- Rust `cell.hyperlink() -> Option<&Hyperlink>`, `set_cell_hyperlink(ref, target)`; Python `cell.hyperlink` (object w/ `.target`).
- RMW losslessness: hyperlinks round-trip; rIds coexist with comments/tables (no collision).

### Out of Scope (deferred)
- Ergonomic authoring of internal `location` links (preserved on RMW, not a headline setter).
- Auto-applying Excel "Hyperlink" cell style (blue/underline).
- Hyperlink removal API. Resurrecting old `relationships.rs` — **forbidden**.

## Capabilities

### New Capabilities
- `worksheet-hyperlinks`: read/write/round-trip cell hyperlinks with rel coexistence.

### Modified Capabilities
- None (additive on `Cell`/`Worksheet`).

## Approach

| Layer | Change |
|-------|--------|
| `sheet_rels.rs` | `HYPERLINK_REL_TYPE` const; add `target_mode: Option<String>` to `SheetRelEntry`; emit `TargetMode` in `write_sheet_rels`; `upsert_hyperlink_rels` allocating rIds **after** comments+tables. |
| `reader.rs` | `Cell.write_hyperlink` field + accessor; parse `<hyperlinks>` (ref, r:id, tooltip, display, location); resolve via parsed `sheet_rels`. |
| `writer.rs` | Emit `<hyperlinks>` **after** `mergeCells` (correct CT_Worksheet order — old branch wrongly placed it before); slot `upsert_hyperlink_rels` into the existing merge pipeline (L318-413); extend rels-part write guard to include hyperlink-only sheets. |
| `python.rs` | `cell.hyperlink` getter + `set_cell_hyperlink`. |

No new ZIP part (unlike comments/tables).

## Risks

| Risk | Likelihood | Mitigation |
|------|-----------|------------|
| `TargetMode` extension breaks existing rels byte-equality | Med | `target_mode=None` default emits identical XML; assert on openpyxl fixtures. |
| rId collision w/ comments/tables on one sheet | Med | Allocate after both via `next_rel_id`; coexistence test. |
| Dropping tooltip/display → non-lossless | Med | Model + parse/emit all attrs. |
| Python string-only deviates from openpyxl | Low | Return object w/ `.target` from cut 1. |

## Rollback

Additive only — revert the feature commits/branch; no schema or data migration.

## Affected Areas

`src/hyperlink.rs` (new), `src/sheet_rels.rs`, `src/reader.rs`, `src/writer.rs`, `src/python.rs`, `src/lib.rs`, integration + pytest, new external-link fixture.

## Success Criteria

- [ ] External hyperlink read + write round-trips (`cargo test`).
- [ ] Real openpyxl-authored fixture round-trips lossless (rels + tooltip/display preserved, no Excel repair).
- [ ] Comments/tables + hyperlinks on one sheet: distinct rIds, all parts intact.
- [ ] `cargo clippy -D warnings` + `cargo fmt --check` clean; pytest green.

## Sizing & Delivery

~450 lines total → **at the 400-line budget**. Recommend a **3-slice chain**: (1) `hyperlink.rs` + `sheet_rels.rs` rels/TargetMode + unit tests; (2) reader+writer round-trip + openpyxl regression; (3) Python bindings + pytest.

## Next Step

`sdd-spec` for `worksheet-hyperlinks`. After landing, delete stale `feat/phase12a-hyperlinks*` branches.
