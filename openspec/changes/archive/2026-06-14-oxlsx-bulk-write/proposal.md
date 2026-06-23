# Proposal: Worksheet Bulk Write API (openpyxl parity)

## Intent

Real openpyxl→oxlsx migrations break: `Worksheet.delete_rows`/`delete_cols` (plus `append`, `insert_rows`, `insert_cols`, `move_range`) are CALLED by users but **do not exist** in `src/`. Audit confirms zero definitions across the crate.

`SPEC.md` and `README.md` falsely mark "Phase 10 — Bulk write API ✅ Done". Stale `.pr-body-phase10-slice*.md` drafts were never merged. The status claim is inaccurate and must be corrected by this change.

## Scope

### In Scope
- Rust `Worksheet`: `append`, `insert_rows`, `delete_rows`, `insert_cols`, `delete_cols`, `move_range`.
- HashMap-rebuild **shift engine** reusing `coords.rs`.
- Shift cells (value + `style_id` + comment + hyperlink — all ride on `Cell`) **plus** `merged_cells` + `row_heights`/`column_widths`.
- Formula cell values move **verbatim — NO reference translation** (exact openpyxl parity).
- `move_range(translate=False)` (move block, clear source).
- **Table-safety guard**: refuse any shift/move whose band intersects a `table.ref_range` (no silent corruption).
- PyO3 bindings with openpyxl signatures incl. dict/list `append`.
- Fix `SPEC.md` + `README.md` Phase 10 status.

### Out of Scope (non-goals)
- `move_range(translate=True)` formula translation → `NotImplemented` error.
- Formula reference rewriting on insert/delete — openpyxl does NOT do this; parity = leave verbatim.
- Auto-shifting `print_area`/`print_title_*`/`defined_names` ranges (openpyxl leaves stale; documented limitation).
- Charts/drawings/data-validation dependency management (openpyxl doesn't either).

## Capabilities

### New Capabilities
- `worksheet-bulk-write`: append, insert/delete rows/cols, move_range, shift semantics, table guard.

### Modified Capabilities
- None. Table interaction is specified inside `worksheet-bulk-write`; the SPEC/README fix is documentation, not a spec change.

## openpyxl parity (verified — openpyxl 3.1.2 in `.venv`)

| openpyxl op | Verified 3.1.2 behavior | oxlsx decision |
|---|---|---|
| `append(list/tuple/gen)` | positional from col A at `max_row+1` | match |
| `append(dict)` | keys = column LETTERS or 1-based int; mixed ok | match (Python layer maps keys) |
| `delete_rows/cols(idx, amount=1)` | shift cell values+styles up/left; formulas NOT translated; **merges & dims NOT shifted** | shift cells (parity) **+ also shift merges & dims** (safety superset); formulas verbatim |
| `insert_rows/cols(idx, amount=1)` | shift down/right; merges & dims NOT shifted | same as above |
| `move_range(rng, rows, cols, translate=False)` | move block, clear source, formulas verbatim | match |
| `move_range(translate=True)` | translates formulas | **DEFER → `NotImplemented`** |

> **Correction to brief:** the premise "openpyxl shifts row dimensions on insert/delete" is FALSE for 3.1.2 — verified: dims and merged ranges stay put. We intentionally EXCEED parity by shifting dims+merges (a non-surprising superset) to avoid data corruption.

## Approach

**Shift engine (HashMap rebuild):** single O(n) pass over `cells: HashMap<String, Cell>`. Parse each ref via `coords::column_index` + row digits; if in/after the affected band compute new `(row,col)`, rebuild the key via `coords::rc_to_a1`, write `cell.reference`, insert into a fresh map. Cells in a deleted band are dropped; inserts leave a gap. Comments/hyperlinks/`style_id` ride on `Cell` (no extra handling). Bounds (`max_row` etc.) recompute from the scan.

- `merged_cells`: parse `A1:B2`, shift endpoints; drop ranges that only partially straddle a deleted band (safest, no corruption).
- `row_heights`/`column_widths`: rekey the index→value maps.
- `tables`: if the band intersects any `ref_range` → return `OxlsxError` (proper ref-shift + dirty-marking deferred).
- `append` (Rust core): positional `append(impl IntoIterator<Item = Value>)` at `max_row+1`; Python layer maps `dict` (letters via `coords`, ints as 1-based) → index map.
- `move_range`: shift a sub-block by `(rows,cols)`, clear source cells not overwritten, formulas verbatim; `translate=True` → error.

**RMW interaction (crux):** a table whose cells move = the table is mutated → it must NOT take the verbatim-retention path. The intersection guard refuses the op rather than silently shipping a stale verbatim `tableN.xml`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/reader.rs` | New | 6 `Worksheet` methods + private shift engine |
| `src/coords.rs` | Reused | range parse + `rc_to_a1`/`column_name` (no API change) |
| `src/python.rs` | New | PyO3 bindings + dict/list `append` |
| `src/error.rs` | Modified | error variants: table-corruption, `translate=True` unimplemented |
| `SPEC.md`, `README.md` | Modified | Phase 10 status ✅→🔲/in-progress; uncheck checklist |
| `tests/` | New | TDD per scenario + openpyxl cross-engine round-trip |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Table ref corruption on shift | High | Intersection guard → error; proper shift deferred |
| Stale merges/dims (if matched openpyxl exactly) | Med | We shift them (superset) |
| Users expect formula translation | Med | Documented parity; `translate=True` deferred explicitly |
| HashMap rebuild perf on huge sheets | Low | Single O(n) pass, no DOM, SAX-friendly |
| Python dict-key ambiguity | Low | Verified semantics: letters + 1-based ints |

## Rollback Plan

Purely additive (new methods + doc edits); no file-format, schema, or `Arc<StyleSheet>` change. Revert the slice PR(s); existing read/write paths untouched.

## Dependencies

- `coords.rs` helpers (reuse), existing `Cell`/`Worksheet` model.
- `.venv` openpyxl 3.1.2 for cross-engine round-trip tests.

## Success Criteria

- [ ] All 6 methods exist on `Worksheet` + PyO3.
- [ ] `append` list & dict (letters + ints, mixed) match openpyxl cell placement.
- [ ] insert/delete shift cells + styles + comments + hyperlinks; formulas verbatim.
- [ ] `merged_cells` & dimensions shift; partial-straddle merges dropped (no corruption).
- [ ] shift/move intersecting a table → `OxlsxError`, never corruption.
- [ ] `move_range(translate=False)` moves+clears; `translate=True` → `NotImplemented` error.
- [ ] Cross-engine round-trip: oxlsx op vs openpyxl op on a fixture → equivalent grid when read back via openpyxl.
- [ ] `SPEC.md` & `README.md` Phase 10 status corrected.
- [ ] `cargo clippy -- -D warnings` + `cargo fmt --check` clean; strict TDD (test-first per scenario).

## Sizing & Delivery

Estimated >400 changed lines across Rust core + PyO3 + tests. **Chained PRs recommended.**

- **Slice 1**: shift engine + insert/delete rows/cols + merges/dims shift + table guard + Rust tests.
- **Slice 2**: `append` (Rust core) + `move_range(translate=False)` + Rust tests + SPEC/README fix.
- **Slice 3**: PyO3 bindings (all 6) + dict-append semantics + openpyxl cross-engine round-trip tests.

`Chained PRs recommended: Yes` · `400-line budget risk: High`
