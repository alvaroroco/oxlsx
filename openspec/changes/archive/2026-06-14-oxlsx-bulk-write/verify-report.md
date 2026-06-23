# Verification Report: oxlsx-bulk-write

**Change**: `oxlsx-bulk-write` (capability `worksheet-bulk-write`, the real Phase 10)
**Mode**: hybrid · Strict TDD ACTIVE · toolchain rustc 1.96.0 · pyo3 0.22 · openpyxl 3.1.2
**Branch**: `feat/bulk-write` (3 slice commits: f3259d8, fd9ba84, f4c297a)
**Verdict**: **PASS WITH WARNINGS**

## Completeness

| Artifact | State |
|----------|-------|
| proposal / spec / design / tasks | Present (Engram + OpenSpec) |
| apply-progress | Present (#344, S1+S2+S3 complete) |
| Tasks complete | All 3 slices committed; working tree clean (only untracked SDD docs) |

## Command Evidence (EXACT)

| Command | Result |
|---------|--------|
| `cargo test` | lib **118 passed/0 failed**; integration **93 passed/0 failed**; doctests 0 |
| `cargo test --features python` | lib **128 passed/0 failed**; integration **93 passed/0 failed** |
| `cargo clippy --all-targets --features python -- -D warnings` | **clean (exit 0)** |
| `cargo fmt --check` | **clean (exit 0)** |
| `maturin develop --features python` | built CPython 3.13 wheel, installed editable |
| `pytest tests/test_python.py -q` | **133 passed/0 failed** (130 baseline + 3 new) |

## Spec Compliance Matrix (14 → test → result)

| # | Scenario | Covering test | Result |
|---|----------|---------------|--------|
| 1 | append_adds_row_at_bottom | `reader::tests::append_adds_row_at_bottom` | ✅ PASS |
| 2 | insert_rows_shifts_cells_down | `reader::tests::insert_rows_shifts_cells_down` | ✅ PASS |
| 3 | delete_rows_shifts_cells_up | `reader::tests::delete_rows_shifts_cells_up` | ✅ PASS |
| 4 | insert_cols_shifts_cells_right | `reader::tests::insert_cols_shifts_cells_right` | ✅ PASS |
| 5 | delete_cols_shifts_cells_left | `reader::tests::delete_cols_shifts_cells_left` | ✅ PASS |
| 6 | move_range_relocates_block | `reader::tests::move_range_relocates_block` | ✅ PASS |
| 7 | delete_rows_shifts_merged_cells_and_dimensions | `reader::tests::delete_rows_shifts_merged_cells_and_dimensions` | ✅ PASS |
| 8 | bulk_op_intersecting_table_returns_err | `reader::tests::bulk_op_intersecting_table_returns_err` | ✅ PASS |
| 9 | move_range_translate_true_is_unsupported | `reader::tests::move_range_translate_true_is_unsupported` | ✅ PASS |
| 10 | delete_rows_does_not_translate_formulas | `reader::tests::delete_rows_does_not_translate_formulas` | ✅ PASS |
| 11 | test_append_list_and_dict | `tests/test_python.py::test_append_list_and_dict` | ✅ PASS |
| 12 | test_delete_and_insert_rows | `tests/test_python.py::test_delete_and_insert_rows` | ✅ PASS |
| 13 | test_move_range | `tests/test_python.py::test_move_range` | ✅ PASS |
| 14 | phase10_docs_reflect_reality | SPEC.md L103 🔲 + README L51/224 unchecked (file inspection) | ⚠️ PASS (stale sub-text — see W1) |

**Bonus coverage** (beyond mandate): `move_range_intersecting_table_returns_err`, `move_range_out_of_bounds_returns_err` — both ✅ PASS.

## Critical-Check Findings

| Check | Result | Evidence |
|-------|--------|----------|
| Shift engine: rebuild, no key collisions | ✅ | `shift_axis` uses `mem::take` + fresh `HashMap` rebuild (reader.rs:576); unparseable keys kept verbatim |
| Formulas verbatim (no translation) | ✅ | S10 asserts `B2 == Formula("=A3+1")` after shift; move keeps text |
| Merges + dims shift; partial-straddle dropped | ✅ | S7 triangulates: before unchanged (A1:B2), straddle dropped (A2:B4 gone, len==2), after shifted (A5:B6→A4:B5), row_heights 5→4, inside-band dropped |
| Table guard: Err + state UNCHANGED | ✅ | Guard runs BEFORE mutation (insert/delete call `guard_tables` first; move_range guards src+dest before any remove). S8 asserts `cells().len()==before` AND `A1=="H1"` post-Err |
| move_range: translate=True→Err, OOB→Err, dest overwrite | ✅ | `UnsupportedTranslate`; OOB `InvalidFormat` (row/col<1); dest "victim"→"a1" verified |
| append dict: letter + 1-based int + mixed, empty-padded | ✅ | `append_column_from_key` (String before i64; <1→ValueError); padded `vec![Value::Empty; max_col]` |
| Byte-equality / verbatim-RMW discriminators intact | ✅ | `table_unmodified_rmw_preserves_raw_table_xml`, `tables_rmw_real_sample_no_duplicate_overrides`, `tables_rmw_strips_stale_content_type_overrides`, `table_mutated_rmw_serializes_from_model`, `tables_readonly_access_keeps_table_verbatim`, hyperlink/comment RMW — ALL ok. `tables_totals.xlsx` exercised (reader.rs + integration_test.rs:2198-2425) |
| Conventions: no `get_`, no public lifetimes, coords reused, Arc<StyleSheet> untouched | ✅ | No `pub fn get_` matches; no `pub struct Worksheet<`/`Cell<`; shift engine uses `coords::{parse_a1,parse_range,rc_to_a1,bands_intersect}`; cells carry their Arc clone unmodified |

## No-Test-Weakening Audit

| Aspect | Result |
|--------|--------|
| Cross-engine pytest truly compares vs openpyxl | ✅ All 3 tests double-assert `grid_x == expected` **AND** `grid_x == grid_o` (openpyxl-derived via `_window`/`iter_rows`); `_do_appends` covers list+letter-dict+int-dict+mixed |
| Guard tests assert state-unchanged | ✅ S8 + move_range guard tests assert cell count + values post-Err |
| Persistence parity | ✅ S12 adds save+reopen round-trip assertion (`grid_x_disk == expected`) |
| Trivial/tautology/empty/smoke/ghost-loop assertions | ✅ None found |

**Assertion quality**: ✅ All assertions verify real behavior.

## TDD Compliance

| Check | Result | Details |
|-------|--------|---------|
| TDD Evidence reported | ✅ | apply-progress #344 has TDD Cycle Evidence table |
| All scenarios have tests | ✅ | 14/14 covered |
| RED confirmed (tests exist) | ✅ | All 13 test fns verified present in source |
| GREEN confirmed (tests pass) | ✅ | 13/13 pass at runtime |
| Triangulation adequate | ✅ | append 4 cases; merge/dim multi-assertion; move src+dest+OOB |
| Safety net | ✅ | 118 lib + 93 int baseline green pre/post |

**TDD note (W2)**: the formal TDD Cycle Evidence *table* in apply-progress explicitly covers Slice 3; Slices 1–2 are documented as RED→GREEN narrative + per-slice commits (adequate, but not in table form).

## Test Layer Distribution

| Layer | Tests | Tool |
|-------|-------|------|
| Unit (Rust) | 10 bulk-write `#[test]` + 2 bonus + coords/error units | cargo test |
| Integration (Python cross-engine) | 3 (scenarios 11,12,13) | pytest + openpyxl 3.1.2 |

**Coverage**: skipped — no coverage tool (tarpaulin/llvm-cov) configured. Not a failure.
**Quality Metrics**: Linter (clippy) ✅ clean; Formatter ✅ clean.

## Issues

**CRITICAL**: None.

**WARNING**:
- **W1 — Doc staleness (the very risk this change targeted)**: SPEC.md L103 / README L51,L224 say "Rust core done; Python bindings pending". Slice 3 COMPLETED the Python bindings (scenarios 11–13 green), so "pending" is now inaccurate. Phase 10 is fully implemented and should read ✅ Done. The literal spec scenario (🔲/in-progress, unchecked) passes, but truthfulness now requires promotion to Done.
- **W2 — TDD evidence table partial**: formal table covers Slice 3 only; Slices 1–2 narrative-only (backed by commits + green tests).

**SUGGESTION**:
- **S1**: `src/writer.rs::parse_a1_coord` remains a separate helper duplicating `coords::parse_a1`. Pre-existing, out of this change's scope; candidate for future consolidation. New bulk-write code correctly reuses `coords`.

## Recommendation

**Archive-ready** after a one-line doc correction. The implementation is complete, all 14 scenarios pass at runtime, conventions hold, and the verbatim-RMW/byte-equality discriminators are intact (no regression). Before/at archive, resolve **W1**: update SPEC.md L103 + README L51/L224 to mark Phase 10 ✅ Done (Rust + Python complete) — closing the doc-accuracy loop this change set out to fix.
