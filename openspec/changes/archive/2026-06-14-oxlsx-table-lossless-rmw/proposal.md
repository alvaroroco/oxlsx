# Proposal: Lossless RMW preservation of unmodeled worksheet-table features

## Intent / Problem & User Value
Phase 13 models a thin slice of a table (`id, name, display_name, ref, headerRowCount, columns, tableStyleInfo`) and **re-serializes every table from that model on save** (`writer.rs:362-372`, `serialize_table_xml`). Table parts are writer-owned (`is_writer_owned_entry` → `xl/tables/` always re-emitted), and `parse_table_xml` ignores everything else. So a file authored in Excel/openpyxl that uses `totalsRowCount`/`totalsRowShown`, per-column `totalsRowLabel`/`totalsRowFunction`/`calculatedColumnFormula`, `dxf` overrides, `autoFilter` criteria, or `sortState` **loses all of them** the moment oxlsx opens it, edits an unrelated cell, and saves. Confirmed against openpyxl 3.1.2: openpyxl round-trips these byte-for-byte (it models them); oxlsx silently drops them. This is genuine data loss in the most common workflow (open → edit one cell → save).

## Scope
### In Scope
- Retain each table's **original raw `tableN.xml` bytes** at read time.
- On save, re-emit the **verbatim original bytes** when the table was NOT mutated through the API (the common case → fully lossless).
- Serialize from the model only when the table WAS mutated, with documented loss of unmodeled features.
- Reconcile id/rId/part-number allocation with retained verbatim ids.
- One real-use regression test (totals-row file → edit unrelated cell → reopen with openpyxl → features survived).

### Out of Scope / Non-Goals
- NO first-class API for `totalsRow*`, `sortState`, `autoFilter` criteria, `dxf`, or `calculatedColumnFormula` (not modeled; preserved opaquely).
- NO extras-bag parsing of unmodeled attributes/children (rejected, see Approach).
- Charts, pivot tables — OUT.

## Capabilities
### New Capabilities
- None — extends behavior of an existing capability.
### Modified Capabilities
- `worksheet-tables`: add a requirement that unmodeled table features survive read-modify-write when the table is not mutated.

## Approach — Recommended: (A) Verbatim-when-unmodified
Mirror the established conservative-passthrough pattern (`defined_names.rs` raw `extras`, `copy_unknown_entries`). At parse time, retain the table's raw bytes alongside the parsed model. At save time, per table: **unmutated → write retained bytes verbatim; mutated → `serialize_table_xml` from the model.**

**Why A over B (extras bag):** the dominant RMW case touches a *cell*, never the table — A makes that case 100% lossless with ~minimal code. B requires parsing every unmodeled attribute/child of `tableN.xml`, storing it opaquely, AND re-emitting it in correct OOXML schema order interleaved with modeled fields — high complexity, schema-ordering corruption risk, and still only *partial* fidelity. A is strictly simpler and strictly more correct for the common case. C (hybrid) deferred: only needed if we later want partial-mutation fidelity, which no current user need justifies.

### Behavior contract
| Case | Behavior | Loss |
|------|----------|------|
| (i) table untouched | re-emit retained raw bytes verbatim | none |
| (ii) style/ref/columns mutated via API | serialize from model; discard raw bytes for that table | unmodeled features (documented) |
| (iii) new table via `add_table` | serialize from model (no raw bytes exist) | n/a — nothing to lose |
| (iv) table removed | drop raw bytes + part + rel + `<tablePart>` | n/a |

## API Impact (minimal / ideally none)
- `ws.tables()`, `add_table()` signatures unchanged. `add_table` has no raw bytes → model path (today's behavior).
- `table_mut()` is the single mutation entry point → returns `&mut` → conservatively marks the table **dirty** (mutated).
- Retention state (`raw_xml`, `dirty`) must NOT become public struct-literal fields on `Table` (that breaks external `Table { .. }` construction). Design decides: store as a **parallel structure on `Worksheet`** aligned to `tables`, or `#[doc(hidden)]` private fields. Prefer the parallel-structure option to keep `Table` a clean public data model.

## Implementation Sketch
- `reader.rs:648-653`: when reading a table part, also capture the raw bytes (read file fully) and store them with `dirty=false`.
- `table.rs` / `reader.rs`: retention store + dirty flag (parallel `Vec<Option<RawTable>>` on `Worksheet`, or private fields). `table_mut` sets dirty.
- `writer.rs:358-372`: in the re-emit loop, if `raw_bytes.is_some() && !dirty` → write the original bytes to `xl/tables/tableN.xml` **without rewriting `id`/`displayName`**; else `serialize_table_xml`.
- `<tableParts>` ordering, content-type override, and rId allocation (`sheet_rels::upsert_table_rels`) unchanged — linkage is by part filename + rId, independent of the internal `<table id>`.
- id allocator (`assign_table_numbers`): the part-number → filename mapping is unchanged; but **API-created** table internal ids must avoid colliding with retained verbatim ids (scan retained ids first).

## Risks / Tradeoffs
| Risk | Likelihood | Mitigation |
|------|-----------|------------|
| **id renumbering vs verbatim bytes** — writer assigns `id=N`==part number; verbatim bytes keep their original internal `id`/`displayName` | High | Decouple: linkage is filename+rId, not internal id. Do NOT rewrite ids inside verbatim bytes. Allocate ids for *model-path* tables avoiding retained verbatim ids → workbook-wide uniqueness preserved |
| stale-vs-fresh: ref/columns derived from cells diverge from retained bytes | Med | Any API mutation flips to model path (fresh, consistent); unmutated stays verbatim (consistent with disk) |
| `table_mut` returns `&mut` even for read-only peeks → false dirty | Low | Acceptable conservative loss; document. Optional fine-grained dirty later |
| Public `Table` struct-literal construction breaks if retention added as fields | Med | Keep retention off the public struct (parallel store) |
| New file (no source ZIP) has no raw bytes | n/a | Model path only — unchanged |

## Rollback Plan
Additive and gated on `raw_bytes.is_some() && !dirty`. Revert the branch; with retention disabled the writer falls back to today's model-only re-emission. Files without tables are byte-unaffected.

## Dependencies
- openpyxl 3.1.2 in `.venv` for the regression fixture (build `samples/tables_totals.xlsx`).

## Success / Acceptance Criteria
- [ ] **Mandatory normal-use regression**: open an openpyxl file with a totals-row table (+ `totalsRowFunction`/`calculatedColumnFormula`), edit an unrelated cell, save, reopen with openpyxl → totals row and unmodeled features SURVIVED (byte/struct assert on `tableN.xml`).
- [ ] Editing the table via `table_mut` (ref/columns/style) → model-path output opens in Excel/openpyxl without repair (documented feature loss accepted).
- [ ] `add_table` on a new workbook unchanged; output opens without repair.
- [ ] Removing a table drops its part, rel, and `<tablePart>`.
- [ ] Workbook-wide id uniqueness holds when verbatim + API tables coexist.
- [ ] `cargo test` + `clippy -D warnings` + `fmt --check` + `pytest` green; no `get_` prefixes, no public lifetimes, `Arc<StyleSheet>` untouched.

## Sizing / Chain Recommendation
Estimated ~150-250 LOC core. **Single focused PR** preferred (under 400 lines). If review slicing is wanted, 2 chained units on `feature/table-lossless-rmw`:
- **Unit 1** — retention + verbatim re-emit (reader capture, writer verbatim path, id reconciliation, regression test). ~150 LOC.
- **Unit 2** — dirty tracking on `table_mut` + mutation→model-path contract + tests. ~100 LOC.
