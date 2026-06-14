# Proposal: Worksheet Tables (`ws.add_table` / `ws.tables`) — Phase 13 final slice

## Intent / Problem & User Value
oxlsx cannot model Excel Tables (structured ranges). Two concrete gaps:
1. **No API** to read or create tables, the last openpyxl worksheet feature in Phase 13.
2. **Latent RMW corruption (confirmed by audit)**: `write_sheet_xml` never emits `<tableParts>` and every `sheetN.xml` is writer-owned (rewritten). A table-bearing file saved by oxlsx today keeps the orphaned `xl/tables/tableN.xml` + sheet rel (copied via `copy_unknown_entries`) but DROPS `<tableParts>` from the rewritten sheet → orphaned/repair-flagged table. This slice closes both.

Mirror openpyxl: `Table(displayName, ref)`, `ws.add_table(tab)`, `ws.tables["Table1"]` (TableList keyed by displayName).

## Scope
### In Scope (minimal Excel-VALID slice)
- Rust `Table` model: `id`, `name`, `display_name`, `ref`, `header_row_count`, `columns: Vec<TableColumn{id,name}>`, optional `style: Option<TableStyleInfo>`.
- `TableStyleInfo` = full openpyxl struct (`name` + 4 bools) — cheap, complete, high-value.
- READ: parse `xl/tables/tableN.xml` via the `.../table` sheet rel; expose `ws.tables()`.
- WRITE: `ws.add_table(table)`; auto-generate `tableColumns` from header row when columns omitted; emit `<autoFilter ref>`=ref; new `xl/tables/tableN.xml` part; CT override `…spreadsheetml.table+xml`; `<tableParts>` as LAST child of `<worksheet>`; sheet rel via reused `sheet_rels.rs`.
- Workbook-wide unique `id` and `displayName` allocation across all sheets/tables.
- RMW preservation: unknown parts untouched; existing typed tables re-emitted with `<tableParts>` restored.
- Python parity: `Table`, `TableStyleInfo`, `ws.add_table`, `ws.tables` (dict-like).

### Out of Scope / Non-Goals (deferred)
- dxf style overrides (`headerRowDxfId`, `dataDxfId`, …), totals row (`totalsRowCount`/functions), calculated-column formulas, autoFilter sort/filter CRITERIA (we emit only bare `<autoFilter ref>`), table style name overrides beyond TableStyleInfo, `sortState`.
- Charts and pivot tables — explicitly OUT.

## Capabilities
### New Capabilities
- `worksheet-tables`: read existing tables and create minimal Excel-valid tables on a worksheet.
### Modified Capabilities
- None at spec level (writer `<tableParts>` fix is implementation of the new capability).

## Approach
Reuse the Phase 12b comments new-part pattern verbatim. Add `src/table.rs` (model + parse + serialize, sibling to `comments.rs`). Add `sheet_rels::upsert_table_rels` mirroring `upsert_comment_rels` (preserve unrelated, allocate next `rIdN`, coexist with comment/hyperlink rels). Gate writer emission on `has_tables(ws)`; allocate workbook-wide table `id` in the sheet loop; pass table rIds into `write_sheet_xml` so `<tableParts>` is appended last. Read path: in `reader.rs`, for each `.../table` rel resolve+parse part into `ws.tables`.

## Affected Areas
| Area | Impact | Description |
|------|--------|-------------|
| `src/table.rs` | New | Table/TableColumn/TableStyleInfo model + XML parse/serialize |
| `src/sheet_rels.rs` | Modified | Add `TABLE_REL_TYPE` + `upsert_table_rels` |
| `src/reader.rs` | Modified | `Worksheet.tables` field + parse via table rel |
| `src/writer.rs` | Modified | `has_tables`, write table part, CT override, `<tableParts>` in sheet XML, id/rId alloc |
| `src/python.rs` | Modified | `PyTable`, `PyTableStyleInfo`, `add_table`, `tables` |

## Risks
| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Excel strictness: malformed table corrupts/repairs file | High | Byte-diff golden vs openpyxl 3.1.5 output; emit required attrs (id,name,displayName,ref,headerRowCount,autoFilter,tableColumn id+name) |
| Workbook-wide `id`/`displayName` collisions | Med | Central allocator scanning all sheets+existing parsed ids before assigning |
| Sheet-rel collision with comments/hyperlinks | Med | Reuse preserve-unrelated upsert; allocate next free rId |
| Header-row ↔ column-name inconsistency | Med | Derive column names from header cells at add_table time; validate count matches ref width |
| `<tableParts>` schema-order (must be last child) | Med | Append after pageMargins/all existing children in write_sheet_xml |

## Rollback Plan
Feature is additive and gated on `has_tables`. Revert the chain's branches; `sheet_rels.rs` table additions are unused without writer wiring. Pre-existing files without tables are byte-unaffected.

## Dependencies
- openpyxl 3.1.5 in `.venv` for parity golden fixtures. Build a `samples/tables.xlsx` fixture.

## Success Criteria
- [ ] Round-trip read: open a table-bearing xlsx → `ws.tables["Table1"].ref` correct; `<tableParts>` no longer dropped on save.
- [ ] `ws.add_table(Table(displayName,ref))` → output opens in Excel/openpyxl without repair; columns auto-derived from header row.
- [ ] Multiple tables across multiple sheets get unique ids/rIds.
- [ ] Python parity tests pass; no `get_` prefixes, no public lifetimes, `Arc<StyleSheet>` untouched.
- [ ] `cargo test` + `clippy -D warnings` + `fmt --check` + `pytest` green.

## Sizing / Chain Recommendation
Estimated >400 lines → **3-unit feature-branch-chain** (matches prior Phase 13 slices):
- **Unit 1 — Rust READ**: `table.rs` model+parse, `Worksheet.tables`, read path, `<tableParts>` preservation fix. (~250-350 LOC)
- **Unit 2 — Rust WRITE**: `add_table`, new-part emission, CT override, sheet rel, id/rId allocation. (~250-350 LOC)
- **Unit 3 — Python parity**: PyTable/PyTableStyleInfo/add_table/tables + pytest. (~200-250 LOC)
Tracker branch `feature/phase13-tables`.
