# Proposal: Worksheet Conditional Formatting (openpyxl parity)

## Intent

Real openpyxl→oxlsx migrations break on `from openpyxl.formatting.rule import FormulaRule`
(`excel/format.py`: `ws.conditional_formatting.add("A1:A10", FormulaRule(...))`). oxlsx has
ZERO conditional formatting (CF) today. CF is the most common openpyxl styling feature after
fonts/fills; without it, styled-report migrations stall. Close the gap with `FormulaRule` +
`CellIsRule` and a differential-format (`dxf`) registry serialized into `xl/styles.xml`.

## Scope

### In Scope
- `ws.conditional_formatting.add(sqref, rule)` (Python) + `ws.add_conditional_formatting(sqref, rule)` (Rust).
- `FormulaRule` → `cfRule type="expression"`; `CellIsRule` → `cfRule type="cellIs" operator=…`.
- `Dxf` differential format with **fill** (PatternFill) and **font** (color/bold/italic) — reusing existing `Fill`/`Font`/`Color`.
- Emit `<dxfs>` in `styles.xml` (after `<cellStyles>`) with sequential `dxfId` assignment; emit `<conditionalFormatting>` in sheet XML (after `mergeCells`, before `hyperlinks`); auto-assign `priority`.
- RMW: parse + preserve existing `<dxfs>` and existing `<conditionalFormatting>`, offsetting new `dxfId`s past preserved ones.
- `stopIfTrue` passthrough. openpyxl cross-engine round-trip (no-repair).

### Out of Scope
- `ColorScaleRule`, `DataBarRule`, `IconSetRule` (color scales / data bars / icon sets).
- `dxf` borders & number formats (font + fill only this change).
- Charts. Multi-`cfRule` ergonomics beyond add-order priority. `Index`-style sugar.

## Capabilities

### New Capabilities
- `worksheet-conditional-formatting`: model + read/write of `<conditionalFormatting>` rules (FormulaRule, CellIsRule), the `Dxf` differential-format registry, and its `<dxfs>` serialization in `styles.xml`.

### Modified Capabilities
- None (no existing styles capability spec; `<dxfs>` arrives with the new capability).

## Approach

1. **Dxf/styles infra** (foundation): add `Dxf { fill: Option<Fill>, font: Option<Font> }`; parse `<dxfs>` in `StyleSheet`; emit `<dxfs>` in `write_styles_xml` after `cellStyles`; preserve source dxfs in `collect_styles` so the writer no longer silently drops them.
2. **Rule model + emission**: `ConditionalFormatting` collection on `Worksheet` (`sqref → Vec<Rule>`); `Rule`/`FormulaRule`/`CellIsRule`; assign each rule's `dxf` a `dxfId` (offset past preserved dxfs) and a sequential `priority`; emit `<conditionalFormatting>` in correct CT_Worksheet slot. Read existing CF on RMW.
3. **Python parity**: expose `FormulaRule`, `CellIsRule` (top-level `oxlsx.*`, mirroring existing flat `oxlsx.PatternFill`/`oxlsx.Font`; optional `oxlsx.formatting.rule` shim) + `ws.conditional_formatting` proxy with `.add(sqref, rule)`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/styles.rs` | Modified | `Dxf` type; parse `<dxfs>`; `dxfs` field on `StyleSheet` |
| `src/conditional_formatting.rs` | New | `Rule`/`FormulaRule`/`CellIsRule`/`ConditionalFormatting` model |
| `src/reader.rs` | Modified | `Worksheet` CF collection; read existing `<conditionalFormatting>` for RMW |
| `src/writer.rs` | Modified | Emit `<dxfs>`; preserve source dxfs; emit `<conditionalFormatting>` in slot |
| `src/python.rs` | Modified | `FormulaRule`/`CellIsRule` pyclasses + `conditional_formatting` proxy |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| `<dxfs>` integration breaks existing CellXf/font/fill serialization or byte-equality discriminators | High | dxfs is an additive trailing block; gate emission on non-empty; keep all existing fixture byte-equality tests green; no-dxf files emit zero `<dxfs>` |
| RMW dangling/colliding `dxfId` (new rules vs preserved dxfs) | High | Parse source dxfs first; assign new dxfIds at `source_dxf_count + i`; preserve source CF verbatim-model |
| Color alpha mismatch (openpyxl writes `00`-prefixed `rgb`) | Med | Match openpyxl's `<color rgb="00…"/>` form in cross-engine assertions |
| `priority` collisions across multiple ranges | Med | Auto-assign globally increasing priority in add order |

## Rollback Plan

Feature is purely additive: revert the slice branch(es). With zero CF rules + zero dxfs, `styles.xml` and sheet XML emit exactly as today (no `<dxfs>`/`<conditionalFormatting>`), so existing fixtures stay byte-identical. No persisted data migration.

## Dependencies

- Existing `Fill`/`Font`/`Color` (styles.rs) and A1 coords (coords.rs) — reuse, no new crates.
- `Arc<StyleSheet>` ownership untouched; `dxfs` lives inside `StyleSheet` (shared) — no new lifetimes, PyO3-safe.

## Sizing & Delivery

- Estimate ~500–650 changed lines (styles/dxf infra ~150, rule model+emission ~220, Python ~120, tests ~150).
- **400-line budget risk: High → Chained PRs recommended: Yes.** Three slices: (1) dxf/styles infra + RMW preservation; (2) rule model + `<conditionalFormatting>` emission + dxfId assignment + Rust API + RMW CF read; (3) Python bindings.

## Success Criteria

- [ ] `ws.add_conditional_formatting("A1:A10", FormulaRule{...})` emits `<conditionalFormatting sqref="A1:A10"><cfRule type="expression" priority="1" dxfId="0">…</cfRule>` + matching `<dxf>` in `<dxfs>`.
- [ ] `CellIsRule` emits `type="cellIs" operator="greaterThan"` with `<formula>`.
- [ ] **Cross-engine**: write a FormulaRule with oxlsx, reopen in openpyxl 3.1.2 — rule + dxf present, file opens with NO repair prompt.
- [ ] RMW of a file with existing CF/dxfs preserves them and appends new ones without dxfId collision.
- [ ] All existing byte-equality / no-repair fixture tests stay green (test.xlsx, multi_sheet, tables*, inline_strings).
- [ ] `cargo test`, `cargo clippy -- -D warnings`, `cargo fmt --check`, `pytest` all pass (strict TDD).
