# worksheet-conditional-formatting Specification

## Purpose

Add openpyxl-parity conditional formatting (CF) to worksheets: `FormulaRule` and `CellIsRule`
with a differential-format (`Dxf`) registry serialized as `<dxfs>` in `xl/styles.xml` and
`<conditionalFormatting>` in the sheet XML. Unblocks `from openpyxl.formatting.rule import FormulaRule`.

## In Scope

- `FormulaRule` (`cfRule type="expression"`), `CellIsRule` (`type="cellIs"` + `operator`).
- `Dxf { fill: Option<Fill>, font: Option<Font> }` reusing existing `Fill`/`Font`/`Color`/`PatternFill`.
- `stopIfTrue`; auto-assigned global `priority`; dxfId assignment; RMW preservation.
- Rust `ws.add_conditional_formatting(sqref, rule)`; Python `ws.conditional_formatting.add(sqref, rule)`, `FormulaRule`, `CellIsRule`.

## Out of Scope (deferred)

- `ColorScaleRule`, `DataBarRule`, `IconSetRule`.
- `Dxf` borders and number formats.
- Charts.

## Requirements

### Requirement: Register conditional-formatting rules on a worksheet

The system MUST let callers attach a rule to an A1 range (`sqref`) via `ws.add_conditional_formatting(sqref, rule)` (Rust) and `ws.conditional_formatting.add(sqref, rule)` (Python), preserving add order.

#### Scenario: Multiple rules get distinct priorities and stopIfTrue (`multiple_rules_priorities_and_stopiftrue`)

- GIVEN three rules added across different ranges, one with `stopIfTrue=true`
- WHEN the workbook is saved
- THEN each `cfRule` carries a unique, sequential `priority` ("1","2","3") in add order
- AND the `stopIfTrue` rule serializes `stopIfTrue="1"` while others omit the attribute

### Requirement: FormulaRule serialization

A `FormulaRule` MUST serialize as `<cfRule type="expression" priority=N dxfId=M>` with each formula as an XML-escaped `<formula>` child, and its differential format MUST appear in `<dxfs>`.

#### Scenario: FormulaRule writes cfRule and dxf (`formula_rule_write_and_dxf_emitted`)

- GIVEN `FormulaRule(formula=["$A1>5"])` with a solid fill added to `A1:A10`
- WHEN saved
- THEN sheet XML contains `<conditionalFormatting sqref="A1:A10">` with a `cfRule` having `type="expression"`, `dxfId="0"`, `priority="1"`, and child `<formula>$A1&gt;5</formula>`
- AND `xl/styles.xml` contains `<dxfs count="1"><dxf><fill><patternFill ...>` with the fill

### Requirement: CellIsRule serialization

A `CellIsRule` MUST serialize as `<cfRule type="cellIs" operator=... priority=N dxfId=M>` with `<formula>` operand children.

#### Scenario: CellIsRule writes operator and formula (`cell_is_rule_write`)

- GIVEN `CellIsRule(operator="greaterThan", formula=["5"])` added to a range
- WHEN saved
- THEN the `cfRule` has `type="cellIs"`, `operator="greaterThan"`, and child `<formula>5</formula>`

### Requirement: Dxf registry and gated `<dxfs>` emission

Differential formats MUST live inside `StyleSheet`. The writer MUST emit `<dxfs>` after `<cellStyles>`. Identical dxfs SHOULD be deduplicated (find-or-push, matching openpyxl). When NO rule exists, NO `<dxfs>` MUST be emitted and `styles.xml` MUST stay byte-identical to current output.

#### Scenario: No CF keeps styles byte-identical (`no_cf_keeps_styles_byte_identical`)

- GIVEN a workbook with no conditional formatting (incl. byte-equality fixtures: `tables_totals.xlsx`, `inline_strings_no_sst.xlsx`)
- WHEN saved
- THEN `styles.xml` contains no `<dxfs>` element and existing byte-equality / no-repair tests still pass

### Requirement: Conditional-formatting round-trips through oxlsx

Saved CF MUST be re-readable by oxlsx, recovering rule type, formulas, operator, priority, and the linked dxf.

#### Scenario: Round-trip read (`conditional_formatting_round_trip`)

- GIVEN a workbook saved with a FormulaRule and a fill dxf
- WHEN reopened with oxlsx
- THEN the rule (type, formula, priority) and its dxf fill are read back

### Requirement: RMW preserves existing dxfs and conditional formatting

On read-modify-write the system MUST preserve source `<dxfs>` and source `<conditionalFormatting>`. New dxfs MUST be appended after preserved ones (new dxfId ≥ `source_dxf_count`) so no dxfId collides, and existing styles MUST NOT be corrupted.

#### Scenario: RMW preserves source dxfs and rules (`cf_rmw_preserves_source_dxfs_and_rules`)

- GIVEN a fixture that already has conditional formatting and dxfs
- WHEN a new rule is added and the workbook is saved
- THEN source dxfs and rules are preserved AND the new rule's dxfId is offset past the source dxfs with no collision

### Requirement: Python parity for FormulaRule and CellIsRule

`FormulaRule` and `CellIsRule` MUST be importable from oxlsx and usable via `ws.conditional_formatting.add(...)`; output MUST open in openpyxl 3.1.x with no repair.

#### Scenario: FormulaRule round-trips to openpyxl (`test_formula_rule_round_trip`)

- GIVEN `ws.conditional_formatting.add("A1:A10", FormulaRule(formula=["$A1>5"], fill=PatternFill(...)))`
- WHEN saved and reopened with openpyxl
- THEN the rule is present and the file opens with no repair prompt

#### Scenario: CellIsRule read by openpyxl (`test_cell_is_rule`)

- GIVEN a `CellIsRule(operator="greaterThan", formula=["5"])` added via Python
- WHEN reopened with openpyxl
- THEN openpyxl reads `operator="greaterThan"` and the formula

#### Scenario: PatternFill survives in the dxf (`test_conditional_formatting_with_pattern_fill`)

- GIVEN a rule whose dxf carries a `PatternFill`
- WHEN saved and reopened with openpyxl
- THEN the dxf fill (fgColor) is present on the rule's differential format

## Notes

- **dxfs-in-StyleSheet**: `Dxf` is stored inside `StyleSheet` (shared via `Arc<StyleSheet>`), no new lifetimes, PyO3 `'static`-safe.
- **Two XML slots**: `<dxfs>` in `styles.xml` immediately after `<cellStyles>` (CT_Stylesheet); `<conditionalFormatting>` in sheet XML between `<mergeCells>` and `<hyperlinks>` (CT_Worksheet: after mergeCells/phoneticPr, before dataValidations/hyperlinks).
- **dxfId/priority**: priority is global, sequential, unique in add order. dxfId via find-or-push dedup — CORRECTION to proposal's "strictly sequential": openpyxl 3.1.2 dedupes identical dxfs (verified: 3 rules sharing a fill → `count="2"`). On RMW, new dxfs append after preserved source dxfs.
- **Serialization (openpyxl-verified)**: cfRule attribute order `type, priority, [operator], dxfId, [stopIfTrue]`; `stopIfTrue="1"` (not "true"); formulas XML-escaped (`&gt;`,`&lt;`); dxf fill emits both `<fgColor>`+`<bgColor>`; colors `00`-prefixed (`rgb="009C0006"`). Scenarios assert attribute VALUES (presence), not brittle ordered substrings.
- **Import-path parity**: openpyxl exposes `FormulaRule`/`CellIsRule` from `openpyxl.formatting.rule`. oxlsx exposes them top-level (`oxlsx.FormulaRule`, `oxlsx.CellIsRule`), mirroring the existing flat `oxlsx.PatternFill`/`oxlsx.Font` precedent; an `oxlsx.formatting.rule` shim MAY be added later.

## Test Mapping (1:1)

| # | Scenario | Test | Kind |
|---|----------|------|------|
| 1 | FormulaRule writes cfRule + dxf | `formula_rule_write_and_dxf_emitted` | cargo |
| 2 | CellIsRule operator+formula | `cell_is_rule_write` | cargo |
| 3 | Round-trip read | `conditional_formatting_round_trip` | cargo |
| 4 | No CF byte-identical styles | `no_cf_keeps_styles_byte_identical` | cargo |
| 5 | RMW preserves dxfs+rules | `cf_rmw_preserves_source_dxfs_and_rules` | cargo |
| 6 | Priorities + stopIfTrue | `multiple_rules_priorities_and_stopiftrue` | cargo |
| 7 | FormulaRule round-trip | `test_formula_rule_round_trip` | pytest |
| 8 | CellIsRule via Python | `test_cell_is_rule` | pytest |
| 9 | PatternFill in dxf | `test_conditional_formatting_with_pattern_fill` | pytest |
