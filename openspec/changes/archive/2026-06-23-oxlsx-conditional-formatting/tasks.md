# Tasks: worksheet-conditional-formatting

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 550–650 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | Slice 1 (dxf infra ~150) → Slice 2 (rule model ~220) → Slice 3 (Python ~120) |
| Delivery strategy | auto-chain |
| Chain strategy | feature-branch-chain |

Decision needed before apply: No
Chained PRs recommended: Yes
Chain strategy: feature-branch-chain
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Base Branch | Scenarios |
|------|------|-----------|-------------|-----------|
| 1 | Dxf infra + RMW dxf preservation + byte guard | PR 1 → `feat/conditional-formatting` | `main` | Scenario 4 partial (byte guard) |
| 2 | Rule model + emission + Rust API + RMW CF | PR 2 → `feat/conditional-formatting` | PR 1 branch | Scenarios 1,2,3,5,6 |
| 3 | Python bindings + cross-engine tests | PR 3 → `feat/conditional-formatting` | PR 2 branch | Scenarios 7,8,9 |

**Rollback per slice**: Each PR is independently revertible; zero CF → byte-identical output.

---

## Slice 1: Dxf Infrastructure + RMW Preservation + Byte Guard (~150 lines)

**Focus**: `Dxf` type, `StyleSheet` parsing/emission, `collect_styles` preservation, byte-equality guard.
**Scenarios**: 4 (partial: byte guard that no CF → no `<dxfs>`).

### Phase 1.1: Define Dxf Type (src/styles.rs)

- [x] 1.1.1 **RED**: Write unit test `dxf_with_fill_serializes` that creates a `Dxf { fill: Some(Fill::solid_red()), font: None }`, serializes to XML string, asserts contains `<dxf><fill><patternFill patternType="solid"><fgColor rgb="00FF0000"/><bgColor rgb="00FF0000"/></patternFill></fill></dxf>`. Test fails (Dxf not defined).
- [x] 1.1.2 **GREEN**: Define `pub struct Dxf { pub fill: Option<Fill>, pub font: Option<Font> }` with `Debug, Clone, PartialEq`.
- [x] 1.1.3 **RED**: Write unit test `dxf_with_font_serializes` for a `Dxf` with bold red font color.
- [x] 1.1.4 **GREEN**: Implement stub serialization (write_dxf helper in writer.rs — wire later). Pass tests.

### Phase 1.2: Parse `<dxfs>` in StyleSheet (src/styles.rs)

- [x] 1.2.1 **RED**: Write unit test `stylesheet_parses_dxfs` with sample XML containing `<dxfs count="1"><dxf><fill>...</fill></dxf></dxfs>`. Assert `sheet.dxfs.len() == 1`.
- [x] 1.2.2 **GREEN**: Add `pub dxfs: Vec<Dxf>` field to `StyleSheet`. Parse `<dxfs>` block in `StyleSheet::parse` SAX loop: track `in_dxfs`, `in_dxf`, `cur_dxf`. Reuse existing `parse_color`, font/fill parsing; push complete `Dxf` on `</dxf>`.
- [x] 1.2.3 **REFACTOR**: Confirm `dxfs` preserved in `original_styles` clone.

### Phase 1.3: Emit `<dxfs>` After `<cellStyles>` (src/writer.rs)

- [x] 1.3.1 **RED**: Write integration test `no_cf_keeps_styles_byte_identical` (scenario 4): open `samples/tables_totals.xlsx`, save to tmp, assert styles.xml contains NO `<dxfs>` and file reopens with no repair. Test fails (no logic yet).
- [x] 1.3.2 **GREEN**: In `write_styles_xml`, after emitting `</cellStyles>`, check `if !collected.dxfs.is_empty()`. If non-empty, emit `<dxfs count=N>`, iterate `collected.dxfs`, call `write_dxf(&mut xml, dxf)` (helper function: emit `<dxf>`, optionally `<font>` with color/bold/italic, optionally `<fill>` with `write_fill`, then `</dxf>`). Gate ensures zero-CF files emit nothing.
- [x] 1.3.3 Test: Confirm `tables_totals.xlsx` and `inline_strings_no_sst.xlsx` byte-equality tests stay green (no `<dxfs>` emitted).

### Phase 1.4: RMW Dxf Preservation in `collect_styles`

- [x] 1.4.1 **RED**: Write integration test `rmw_preserves_source_dxfs`: Create fixture with openpyxl containing CF + dxf; open in oxlsx, save without adding CF, reopen, assert dxfs count unchanged. Test fails (dxfs dropped).
- [x] 1.4.2 **GREEN**: In `collect_styles`, when `preserve_original` and `original_styles.is_some()`, clone `original_styles.dxfs` into `CollectedStyles.dxfs` as the base. New rules (Slice 2) will `find_or_push` into this vec.
- [x] 1.4.3 Test: RMW fixture confirms dxfs preserved.

---

## Slice 2: Rule Model + Emission + Rust API + RMW CF Read (~220 lines)

**Focus**: `ConditionalFormatting`, `ConditionalRule` structs; add CF to Worksheet; parse CF from sheet XML; emit `<conditionalFormatting>` with dxfId mapping; priority assignment.
**Scenarios**: 1,2,3,5,6.

### Phase 2.1: Define ConditionalRule + ConditionalFormatting (src/conditional_formatting.rs — NEW)

- [x] 2.1.1 Create `src/conditional_formatting.rs` with `pub struct ConditionalRule { pub type_ext: String, pub priority: Option<u32>, pub dxf_id: Option<u32>, pub stop_if_true: bool, pub operator: Option<String>, pub formulas: Vec<String>, pub dxf: Option<Dxf>, pub unmodeled_attrs: Vec<(String, String)>, pub unmodeled_children: String }`.
- [x] 2.1.2 Add `pub struct ConditionalFormatting { pub sqref: String, pub rules: Vec<ConditionalRule> }`.
- [x] 2.1.3 Add `mod conditional_formatting; pub use conditional_formatting::*;` to `src/lib.rs`.

### Phase 2.2: Add CF Collection to Worksheet (src/reader.rs)

- [x] 2.2.1 Add `pub conditional_formatting: Vec<ConditionalFormatting>` to `Worksheet` struct; initialize as empty vec in constructors.
- [x] 2.2.2 Add `pub fn add_conditional_formatting(&mut self, sqref: impl Into<String>, mut rule: ConditionalRule)` that assigns `rule.priority = Some(self.max_cf_priority() + 1)`, finds or creates `ConditionalFormatting` for `sqref`, and pushes the rule.
- [x] 2.2.3 Add helper `fn max_cf_priority(&self) -> u32` that iterates all CF rules and returns max priority (0 if none).

### Phase 2.3: Emit `<conditionalFormatting>` in Sheet XML (src/writer.rs)

- [x] 2.3.1 **RED**: Write integration test `formula_rule_write_and_dxf_emitted` (scenario 1): create `ConditionalRule { type_ext: "expression", formulas: vec!["$A1>5"], dxf: Some(solid_red_dxf), ... }`, add to sheet, save, unzip, assert sheet XML contains `<conditionalFormatting sqref="A1:A10">` with `<cfRule type="expression" priority="1" dxfId="0"><formula>$A1&gt;5</formula></cfRule>` AND styles.xml has `<dxfs count="1">`. Test fails.
- [x] 2.3.2 **GREEN**: In `write_sheet_xml`, after `</mergeCells>` (before `<hyperlinks>`), iterate `ws.conditional_formatting`. For each CF block, emit `<conditionalFormatting sqref=...>`, iterate its rules, for each rule: look up `dxfId` from a pre-computed map `rule_dxf_ids: HashMap<(sheet_idx, sqref, rule_idx), u32>`, emit `<cfRule type=... priority=... [operator=...] dxfId=... [stopIfTrue="1"]>`, iterate `formulas` and emit XML-escaped `<formula>`, append `unmodeled_children` as-is, `</cfRule>`. Close `</conditionalFormatting>`.
- [x] 2.3.3 Build `rule_dxf_ids` map in `collect_styles`: iterate all sheets' CF rules; for each `rule.dxf`, call `find_or_push(&mut dxfs, dxf)` to get `dxf_id`; store `(sheet_idx, sqref, rule_idx) -> dxf_id` in map. Pass map to `write_sheet_xml`.
- [x] 2.3.4 Test: scenario 1 passes.

### Phase 2.4: CellIsRule Serialization (scenario 2)

- [x] 2.4.1 **RED**: Write integration test `cell_is_rule_write` (scenario 2): `ConditionalRule { type_ext: "cellIs", operator: Some("greaterThan"), formulas: vec!["5"], ... }`, assert cfRule has `operator="greaterThan"`.
- [x] 2.4.2 **GREEN**: Writer already emits `operator` attribute if `rule.operator.is_some()`. Pass test.

### Phase 2.5: Round-Trip Read of CF (src/reader.rs)

- [x] 2.5.1 **RED**: Write integration test `conditional_formatting_round_trip` (scenario 3): save a workbook with CF, reopen with oxlsx, assert `ws.conditional_formatting` not empty, rule type/formula/priority/dxf present.
- [x] 2.5.2 **GREEN**: In `parse_sheet` (or new helper `parse_conditional_formatting`), SAX-parse `<conditionalFormatting>` after cells: extract `sqref`, iterate `<cfRule>`, extract `type`, `priority`, `dxfId`, `operator`, `stopIfTrue`, collect `<formula>` children (unescape XML), capture unmodeled attrs/children as strings. Resolve `dxfId` to a cloned `Dxf` from `styles.dxfs[dxf_id]`. Build `ConditionalRule` and push to `ws.conditional_formatting`. Priority remains `Some(parsed_priority)` (no reassignment on read).
- [x] 2.5.3 Test: Round-trip passes.

### Phase 2.6: RMW CF Preservation + dxfId Offset (scenario 5)

- [x] 2.6.1 **RED**: Write integration test `cf_rmw_preserves_source_dxfs_and_rules` (scenario 5): open fixture with CF, add a new rule with a new dxf, save, assert original CF + dxf preserved AND new rule has `dxfId >= source_dxf_count`. Reopen in openpyxl, confirm no repair.
- [x] 2.6.2 **GREEN**: Logic already works: `collect_styles` clones source dxfs first, `find_or_push` for new dxf appends at end, mapping is correct. Confirm test passes.

### Phase 2.7: Multiple Rules + stopIfTrue (scenario 6)

- [x] 2.7.1 **RED**: Write integration test `multiple_rules_priorities_and_stopiftrue` (scenario 6): add 3 rules to different ranges, second rule has `stop_if_true: true`. Assert priorities are "1","2","3" and second rule has `stopIfTrue="1"` attribute.
- [x] 2.7.2 **GREEN**: Writer emits `stopIfTrue="1"` if `rule.stop_if_true` is true; priority assignment in `add_conditional_formatting` handles sequential numbering. Pass test.

---

## Slice 3: Python Bindings + Cross-Engine Tests (~120 lines)

**Focus**: PyO3 classes for `FormulaRule`, `CellIsRule`, `ConditionalFormattingProxy`; pytest cross-engine validation.
**Scenarios**: 7,8,9.

### Phase 3.1: PyFormulaRule + PyCellIsRule (src/python.rs)

- [ ] 3.1.1 **RED**: Write pytest `test_formula_rule_round_trip` (scenario 7): `ws.conditional_formatting.add("A1:A10", FormulaRule(formula=["$A1>5"], fill=PatternFill(...)))`, save, reopen with openpyxl 3.1.2, assert rule present + file opens with no repair. Test fails (class not defined).
- [ ] 3.1.2 **GREEN**: Define `#[pyclass] pub struct PyFormulaRule { formula: Vec<String>, stop_if_true: Option<bool>, fill: Option<Py<PyPatternFill>>, font: Option<Py<PyFont>> }` with `#[new]` constructor. Add `#[pyclass] pub struct PyCellIsRule` with `operator: String`. Add both to `oxlsx` module (`m.add_class`).
- [ ] 3.1.3 Test: Import succeeds in Python.

### Phase 3.2: PyConditionalFormattingProxy (src/python.rs)

- [ ] 3.2.1 **RED**: Write pytest asserting `ws.conditional_formatting.add(...)` works.
- [ ] 3.2.2 **GREEN**: Define `#[pyclass] pub struct PyConditionalFormattingProxy { workbook: Py<PyWorkbook>, sheet_idx: usize }`. Implement `#[pymethods] fn add(&self, sqref: &str, rule: &PyAny)` that locks workbook, extracts `PyFormulaRule` or `PyCellIsRule` from `rule`, builds Rust `ConditionalRule` with resolved `Dxf`, calls `ws.add_conditional_formatting`. Expose `ws.conditional_formatting` property on `PyWorksheet` returning the proxy.
- [ ] 3.2.3 Test: Python add() succeeds, saves, reopens.

### Phase 3.3: Cross-Engine Validation (pytest)

- [ ] 3.3.1 **RED**: Write `test_cell_is_rule` (scenario 8): use `CellIsRule(operator="greaterThan", formula=["5"])`, save, reopen with openpyxl, assert operator read correctly.
- [ ] 3.3.2 **GREEN**: Ensure `operator` serialization correct. Pass test.
- [ ] 3.3.3 **RED**: Write `test_conditional_formatting_with_pattern_fill` (scenario 9): add rule with `PatternFill`, save, reopen with openpyxl, assert dxf fill present.
- [ ] 3.3.4 **GREEN**: Dxf fill serialization already correct from Slice 1. Pass test.

### Phase 3.4: Final Verification

- [ ] 3.4.1 Run full test suite: `cargo test`, `cargo clippy -- -D warnings`, `cargo fmt --check`, `pytest`.
- [ ] 3.4.2 Confirm all 9 scenarios pass (1:1 mapping).
- [ ] 3.4.3 Verify byte-equality discriminators still green (`tables_totals`, `inline_strings_no_sst`).
