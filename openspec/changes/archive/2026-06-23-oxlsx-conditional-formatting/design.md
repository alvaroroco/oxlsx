# Design: oxlsx-conditional-formatting

## Technical Approach

Introduce a `ConditionalFormatting` collection on `Worksheet` containing `ConditionalRule` objects, backed by a `Dxf` (differential format) registry inside `StyleSheet`. At write time, rules look up or append their `Dxf` into the global registry (`find-or-push` dedup) to acquire a `dxfId`, and receive an auto-assigned `priority`. RMW cycles parse and preserve source `<dxfs>` and source `<conditionalFormatting>` verbatim (with exact bytes for unmodeled properties) so new rules safely append without collision.

## Architecture Decisions

### Decision: Dxf Registry Location

**Choice**: Add `dxfs: Vec<Dxf>` to `src/styles.rs` `StyleSheet`.
**Alternatives considered**: Store `Dxf` list directly on the `Workbook` or build it ephemerally during write.
**Rationale**: `StyleSheet` is already shared across all sheets via `Arc<StyleSheet>` and handles `fonts` and `fills`. Putting `dxfs` here avoids lifetime issues with PyO3 (`'static` bound), reuses `Font`/`Fill` parsing, and seamlessly integrates with `original_styles` cloning for RMW preservation.

### Decision: dxfId deduplication vs sequential assignment

**Choice**: Use `find-or-push` against `CollectedStyles.dxfs` at write time.
**Alternatives considered**: Strictly sequential `dxfId` assignment per rule.
**Rationale**: Openpyxl deduplicates identical differential formats (e.g., 3 rules with the same red fill share `dxfId="1"`). `find_or_push` perfectly mirrors this AND automatically maps parsed RMW rules to their original `dxfId` because the preserved `Dxf` exists at the exact original index in `CollectedStyles.dxfs`.

### Decision: Priority Assignment Timing

**Choice**: Assign `priority` at ADD time (`ws.add_conditional_formatting()`).
**Alternatives considered**: Assign `priority` at WRITE time.
**Rationale**: Excel requires unique priorities. Assigning at add-time (`max(existing) + 1`) ensures predictable sequential uniqueness that aligns with openpyxl's observed behavior, without requiring mutable access to the sheet model during the read-only serialization phase.

### Decision: RMW Preservation Strategy

**Choice**: Extract known fields (`type`, `priority`, `dxfId`) and capture unmodeled attributes + children as raw strings (`unmodeled_attrs`, `unmodeled_children`).
**Alternatives considered**: Capture the entire `<cfRule>` element as verbatim bytes (`RawTable` style).
**Rationale**: Verbatim byte preservation makes it impossible to rewrite `dxfId` or `priority` if they needed shifting. By parsing the container but keeping unknown inners (like `<colorScale>`) as raw strings, we preserve unmodeled openpyxl data while still participating in the structured CF layout.

## Data Flow

1. **Python API**: `ws.conditional_formatting.add(sqref, FormulaRule(...))`
2. **Proxy**: Translates `FormulaRule` to Rust `ConditionalRule { type_ext: "expression", dxf: Some(Dxf), formulas: [...] }`.
3. **Add**: `ws.add_conditional_formatting` assigns `rule.priority = max_prio + 1` and pushes to `ws.conditional_formatting`.
4. **Write (Styles)**: `collect_styles` clones `original_styles.dxfs`, iterates all CF rules, finds/pushes `rule.dxf`, and builds `rule_dxf_ids` mapping `(sheet_idx, sqref, rule_idx) -> dxf_id`.
5. **Write (Sheet)**: `<conditionalFormatting>` block emitted between `<mergeCells>` and `<hyperlinks>` using the mapped `dxf_id`.

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/conditional_formatting.rs` | Create | Define `ConditionalFormatting` and `ConditionalRule` structs. |
| `src/styles.rs` | Modify | Define `Dxf { fill, font }`; add `dxfs: Vec<Dxf>` to `StyleSheet`; parse `<dxfs>`. |
| `src/reader.rs` | Modify | Add `conditional_formatting` to `Worksheet`; parse `<conditionalFormatting>` in `parse_sheet_data`. |
| `src/writer.rs` | Modify | Update `collect_styles` for dxfs mapping; emit `<dxfs>` after `<cellStyles>`; emit `<conditionalFormatting>` between mergeCells/hyperlinks. |
| `src/python.rs` | Modify | Add `PyFormulaRule`, `PyCellIsRule`, and `PyConditionalFormattingProxy` (`ws.conditional_formatting`). |

## Interfaces / Contracts

```rust
// src/conditional_formatting.rs
#[derive(Debug, Clone)]
pub struct ConditionalFormatting {
    pub sqref: String,
    pub rules: Vec<ConditionalRule>,
}

#[derive(Debug, Clone)]
pub struct ConditionalRule {
    pub type_ext: String,
    pub priority: Option<u32>,
    pub dxf_id: Option<u32>,     // For RMW rules
    pub stop_if_true: bool,
    pub operator: Option<String>,
    pub formulas: Vec<String>,
    pub dxf: Option<Dxf>,        // For new/modeled rules
    pub unmodeled_attrs: Vec<(String, String)>,
    pub unmodeled_children: String,
}
```

## Testing Strategy (3-Slice Chain)

| Slice | What to Test | Approach |
|-------|--------------|----------|
| **1: Infra + Guard** | Dxf infra, styles.xml emission, byte-equality guard | Unit tests for `<dxfs>` parsing/writing. `no_cf_keeps_styles_byte_identical` integration test confirming no `<dxfs>` emission on clean sheets. |
| **2: Model + Read/Write** | FormulaRule/CellIsRule serialization, `priority` assignment, RMW dxf preservation | `formula_rule_write_and_dxf_emitted`, `cell_is_rule_write`, `conditional_formatting_round_trip`, `cf_rmw_preserves_source_dxfs_and_rules`, `multiple_rules_priorities_and_stopiftrue`. |
| **3: Python Parity** | PyO3 proxies, cross-engine compatibility | Pytest fixtures: `test_formula_rule_round_trip`, `test_cell_is_rule`, `test_conditional_formatting_with_pattern_fill` verifying openpyxl opens output with no repair. |

## Migration / Rollout

No data migration required. Feature is strictly additive. Workbooks without conditional formatting will emit byte-identical structures.

## Open Questions

- None. All openpyxl schema nuances (`find-or-push` dxfId, `stopIfTrue="1"`, unmodeled rule preservation) have been addressed.
