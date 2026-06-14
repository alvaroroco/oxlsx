# Tasks: oxlsx-phase4-write

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 450-650 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 → PR 2 |
| Delivery strategy | auto-chain |
| Chain strategy | feature-branch-chain |

Decision needed before apply: No
Chained PRs recommended: Yes
Chain strategy: feature-branch-chain
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Workbook/Worksheet write API + style model plumbing | PR 1 | Base = feature/tracker branch; covers creation, mutation, and style collection helpers. |
| 2 | OOXML writer + integration/regression verification | PR 2 | Base = PR 1 branch; covers zip parts, round-trip tests, and full regression run. |

## Phase 1: Foundation / API surface

- [ ] 1.1 Add failing tests for `Workbook::new`, `Workbook::create_sheet`, `Workbook::active_mut`, and `Value` `From` impls in the workbook/worksheet test module.
- [ ] 1.2 Implement `Workbook::new/create_sheet/active_mut` and `impl From<&str/i32/f64/bool> for Value` in the workbook/value modules.
- [ ] 1.3 Verify `active_mut` returns a mutable sheet reference and newly created sheets become active by default.

## Phase 2: Core Worksheet mutation

- [ ] 2.1 Add failing tests for `Worksheet::set_cell`, `Worksheet::set_cell_font`, and `Worksheet::set_cell_fill` covering string, number, bool, font, and fill writes.
- [ ] 2.2 Implement the worksheet mutation APIs so they update cell value, style, and sheet cell map without breaking existing reads.
- [ ] 2.3 Ensure style assignment preserves shared style ownership and does not clone sheet-level style state unnecessarily.

## Phase 3: Style sorting and deduplication

- [ ] 3.1 Add tests for `cell_sort_key` ordering and `collect_styles` deduplication on repeated font/fill combinations.
- [ ] 3.2 Implement `cell_sort_key` and refactor `collect_styles` to emit stable, de-duplicated style records for writer consumption.
- [ ] 3.3 Verify sorted cell iteration keeps style references deterministic across single-sheet and multi-sheet cases.

## Phase 4: OOXML writer implementation

- [ ] 4.1 Add failing tests for `write_xlsx` generating the 6 OOXML parts: `[Content_Types].xml`, `_rels/.rels`, `docProps/*`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/styles.xml`, `xl/worksheets/sheet*.xml`.
- [ ] 4.2 Implement `writer.rs::write_xlsx` to serialize workbook metadata, styles, worksheets, and package relationships into the zip container.
- [ ] 4.3 Confirm the writer emits workbook, worksheet, and style XML that matches openpyxl-compatible structure for the supported sample data.

## Phase 5: Integration and regression verification

- [ ] 5.1 Add round-trip integration tests for string, number, bool, font, fill, and multi-sheet write/read behavior using the sample fixtures.
- [ ] 5.2 Run the full existing test suite to ensure the current 29 tests still pass after write-path changes.
- [ ] 5.3 Run `cargo clippy -- -D warnings` and `cargo fmt --check` to validate final polish and lint cleanliness.
