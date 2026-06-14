# Tasks: oxlsx Phase 12b Comments

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 320-520 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 → rels/model + Rust read/write core; PR 2 → Python parity + fixtures/tests |
| Delivery strategy | auto-chain |
| Chain strategy | feature-branch-chain |

Decision needed before apply: No
Chained PRs recommended: Yes
Chain strategy: feature-branch-chain
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Comment model + per-sheet rels + Rust comment round-trip | PR 1 | Base `feat/phase12b-comments`; include read/write core and Rust tests |
| 2 | Python comment parity + fixture polish + preservation test | PR 2 | Base PR 1 branch; include Python wrappers/tests and RMW preservation check |

## Phase 1: Foundation / rels model

- [x] 1.1 Add `src/comments.rs` and `src/sheet_rels.rs` with `Comment`, `SheetRelEntry`, and SAX parse/write helpers for comments + per-sheet rels.
- [x] 1.2 Extend `Cell` in `src/reader.rs` with `comment: Option<Comment>` and add `Worksheet::set_cell_comment(...)` plus `Cell::comment()`.
- [x] 1.3 Add `Workbook`/sheet rel tracking needed to preserve unknown sheet relationships in RMW flows.

## Phase 2: Core read/write implementation

- [x] 2.1 Parse sheet rels, resolve comment targets, and attach parsed comments to cells on `Workbook::open`.
- [x] 2.2 Serialize comments XML, legacyDrawing/VML plumbing, and merged sheet rels on `Workbook::save`.
- [x] 2.3 Extend `[Content_Types].xml` handling for comments XML and VML defaults/overrides.
- [x] 2.4 Update ZIP preservation logic so known comment-owned entries are rewritten while unknown sheet rels remain copied.

## Phase 3: Integration / parity wiring

- [x] 3.1 Expose `Comment` through `src/lib.rs` and wire `PyComment`, `PyCell::comment`, and `PyWorksheet::set_cell_comment` in `src/python.rs`.
- [x] 3.2 Add/refresh the comments fixture workbook with a plain-text comment, author metadata, and at least one unrelated sheet relationship.
- [x] 3.3 Map openpyxl-like assertions to the Rust/Python APIs without adding rich-comment behavior.

## Phase 4: Testing / verification

- [x] 4.1 Add Rust integration tests for `comment_round_trip` and `comment_write_new_cell` using the comments fixture.
- [x] 4.2 Add Rust integration test for `sheet_comment_relationships_preserved` that inspects the saved ZIP rels after RMW.
- [x] 4.3 Add Python parity tests `test_comment_round_trip` and `test_comment_write_new_cell`.
- [x] 4.4 Add focused unit tests for comments XML parsing and per-sheet rel merging/serialization.

## Phase 5: Cleanup / deferred scope guard

- [ ] 5.1 Remove any temporary scaffolding once the comment path is stable.
- [ ] 5.2 Keep rich text, dimensions, threaded comments, and general VML editing explicitly out of scope for this phase.
