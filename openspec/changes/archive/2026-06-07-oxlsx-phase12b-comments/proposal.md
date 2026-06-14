# Proposal: oxlsx Phase 12b Comments

## Intent

Add openpyxl-style comment support so cells can read/write plain text comments plus author metadata, and preserve them across save/reload without disturbing unrelated OOXML parts.

## Scope

### In Scope
- Read plain text comment text + author into `cell.comment`
- Write/update comments and round-trip them through `Workbook::save`
- Add Python parity via `cell.comment` getter plus worksheet-level comment setter compatible with `'static` PyO3 wrappers

### Out of Scope
- Rich text formatting, box dimensions, and comment styling fidelity
- Threaded comments / modern notes beyond legacy OOXML comments
- Generalized drawing/VML editing beyond minimum comment plumbing

## Capabilities

### New Capabilities
- `cell-comments`: Read, mutate, and serialize worksheet cell comments with plain text and author metadata

### Modified Capabilities
- None

## Approach

Treat comments as a narrow slice: plain text + author only in this phase, rich comment fidelity later. Add a lightweight `Comment` domain model stored alongside cell/write state, reuse per-sheet rel handling from hyperlinks, parse `xl/comments*.xml` into a ref→comment map, and write the minimum package additions: comments part, sheet rel entry, sheet `legacyDrawing`, VML stub, and required content-type overrides. In RMW mode, preserve untouched ZIP entries and merge only comment-owned rel/content-type changes.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/reader.rs` | Modified | Parse sheet comment rels, comments XML, and attach comments by cell ref |
| `src/writer.rs` | Modified | Emit comments parts, VML/rels/content-types updates, and preserve unrelated entries |
| `src/workbook_rels.rs` | Modified | Extend rel parsing/helpers for per-sheet relationship targets |
| `src/python.rs` | Modified | Expose comment getter plus worksheet-level setter for Python parity |
| `src/lib.rs` | Modified | Re-export comment type if made public |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Legacy VML plumbing breaks Excel compatibility | Med | Keep generated VML minimal, validate with round-trip fixtures and Excel/openpyxl-open tests |
| RMW overwrites unrelated sheet relationships | High | Merge rels/content-types surgically and copy unknown ZIP entries unchanged |
| PyO3 API drifts from openpyxl ergonomics | Med | Keep read at `cell.comment`; use worksheet setter only as Rust/PyO3 mutation bridge |

## Rollback Plan

Revert comment serialization/parsing changes and disable Python comment bindings, leaving hyperlinks and existing RMW preservation paths untouched.

## Dependencies

- Existing hyperlink-era per-sheet relationship infrastructure
- Comment fixture workbook(s) covering existing-comment read and save/reload

## Success Criteria

- [ ] Loading a workbook with plain text comments exposes comment text and author on the target cell
- [ ] Saving and reloading preserves comments and unrelated workbook relationships/content
- [ ] Python bindings support reading `cell.comment` and writing comments without non-`'static` references
