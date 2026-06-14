# Design: oxlsx Phase 12b Comments

## Technical Approach

Plain text comments require introducing the first per-sheet rels infrastructure: today `workbook_rels.rs` handles only workbook rels, and `copy_unknown_entries` in `writer.rs` transparently preserves untouched sheet rels/comments/VML. The proposal's "hyperlink-era plumbing" does not yet exist.

Round-trip rule: any sheet with at least one `cell.comment.is_some()` is regenerated end-to-end (sheet XML + sheet rels + comments XML + VML); other sheets fall through `copy_unknown_entries` untouched.

## Architecture Decisions

### Comment model on `Cell`

| Option | Tradeoff | Decision |
|--------|----------|----------|
| Field on `Cell` | Mirrors `cell.comment` openpyxl shape | **Chosen** |
| `HashMap<ref, Comment>` on `Worksheet` | Forces lookup per read; breaks parity | Rejected |

`Comment { text: String, author: String }` is `Clone + 'static`, PyO3-safe.

### New `src/sheet_rels.rs`, not extension of `workbook_rels.rs`

Workbook rels are writer-owned; sheet rels are partially writer-owned (we own `comments`/`vmlDrawing` only). Mixing risks dropping unrelated rels and violating the spec's preservation requirement. Separation also keeps each module under ~200 lines.

### Sheet rels are RMW-merged

When a sheet has comments, read existing `xl/worksheets/_rels/sheetN.xml.rels` from source, drop existing `comments` + `vmlDrawing` entries, append fresh ones with stable ids, re-emit. New sheets: emit only the two comment-related rels. Regeneration from scratch would silently drop printer settings, drawings, tables — explicit spec violation.

### VML emitted from a constant template

| Option | Tradeoff | Decision |
|--------|----------|----------|
| String template | openpyxl-canonical, ~30 lines | **Chosen** |
| Real VML serializer | Brittle mso-namespaced model, out of scope | Rejected |

Shape ids follow openpyxl's `_x0000_s{1024+i}` convention.

### Public API mirrors existing `set_cell_font` / `set_cell_fill`

Rust: `Worksheet::set_cell_comment(&mut self, reference: &str, comment: Comment)`. Python: `ws.set_cell_comment(coord, text, author)` + `cell.comment` getter. Openpyxl's mutable `cell.comment =` would require lifetimes on the PyO3 cell proxy — forbidden by project rules.

## Data Flow

```
open:  sheet rels → comments target → parse_comments_xml → attach to cells
save:  for each sheet with comments → write sheet XML (legacyDrawing rid)
                                   → write comments XML
                                   → write VML
                                   → merge + write sheet rels
       other sheets → fall through copy_unknown_entries
       [Content_Types].xml → add comments + vml overrides
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/comments.rs` | Create | `Comment` struct + `parse_comments_xml` SAX parser |
| `src/sheet_rels.rs` | Create | `SheetRelEntry`, `parse_sheet_rels`, `write_sheet_rels` (merge-aware) |
| `src/reader.rs` | Modify | `Cell.comment`, `Cell::comment()`; in `open` parse per-sheet rels + comments; `Worksheet::set_cell_comment`; store `original_sheet_rels` on `Workbook` |
| `src/writer.rs` | Modify | Emit `<legacyDrawing r:id=...>` in sheet XML; new `write_comments_xml` / `write_vml_drawing` / `write_sheet_rels`; extend `write_content_types`; track written paths so `copy_unknown_entries` skips them |
| `src/lib.rs` | Modify | `pub mod comments; pub mod sheet_rels; pub use comments::Comment;` |
| `src/python.rs` | Modify | `PyComment`, `PyCell::comment`, `PyWorksheet::set_cell_comment` |
| `tests/integration_test.rs` | Modify | 3 new Rust tests (see scenario mapping) |
| `tests/test_python.py` | Modify | 2 new Python tests |
| `samples/comments.xlsx` | Create | Authored via openpyxl; fixture with one comment + one unrelated sheet rel |

## Interfaces

```rust
// src/comments.rs
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment { pub text: String, pub author: String }

// src/sheet_rels.rs
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetRelEntry { pub id: String, pub rel_type: String, pub target: String }

// additions
impl Cell { pub fn comment(&self) -> Option<&Comment>; }
impl Worksheet { pub fn set_cell_comment(&mut self, reference: &str, comment: Comment); }
```

Constants in `sheet_rels.rs`:
- comments rel type: `…/officeDocument/2006/relationships/comments`
- vmlDrawing rel type: `…/officeDocument/2006/relationships/vmlDrawing`
- content-types: `/xl/comments{N}.xml` override + `vml` Default extension

## Testing Strategy

| Layer | What | How |
|-------|------|-----|
| Unit | `parse_comments_xml`, `parse_sheet_rels` round-trip, VML template substitution | `#[cfg(test)] mod tests` with `Cursor<&str>` (pattern from `workbook_rels.rs`) |
| Integration (Rust) | 3 spec scenarios | `tests/integration_test.rs` + `samples/comments.xlsx` + `tempfile::NamedTempFile` (pattern from `phase4_round_trip_*`) |
| Integration (Python) | 2 spec scenarios | `tests/test_python.py`, openpyxl-shaped assertions |

Scenario → test (1:1, mandatory):

| Scenario | Test |
|----------|------|
| `comment_round_trip` | `#[test] fn comment_round_trip()` |
| `comment_write_new_cell` | `#[test] fn comment_write_new_cell()` |
| `sheet_comment_relationships_preserved` | `#[test] fn sheet_comment_relationships_preserved()` |
| `test_comment_round_trip` | `def test_comment_round_trip` |
| `test_comment_write_new_cell` | `def test_comment_write_new_cell` |

The preservation test uses a fixture whose sheet rels include an unrelated entry (e.g. printerSettings); after open→save→reopen the unrelated rel must still exist in the ZIP.

## Migration / Rollout

No migration. Workbooks without comments never enter the new code paths (gated on `cell.comment.is_some()`); `copy_unknown_entries` continues to handle them as today.

## Out of Scope (deferred)

- Rich text comments (runs, fonts, colors)
- Comment box dimensions / positioning beyond minimum VML
- Threaded comments / modern notes
- Generic drawing/VML editing beyond the comment shape stub

## Open Questions

- [ ] `samples/comments.xlsx` should be authored by openpyxl (not by hand) to guarantee Excel-real VML and committed as binary fixture.
