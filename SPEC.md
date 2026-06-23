# Specification: oxlsx (Rust OpenXML Library)

## 1. Goal

Create a high-performance, memory-efficient Rust library to read and write `.xlsx` files with full formatting support, designed as an **openpyxl equivalent in Rust** with Python bindings (PyO3).

The #1 priority is **API ergonomics**: the Rust API must mirror openpyxl as closely as the language allows.

## 2. Architecture

- **Engine**: `quick-xml` (SAX-based) for high-speed parsing and writing
- **Storage**: `zip-rs` to handle the Office Open XML container (deflate only)
- **Memory Model**: Iterator-based reader to handle large files without loading the entire workbook into RAM (Phase 5)
- **Style sharing**: `Arc<StyleSheet>` shared across `Workbook → Worksheet → Cell` — no lifetime parameters (PyO3-compatible)
- **Type safety**: `StyleId(usize)` newtype prevents index confusion at compile time

## 3. Data Models

### Styles
- `Color`: Enum — `Argb(String)`, `Indexed(u32)`, `Theme(u32)`
- `Fill`: Pattern fills (solid, gradient) with fg/bg color
- `Font`: Bold, italic, size, family, color
- `StyleId(usize)`: Newtype for cell `s` attribute index
- `StyleSheet`: Central registry mapping `StyleId` → `Font` + `Fill` + `numFmtId`

### Workbook / Worksheet / Cell
- `Workbook`: Main entry point — open existing or create new
- `Worksheet`: Access and set cells by Excel-style reference (`"A1"`, `"B3"`)
- `Cell`: Contains `Value`, `StyleId`, `Arc<StyleSheet>`, and optional `Comment` / `Hyperlink`
- `Value`: `Empty | String(String) | Number(f64) | Bool(bool) | Date(NaiveDate) | Formula(String)`

### Sheet & workbook features
- `DefinedName` / `DefinedNames`: named ranges — global (`wb.defined_names()`) and sheet-local (`ws.defined_names()`)
- `Table` / `TableColumn` / `TableStyleInfo`: worksheet tables with lossless RMW
- `Comment`, `Hyperlink`: cell-attached, ride along with the `Cell`
- `Dxf`: differential format (fill/font) for conditional formatting; `ConditionalFormatting` holds `FormulaRule` / `CellIsRule` per `sqref`
- Sheet-level: freeze panes, print settings (`_xlnm.*`), `auto_filter`

## 4. Phase Scope

### ✅ Phase 1 — Read: ZIP + Styles + Cell API
1. Open `.xlsx` container from ZIP directly (no temp disk extraction)
2. Parse `xl/styles.xml` → `StyleSheet` via SAX
3. Parse `xl/worksheets/sheet1.xml` → cells with `StyleId`
4. `cell.font()` / `cell.fill()` resolve via `Arc<StyleSheet>`

### ✅ Phase 2 — Read: Shared Strings + Value Completion
1. Parse `xl/sharedStrings.xml` → `Vec<String>` (eager, optional file)
2. Resolve `t="s"` cells to `Value::String` at parse time
3. `t="b"` → `Value::Bool`, `t="inlineStr"` → `Value::String`
4. `Cell::text() -> Option<&str>` convenience accessor

### ✅ Phase 3 — Read: Multi-sheet + Dates
1. Dynamic sheet discovery via `xl/workbook.xml` + `xl/_rels/workbook.xml.rels`
2. `wb.worksheets()`, `wb.sheetnames()`, `wb.sheet_by_name()`, `wb.active()`
3. `ws.title()` — sheet name
4. `numFmtId` captured in `CellXf`; `Value::Date(NaiveDate)` resolved at parse time
5. Excel serial date conversion (epoch 1899-12-30)

### ✅ Phase 4 — Write: New Workbooks
1. `Workbook::new()` — default sheet named "Sheet"
2. `wb.create_sheet(name)`, `wb.active_mut()`
3. `ws.set_cell(ref, value)` — `impl Into<Value>`
4. `ws.set_cell_font(ref, Font)` / `ws.set_cell_fill(ref, Fill)`
5. `wb.save(path)` — valid `.xlsx` with all required OOXML parts
6. Style deduplication on save; inline strings (`t="inlineStr"`)

### ✅ Phase 5 — `open_readonly()` Streaming Reader
**Decision**: Option A — fiel a openpyxl. `open_readonly()` expone solo `ws.rows()`, sin `ws.cell("A1")`.

| Mode | `ws.cell("A1")` | `ws.rows()` | Metadata speed |
|------|:-:|:-:|:-:|
| `Workbook::open()` (current) | ✅ | ❌ | 21ms (eager) |
| `Workbook::open_readonly()` (Phase 5) | ❌ | ✅ | ~1ms (lazy) |

1. `Workbook::open_readonly(path) -> Result<ReadonlyWorkbook, OxlsxError>`
2. `ReadonlyWorkbook::sheet(0) -> Option<&ReadonlyWorksheet>`
3. `ReadonlyWorkbook::sheetnames() -> Vec<&str>`
4. `ReadonlyWorksheet::rows() -> impl Iterator<Item = Vec<Cell>>`
5. Parse cells on demand — no HashMap, stream directly from ZIP entry
6. Styles and shared strings still parsed eagerly (small, needed for every cell)
7. Target: metadata 21ms → ~1ms, beat openpyxl on ALL benchmarks

### ✅ Phase 6 — PyO3 Python Bindings
1. `#[pyclass]` / `#[pymethods]` on all public structs
2. `maturin` build system + PyPI publish workflow
3. `cell.value` → native Python `str / int / float / datetime.date`
4. `wb["Sheet1"]` via `__getitem__`, `ws["A1"] = value` via `__setitem__`
5. Parametric parity test suite — same tests run against openpyxl and oxlsx

### ✅ Phase 7 — Read-Modify-Write
1. Open existing `.xlsx`, modify cells, save back
2. Preserve unsupported parts (charts, images, macros, pivot tables)
3. Atomic same-path save via `tempfile`

### ✅ Phase 8 — Extended Write Features
1. Formulas (`Value::Formula`) — `=` prefix auto-detection in Python
2. Merged cells, column widths, row heights
3. Date `numFmtId` on write (`Value::Date` → `numFmtId=14`)
4. Shared strings optimization — `xl/sharedStrings.xml`, `t="s"`

### ✅ Phase 9 — Iteration API
Mirror openpyxl's most-used read patterns:
1. `ws.iter_rows(min_row, max_row, min_col, max_col, values_only)` → generator
2. `ws.iter_cols(min_col, max_col, min_row, max_row, values_only)` → generator
3. `ws.rows` / `ws.columns` / `ws.values` properties
4. `ws.max_row`, `ws.max_column`, `ws.min_row`, `ws.min_column` properties
5. Python bindings for all of the above

### ✅ Phase 10 — Bulk Write API
Common patterns for programmatic report generation:
1. `ws.append(iterable)` — append a row of values at the bottom (list, tuple, dict)
2. `ws.delete_rows(idx, amount)` / `ws.delete_cols(idx, amount)`
3. `ws.insert_rows(idx, amount)` / `ws.insert_cols(idx, amount)`
4. `ws.move_range(cell_range, rows, cols)`
5. Python bindings for all of the above

### ✅ Phase 11 — Rich Cell Formatting
Complete the formatting surface to match openpyxl:
1. `cell.number_format` — read and write custom number format strings
2. `cell.alignment` — `Alignment` struct: horizontal, vertical, wrap_text, indent
3. `cell.border` — `Border` / `Side` structs: style, color per edge
4. `cell.protection` — locked, hidden flags
5. Full style round-trip: open → read styles → modify → save with correct styles

### ✅ Phase 12 — Hyperlinks and Comments
1. `cell.hyperlink` — read/write external hyperlinks (`<hyperlinks>` element + external rels via `sheet_rels.rs`); internal `location` links + tooltip/display preserved on RMW
2. `cell.comment` — read and write plain-text comments (`xl/comments*.xml` + VML drawings)

### ✅ Phase 13 — Freeze Panes, Print Settings, Named Ranges, Tables
1. `ws.freeze_panes` — freeze rows/columns
2. `ws.print_area`, `ws.print_title_rows`, `ws.print_title_cols` (reserved `_xlnm.*` defined names)
3. Named ranges — `wb.defined_names()` (global) + `ws.defined_names()` (sheet-local), conservative lossless passthrough; reserved-name rejection
4. Tables — `ws.add_table(Table)`, `ws.tables`, `ws.remove_table()`, `TableStyleInfo`; **lossless RMW** (unmodeled table features preserved verbatim when the table is not mutated)

### ✅ Phase 14 — Distribution: PyPI + crates.io + CI
1. `maturin` abi3 wheels (single `cp38-abi3` per platform; Linux x86_64/aarch64, macOS x86_64/aarch64, Windows) + sdist
2. PyPI publish via OIDC **Trusted Publishing** (no stored tokens; `pypi`/`testpypi` environments)
3. `crates.io` publish (`cargo publish`, `CARGO_REGISTRY_TOKEN`)
4. Release workflow on GitHub Release published (+ `workflow_dispatch` TestPyPI dry-run); `Cargo.toml` / `pyproject.toml` version sync
5. Shipped as **v0.1.1** to both registries

### ✅ Post-v0.1.0 — openpyxl migration parity
Gaps surfaced by a real openpyxl→oxlsx migration, all delivered with strict TDD + cross-engine parity:
1. **Bulk write** (the real Phase 10) — `append` (list/dict), `insert_rows`/`insert_cols`, `delete_rows`/`delete_cols`, `move_range`; sparse shift engine, merges/dims shift, table-intersection guard, formulas verbatim, Excel grid-bound + overflow validation
2. **Conditional formatting** — `FormulaRule`, `CellIsRule` + `Dxf` (gated `<dxfs>`; no-CF files byte-identical); RMW preserves source dxfs + unmodeled rules (color scales) verbatim
3. **Sheet `auto_filter`** — `ws.auto_filter.ref`
4. **`PatternFill`** + `ws.set_cell_fill`
5. **Utils** — `get_column_letter`, `column_index_from_string`, `dataframe_to_rows` (pandas, duck-typed, no hard dep)

### 🔲 Phase 15 — Pandas engine / tabular adapter
1. `engine="oxlsx"` integration via a dedicated adapter
2. DataFrame ↔ XLSX mapping, headers/index handling, type conversions
3. Explicit adapter instead of monkeypatching by default

## 5. Performance Targets

| Metric | Target | Status |
|--------|--------|--------|
| Read 100k rows with formatting | < 500ms | ✅ ~250ms |
| Memory for 1GB file | < 50MB | ✅ Streaming via Phase 5 |
| Write 10k rows | < 50ms | ✅ 42ms |
| Style resolution (font + fill) | < 10ns | ✅ 1.5ns |
| Beat openpyxl on ALL benchmarks | Phase 5+ | Write: ✅ Read metadata: ✅ |

## 6. API Parity Reference

See `AGENTS.md` for the full openpyxl mapping table and API review checklist.
All public APIs must be verified against:
- https://openpyxl.readthedocs.io/en/stable/api/openpyxl.html
- https://openpyxl.readthedocs.io/en/stable/api/openpyxl.worksheet.worksheet.html
- https://openpyxl.readthedocs.io/en/stable/api/openpyxl.cell.cell.html

## 7. Dependencies

| Crate | Role | Constraint |
|-------|------|-----------|
| `quick-xml` | SAX parser + writer | No DOM |
| `zip` | OOXML container | deflate only (PyO3 wheel size) |
| `thiserror` | Error types | — |
| `chrono` | Date resolution | `default-features = false` |
| `tempfile` | Atomic same-path save | — |
| `indexmap` | Ordered shared strings dedup | — |
