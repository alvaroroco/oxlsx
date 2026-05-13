# Specification: oxlsx (Rust OpenXML Library)

## 1. Goal

Create a high-performance, memory-efficient Rust library to read and write `.xlsx` files with full formatting support, designed as an **openpyxl equivalent in Rust** with Python bindings (PyO3) planned.

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
- `Cell`: Contains `Value`, `StyleId`, and `Arc<StyleSheet>`
- `Value`: `Empty | String(String) | Number(f64) | Bool(bool) | Date(NaiveDate)`

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

### 🔲 Phase 5 — Lazy / Streaming Reader
1. Iterator-based row streaming: `ws.rows() -> impl Iterator<Item = Row>`
2. Parse cells on demand — no full HashMap load at open time
3. Target: beat openpyxl on metadata-only read (currently 21ms vs 1.1ms)
4. Required for 100k+ rows within the 50MB memory target

### 🔲 Phase 6 — PyO3 Python Bindings
1. `#[pyclass]` / `#[pymethods]` on all public structs
2. `maturin` build system + PyPI publish workflow
3. `cell.value` → native Python `str / int / float / datetime.date`
4. `wb["Sheet1"]` via `__getitem__`, `ws["A1"] = value` via `__setitem__`

### 🔲 Phase 7 — Read-Modify-Write
1. Open existing `.xlsx`, modify cells, save back
2. Preserve unsupported parts (charts, images, macros, pivot tables)

### 🔲 Phase 8 — Extended Write Features
1. Formulas (`Value::Formula`)
2. Merged cells, column widths, row heights
3. Date `numFmtId` on write (`Value::Date` → correct format code)
4. Shared strings optimization for repeated text

## 5. Performance Targets

| Metric | Target | Current (Phase 4) |
|--------|--------|-------------------|
| Read 100k rows with formatting | < 500ms | ~250ms est. |
| Memory for 1GB file | < 50MB | Eager load — Phase 5 required |
| Write 10k rows | < 50ms | **42ms** ✅ |
| Style resolution (font + fill) | < 10ns | **1.5ns** ✅ |
| Beat openpyxl on ALL benchmarks | Phase 5 | Write: ✅ Read metadata: Phase 5 |

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
