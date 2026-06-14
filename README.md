# oxlsx

A high-performance Rust library for reading and writing `.xlsx` files with full formatting support — designed as an **openpyxl equivalent in Rust**, with Python bindings (PyO3).

> **5–10x faster than openpyxl** on write operations. Streaming reads via `open_readonly()`.

## Status

**v0.1.0 released** — feature-complete (phases 1–13). First stable API.

## Install

### Rust
```toml
# Cargo.toml
[dependencies]
oxlsx = { git = "https://github.com/alvaroroco/oxlsx", tag = "v0.1.0" }
```
> Not yet on crates.io (Phase 14).

### Python (build from source via maturin)
```bash
pip install maturin
maturin develop --features python   # or: maturin build --features python
```
> PyPI wheels are planned for Phase 14.


| Phase | Scope | Status |
|-------|-------|--------|
| 1 | ZIP parsing, styles, `cell.font()` / `cell.fill()` | ✅ Done |
| 2 | Shared strings, `Value::String/Bool`, `cell.text()` | ✅ Done |
| 3 | Multi-sheet, `Value::Date`, openpyxl API parity | ✅ Done |
| 4 | Write support — `Workbook::new()` + `wb.save()` | ✅ Done |
| 5 | Lazy / streaming reader (`open_readonly()`) | ✅ Done |
| 6 | PyO3 Python bindings | ✅ Done |
| 7 | Read-modify-write (open + edit + save) | ✅ Done |
| 8 | Extended write features | ✅ Done |
| 9 | Worksheet iteration API | ✅ Done |
| 10 | Bulk write API | ✅ Done |
| 11 | Rich formatting | ✅ Done |
| 12 | Hyperlinks + plain-text comments | ✅ Done |
| 13 | Freeze panes, print settings, named ranges, tables (lossless RMW) | ✅ Done |
| 14 | Distribution / CI hardening | 🔲 Planned |
| 15 | Pandas engine / tabular adapter | 🔲 Planned |

## Usage

### Reading

```rust
use oxlsx::{Workbook, Value};

fn main() -> Result<(), oxlsx::OxlsxError> {
    let wb = Workbook::open("file.xlsx")?;

    // Access by index or name
    let ws = wb.sheet(0).unwrap();
    let ws = wb.sheet_by_name("Employees").unwrap();

    println!("Sheets: {:?}", wb.sheetnames());
    println!("Active: {}", wb.active().unwrap().title());

    if let Some(cell) = ws.cell("A1") {
        match cell.value() {
            Value::String(s)  => println!("text: {s}"),
            Value::Number(n)  => println!("number: {n}"),
            Value::Bool(b)    => println!("bool: {b}"),
            Value::Date(d)    => println!("date: {d}"),
            Value::Empty      => println!("empty"),
        }

        if let Some(font) = cell.font() {
            println!("bold: {}", font.bold);
        }
        if let Some(fill) = cell.fill() {
            println!("fill: {}", fill.pattern_type);
        }
    }

    Ok(())
}
```

### Writing

```rust
use oxlsx::{Color, Fill, Font, Workbook};

fn main() -> Result<(), oxlsx::OxlsxError> {
    let mut wb = Workbook::new();

    let ws = wb.active_mut().unwrap();
    ws.set_cell("A1", "Hello");
    ws.set_cell("B1", 42.0_f64);
    ws.set_cell("C1", true);
    ws.set_cell_font("A1", Font { bold: true, ..Default::default() });
    ws.set_cell_fill("A1", Fill {
        pattern_type: "solid".to_string(),
        fg_color: Some(Color::Argb("FFFF00".to_string())),
        bg_color: None,
    });

    wb.create_sheet("Data").set_cell("A1", "Sheet 2");

    wb.save("output.xlsx")?;
    Ok(())
}
```

## API (openpyxl parity)

| openpyxl (Python) | oxlsx (Rust) |
|-------------------|--------------|
| `load_workbook("f.xlsx")` | `Workbook::open("f.xlsx")` |
| `Workbook()` | `Workbook::new()` |
| `wb.worksheets` | `wb.worksheets()` |
| `wb.sheetnames` | `wb.sheetnames()` |
| `wb.active` | `wb.active()` / `wb.active_mut()` |
| `wb["Sheet1"]` | `wb.sheet_by_name("Sheet1")` |
| `ws.title` | `ws.title()` |
| `ws["A1"]` | `ws.cell("A1")` |
| `ws["A1"] = value` | `ws.set_cell("A1", value)` |
| `cell.value` → `str/int/float/date` | `cell.value()` → `&Value` |
| `cell.font.bold` | `cell.font()?.bold` |
| `cell.fill.fgColor` | `cell.fill()?.fg_color` |
| `cell.comment` | `cell.comment()` / `ws.set_cell_comment(...)` |
| `cell.hyperlink` | `cell.hyperlink()` / `ws.set_cell_hyperlink(...)` |
| `ws.freeze_panes` | `ws.freeze_panes()` / `ws.set_freeze_panes(...)` |
| `ws.print_area` | `ws.print_area()` / `ws.set_print_area(...)` |
| `ws.print_title_rows` / `print_title_cols` | `ws.print_title_rows()` / `ws.print_title_cols()` (+ setters) |
| `wb.defined_names` | `wb.defined_names()` (global) / `ws.defined_names()` (sheet-local) |
| `ws.add_table(Table(...))` | `ws.add_table(Table::new(...))` |
| `ws.tables` / `del ws.tables["T"]` | `ws.tables()` / `ws.remove_table("T")` |
| `wb.save("out.xlsx")` | `wb.save("out.xlsx")` |

## Benchmarks

Measured against openpyxl 3.1.2 on a 10k-row file (5 columns, mixed types). Absolute numbers are machine-dependent; the ratios are what matter. Reproduce with `cargo bench` + `python3 benches/openpyxl_comparison.py`.

### Write

| Operation | openpyxl | oxlsx | Speedup |
|-----------|:--------:|:-----:|:-------:|
| Single cell | 1.90ms | 0.18ms | **~10.8x** |
| 1k rows × 5 cols | 23.2ms | 3.68ms | **~6.3x** |
| 10k rows × 5 cols (50k cells) | 222.4ms | 54.1ms | **~4.1x** |
| 3 sheets × 1k rows | 35.9ms | 6.80ms | **~5.3x** |

### Read (10k rows, 5 cols)

| Operation | openpyxl | oxlsx | Speedup |
|-----------|:--------:|:-----:|:-------:|
| Metadata only (`read_only` / `open_readonly()`) | 1.47ms | 0.072ms | **~20x** |
| Open + iterate all cells | 151ms (`read_only`) / 216ms (full) | 33.4ms | **~4.5–6.5x** |
| Eager `open()` metadata, 10k | 1.47ms | 28ms | tradeoff: eager loads the whole sheet for random access — use `open_readonly()` for streaming metadata |

> For metadata or large-file streaming, use `open_readonly()` (lazy, ~20x faster than openpyxl). Eager `open()` loads the full sheet to enable random `ws.cell("A1")` access.

Run benchmarks yourself:
```bash
cargo bench                          # Rust (criterion)
python3 benches/openpyxl_comparison.py  # Python baseline
```

## Data Model

| Type | Description |
|------|-------------|
| `Workbook` | Entry point — open or create, access sheets |
| `Worksheet` | Access / set cells by Excel reference (`"A1"`, `"B3"`) |
| `Cell` | Holds `Value` + `StyleId` + shared `Arc<StyleSheet>` |
| `Value` | `Empty \| String(String) \| Number(f64) \| Bool(bool) \| Date(NaiveDate) \| Formula(String)` |
| `StyleSheet` | Registry mapping `StyleId` → `Font` + `Fill` |
| `StyleId` | Newtype `(usize)` — prevents index confusion at compile time |
| `Font` | Bold, italic, size, family, color |
| `Fill` | Pattern type, foreground/background color |
| `Color` | `Argb(String) \| Indexed(u32) \| Theme(u32)` |
| `DefinedName` | Named range — `name`, `value` (+ `comment`, `hidden`); global or sheet-local |
| `Table` | Worksheet table — `display_name`, `ref`, `columns`, optional `TableStyleInfo`; lossless RMW for unmodeled features |

## Architecture

- **SAX only** — `quick-xml` event loop, no DOM, no full-file loading
- **ZIP direct** — reads from the container, no temp disk extraction
- **`Arc<StyleSheet>`** — shared across `Workbook → Worksheet → Cell`, no lifetime params
- **PyO3-ready** — all public structs are `'static` compatible
- **Newtype pattern** — `StyleId(usize)` prevents index confusion at compile time

## Dependencies

| Crate | Role |
|-------|------|
| `quick-xml` | SAX XML parser + writer |
| `zip` | OOXML container (deflate only, minimal wheel size) |
| `thiserror` | Typed error enum |
| `chrono` | Date resolution (`NaiveDate`) |

## Roadmap

### ✅ Done
- [x] ZIP container reading + SAX parsing
- [x] Styles: fonts, fills, colors, `StyleId` newtype
- [x] Shared strings → `Value::String`
- [x] `Value::Bool`, `Value::Date` (resolved at parse time)
- [x] Multi-sheet: `wb.sheetnames()`, `wb.sheet_by_name()`, `ws.title()`
- [x] Write support: `Workbook::new()`, `ws.set_cell()`, `wb.save()`
- [x] Style round-trip: bold font + solid fill survive write → open cycle
- [x] Criterion benchmarks + openpyxl comparison script
- [x] Streaming reader (`open_readonly()`), PyO3 bindings, read-modify-write
- [x] Extended write: formulas, merged cells, column widths, date numFmt
- [x] Iteration API (`iter_rows`/`iter_cols`/`rows`/`columns`/`values`, bounds)
- [x] Bulk write (`append`, insert/delete rows/cols, `move_range`)
- [x] Rich formatting (`number_format`, `alignment`, `border`, `protection`)
- [x] Hyperlinks + plain-text comments
- [x] Freeze panes; print area / print titles
- [x] Named ranges — `wb.defined_names()` (global) + `ws.defined_names()` (sheet-local), conservative lossless passthrough
- [x] Tables — `ws.add_table()`, `ws.tables`, `TableStyleInfo`; **lossless RMW** (unmodeled table features preserved verbatim when untouched)

### 🔲 Planned

- [ ] **Phase 14 — PyPI publish + CI/CD**
  - `maturin publish` workflow for PyPI releases
  - GitHub Actions wheel matrix for Linux, macOS, and Windows
  - Version management — `Cargo.toml` + `pyproject.toml` in sync
  - `crates.io` publish for the Rust library

- [ ] **Phase 15 — Pandas engine / tabular adapter**
  - `engine="oxlsx"` integration via a dedicated adapter
  - DataFrame ↔ XLSX mapping, headers/index handling, and type conversions
  - Explicit adapter instead of monkeypatching by default

## License

Licensed under either of:

- [MIT License](LICENSE-MIT)
- [Apache License, Version 2.0](LICENSE-APACHE)

at your option.
