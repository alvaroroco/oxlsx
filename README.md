# oxlsx

A high-performance Rust library for reading and writing `.xlsx` files with full formatting support — designed as an **openpyxl equivalent in Rust**, with Python bindings (PyO3) planned.

> **5–10x faster than openpyxl** on write operations. Read benchmarks improve further with lazy loading (see Roadmap).

## Status

Active development — Phases 1–4 complete. API stabilizing.

| Phase | Scope | Status |
|-------|-------|--------|
| 1 | ZIP parsing, styles, `cell.font()` / `cell.fill()` | ✅ Done |
| 2 | Shared strings, `Value::String/Bool`, `cell.text()` | ✅ Done |
| 3 | Multi-sheet, `Value::Date`, openpyxl API parity | ✅ Done |
| 4 | Write support — `Workbook::new()` + `wb.save()` | ✅ Done |
| 5 | Lazy / streaming reader (beat openpyxl on all benchmarks) | 🔲 Planned |
| 6 | PyO3 Python bindings | 🔲 Planned |
| 7 | Read-modify-write (open + edit + save) | 🔲 Planned |

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
| `wb.save("out.xlsx")` | `wb.save("out.xlsx")` |

## Benchmarks

Measured against openpyxl 3.1.5 on the same 10k-row file (5 columns, mixed types).

### Write

| Operation | openpyxl | oxlsx | Speedup |
|-----------|:--------:|:-----:|:-------:|
| Single cell | 1.57ms | 0.16ms | **9.8x** |
| 1k rows × 5 cols | 22.9ms | 3.1ms | **7.3x** |
| 10k rows × 5 cols | 214.7ms | 42.7ms | **5.0x** |
| 3 sheets × 1k rows | 35.0ms | 4.3ms | **8.1x** |

### Read (10k rows, 5 cols)

| Operation | openpyxl | oxlsx | Notes |
|-----------|:--------:|:-----:|-------|
| Metadata only | 1.1ms | 21ms | oxlsx loads eagerly — `open_readonly()` in Phase 5 fixes this |
| Iterate all cells | 144.7ms | 25.5ms | **5.7x faster** |
| Style resolution | — | **1.5ns** | `Arc<StyleSheet>` lookup |

> Read metadata will improve dramatically in Phase 5 (`open_readonly()` streaming reader).

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
| `Value` | `Empty \| String(String) \| Number(f64) \| Bool(bool) \| Date(NaiveDate)` |
| `StyleSheet` | Registry mapping `StyleId` → `Font` + `Fill` |
| `StyleId` | Newtype `(usize)` — prevents index confusion at compile time |
| `Font` | Bold, italic, size, family, color |
| `Fill` | Pattern type, foreground/background color |
| `Color` | `Argb(String) \| Indexed(u32) \| Theme(u32)` |

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

### 🔲 Planned

- [ ] **Phase 5 — `open_readonly()` streaming reader**
  - `Workbook::open_readonly(path)` — mirrors openpyxl's `read_only=True`
  - `ws.rows()` iterator — streaming SAX, no HashMap, no random access
  - Lazy cell parsing: parse on demand, not at open time
  - Target: metadata 21ms → ~1ms — beat openpyxl on **all** benchmarks
  - Required for files > 100k rows within the 50MB memory target

- [ ] **Phase 6 — PyO3 Python bindings**
  - `#[pyclass]` / `#[pymethods]` on all public structs
  - `maturin` build + publish to PyPI
  - `cell.value` returns native Python `str / int / float / datetime.date`
  - `wb["Sheet1"]` via `__getitem__`

- [ ] **Phase 7 — Read-modify-write**
  - Open existing `.xlsx`, modify cells, save back
  - Preserve unsupported parts (charts, images, macros)

- [ ] **Phase 8 — Extended write features**
  - Formulas, merged cells, column widths
  - Date numFmt on write (`Value::Date` → proper `numFmtId`)
  - Shared strings optimization for repeated text

## License

Licensed under either of:

- [MIT License](LICENSE-MIT)
- [Apache License, Version 2.0](LICENSE-APACHE)

at your option.
