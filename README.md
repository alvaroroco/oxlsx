# oxlsx

A fast, low-level Rust library for reading `.xlsx` files with full formatting support.

Targets the gap between fast-but-format-blind parsers and slow-but-complete ones: oxlsx gives you cell values **and** their styles (fonts, fills, colors) at high throughput.

## Status

Early development — MVP scope (Phase 1). API will change.

## What it does

- Reads `.xlsx` containers (Office Open XML / ZIP)
- Parses `xl/styles.xml` into a typed `StyleSheet` (fonts, fills, colors)
- Parses `xl/worksheets/sheet1.xml` and exposes cells with their associated style
- Resolves `StyleId` → `Font` / `Fill` directly from each `Cell`

## Usage

```rust
use oxlsx::Workbook;

fn main() -> Result<(), oxlsx::OxlsxError> {
    let wb = Workbook::open("file.xlsx")?;
    let ws = wb.sheet(0).unwrap();

    if let Some(cell) = ws.cell("A1") {
        println!("value: {:?}", cell.value());

        if let Some(font) = cell.font() {
            println!("bold: {}", font.bold);
        }
        if let Some(fill) = cell.fill() {
            println!("fill pattern: {}", fill.pattern_type);
        }
    }

    Ok(())
}
```

## Data model

| Type | Description |
|------|-------------|
| `Workbook` | Entry point — open a file, access sheets by index |
| `Worksheet` | Access cells by reference string (`"A1"`, `"B3"`) |
| `Cell` | Holds a `Value`, a `StyleId`, and a shared `StyleSheet` ref |
| `Value` | `Empty`, `String`, `Number(f64)`, `Bool` |
| `StyleSheet` | Registry mapping `StyleId` → `Font` + `Fill` |
| `Font` | Bold, italic, size, family, color |
| `Fill` | Pattern type, foreground/background color |
| `Color` | ARGB hex, theme ID, or indexed |

## Performance targets

- 100k rows with formatting in < 500ms
- Memory < 50MB for files up to 1GB

Achieved via SAX-based parsing (`quick-xml`) and iterator-friendly design — the file is never loaded fully into memory.

## Dependencies

| Crate | Role |
|-------|------|
| `quick-xml` | SAX XML parser |
| `zip` | OOXML container (deflate only) |
| `thiserror` | Typed error enum |

## Roadmap

- [ ] Shared strings table (`xl/sharedStrings.xml`)
- [ ] Multi-sheet support
- [ ] Iterator-based row streaming
- [ ] Write support
- [ ] `Date` value variant

## License

Licensed under either of:

- [MIT License](LICENSE-MIT)
- [Apache License, Version 2.0](LICENSE-APACHE)

at your option.
