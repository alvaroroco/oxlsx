# Rust consumer guide

Target: oxlsx 0.2.0. Inspect the application's Cargo manifest and lockfile before changing dependencies. A minimal dependency is `oxlsx = "0.2.0"`; match the consuming project's toolchain requirements.

## Create, style, save, and reopen

```rust
use oxlsx::{Font, OxlsxError, Value, Workbook};

fn main() -> Result<(), OxlsxError> {
    let mut wb = Workbook::new();
    let ws = wb.active_mut().expect("new workbook has an active sheet");
    ws.set_cell("A1", "Amount");
    ws.set_cell("A2", 42.0_f64);
    ws.set_cell("A3", Value::Formula("SUM(A2:A2)".into()));
    ws.set_cell_font("A1", Font { bold: true, ..Default::default() });
    wb.create_sheet("Notes").set_cell("A1", "Generated report");
    wb.save("report.xlsx")?;

    let loaded = Workbook::open("report.xlsx")?;
    let sheet = loaded.sheet_by_name("Sheet").expect("saved sheet exists");
    assert_eq!(sheet.cell("A2").unwrap().value(), &Value::Number(42.0));
    assert_eq!(
        sheet.cell("A3").unwrap().value(),
        &Value::Formula("SUM(A2:A2)".into())
    );
    assert!(sheet.cell("A1").unwrap().font().unwrap().bold);
    assert_eq!(loaded.sheetnames(), vec!["Sheet", "Notes"]);
    Ok(())
}
```

The `expect`/`unwrap` calls above assert invariants of this self-created example. For external workbooks, handle missing sheets, cells, and styles explicitly; propagate `OxlsxError` with `?` rather than panicking.

## Choose eager or lazy reading

Use `Workbook::open(path)` for random access, editing, or workflows requiring reported parse errors. Read via `sheet_by_name` / `sheet`, and edit via `sheet_mut(index)` or `active_mut()`; write with `set_cell`, then `save`.

For sequential reads of trusted large workbooks, this block can replace the body of the example's `main` (retain `Ok(())`):

```rust
let wb = Workbook::open_readonly("report.xlsx")?;
if let Some(sheet) = wb.active() {
    for row in sheet.rows() {
        for cell in row {
            println!("{:?}", cell.value());
        }
    }
}
```

- Rows are lazy `Vec<Cell>` items, not `Result` items: do not write `row?`.
- Consume/drop one row iterator before opening another on the same workbook; the ZIP reader is shared through interior mutability.
- Row-stream setup failures can become empty iteration, and XML errors can terminate iteration. Do not present successful iteration as proof that a malformed input was fully read.
- Styles and shared strings are loaded up front; streaming does not imply constant total memory.

## Types and style ownership

`Value` includes `Empty`, `String`, `Number(f64)`, `Bool`, `Date`, `DateTime`, `Time`, `Duration`, and `Formula`. Dates are decoded as typed values, not through a separate date accessor. Add `chrono` as a direct dependency when constructing its temporal types in application code.

Use `Value::Formula("SUM(A1:A2)".into())`, without a leading `=`. A Rust string is not automatically the same as Python's formula assignment.

Read `cell.font()` / `cell.fill()` as optional references. Use `Font { ..., ..Default::default() }` when changing a subset of fields. `Color::Theme(index, tint)` has two fields in 0.2.0; older single-field examples are incompatible. Prefer eight-digit `Color::Argb` values.
