# Python consumer guide

Target: oxlsx 0.2.0. Check `importlib.metadata.version("oxlsx")` in the interpreter running the application. If installation is needed, use `python -m pip install 'oxlsx==0.2.0'` in the application's virtual environment; preserve existing dependency constraints.

## Create, style, save, and reopen

```python
from datetime import date
import oxlsx

wb = oxlsx.Workbook()
ws = wb.active
if ws is None:
    raise RuntimeError("Workbook has no active worksheet")

ws.append(["Name", "Amount", "Date"])
ws.append(["Alice", 42, date(2026, 1, 21)])
ws["B3"] = "=SUM(B2:B2)"
ws.set_cell_fill(
    "A1", oxlsx.PatternFill(fill_type="solid", fgColor="FFFFCC00")
)
ws.freeze_panes = "A2"
wb.save("report.xlsx")

loaded = oxlsx.load_workbook("report.xlsx")
sheet = loaded["Sheet"]
assert sheet["A2"].value == "Alice"
assert sheet["B2"].value == 42
assert sheet["C2"].value == date(2026, 1, 21)
assert sheet["B3"].value == "=SUM(B2:B2)"
assert sheet.freeze_panes == "A2"
assert sheet["A1"].fill is not None

sheet["B2"] = 84
loaded.save("report-edited.xlsx")
assert oxlsx.load_workbook("report-edited.xlsx")["Sheet"]["B2"].value == 84
```

For an existing workbook, start with `oxlsx.load_workbook("input.xlsx")`, select a title from `wb.sheetnames`, assign through `ws["A1"] = value`, and save to a different path. Do not replace the input with a newly created workbook when other sheets must survive.

## Access and iteration

```python
for row in sheet.iter_rows(
    min_row=1, max_row=3, min_col=1, max_col=3, values_only=True
):
    print(row)

cell = sheet.cell(row=2, column=2)  # keyword-only, one-based
print(cell.value)
```

`iter_rows`, `iter_cols`, `rows`, `columns`, and `values` materialize lists; they are not streaming generators. Bound ranges for sparse or large sheets. `wb.active` can be `None`; sheet lookup can fail when a title is absent.

## openpyxl migration boundaries

| Task | oxlsx Python |
|---|---|
| Read / create | `oxlsx.load_workbook(path)` / `oxlsx.Workbook()` |
| Select sheet | `wb[title]`, `wb.worksheets`, `wb.active` |
| Write value | `ws["A1"] = value`, not `cell.value = value` |
| Read style | `cell.font`, `cell.fill`; check optional values |
| Write fill | `ws.set_cell_fill(coord, oxlsx.PatternFill(...))` |
| Create sheets / set fonts | Not exposed in this Python API; do not copy Rust examples |
| Read-only / cached formula / VBA flags | Not accepted by `load_workbook(path)` |

Use eight-digit ARGB colors, such as `FFFFCC00` (opaque yellow). Do not assume every openpyxl style class or setter exists.

## Export DataFrame values

With a DataFrame `df` and worksheet `ws` already available:

```python
for row in oxlsx.dataframe_to_rows(df, index=False, header=True):
    ws.append(row)
```

The helper returns a materialized list; MultiIndex columns/index are unsupported. It exports row values, not pandas Styler formatting. Do not use `df.to_excel(engine="oxlsx")`: an oxlsx pandas writer engine is not implemented in this version. See [best practices](best-practices.md) for unsupported workflows.
