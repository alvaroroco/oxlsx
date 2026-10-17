# Validate migration parity

Use an isolated directory and representative non-sensitive copies. Keep the original implementation available as the oracle. Do not execute application entry points until their file, network, and other side effects are understood and redirected.

## Supported before/after example

Both examples read the same `input.xlsx` containing a worksheet titled `Data`. Neither overwrites that input. The changes deliberately adapt setters, not just imports.

Before, using openpyxl:

```python
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

wb = load_workbook("input.xlsx")
ws = wb["Data"]
ws.cell(row=2, column=2, value=42)
ws["B3"] = "=SUM(B2:B2)"
ws["A1"].fill = PatternFill(fill_type="solid", fgColor="FFFFCC00")
ws.freeze_panes = "A2"
wb.save("baseline.xlsx")
```

After, using oxlsx:

```python
import oxlsx

wb = oxlsx.load_workbook("input.xlsx")
ws = wb["Data"]
ws["B2"] = 42
ws["B3"] = "=SUM(B2:B2)"
ws.set_cell_fill("A1", oxlsx.PatternFill(fill_type="solid", fgColor="FFFFCC00"))
ws.freeze_panes = "A2"
wb.save("migrated.xlsx")
```

Independently reopen both with openpyxl for this test; it can stay a development dependency even when application imports migrate:

```python
from openpyxl import load_workbook

baseline = load_workbook("baseline.xlsx", data_only=False)
migrated = load_workbook("migrated.xlsx", data_only=False)
try:
    assert baseline.sheetnames == migrated.sheetnames
    for name in baseline.sheetnames:
        left, right = baseline[name], migrated[name]
        assert (left.max_row, left.max_column) == (right.max_row, right.max_column)
        assert list(left.values) == list(right.values)
    left, right = baseline["Data"], migrated["Data"]
    assert left["B2"].value == right["B2"].value == 42
    assert left["B3"].value == right["B3"].value == "=SUM(B2:B2)"
    assert left["A1"].fill.patternType == right["A1"].fill.patternType == "solid"
    assert left["A1"].fill.fgColor.rgb == right["A1"].fill.fgColor.rgb == "FFFFCC00"
    assert left.freeze_panes == right.freeze_panes == "A2"
finally:
    baseline.close()
    migrated.close()
```

This proves only the example's exercised cells, sheet order, fill, and freeze panes. It is not a universal Excel compatibility test. Comparing ZIP bytes is normally inappropriate because serializers can reorder equivalent XML.

## Extend checks for the actual application

1. Capture baseline return values, exceptions, emitted files, and workbook observations for each migrated path. Add regression cases for the input shapes the application actually consumes.
2. Compare sheet order/titles, value types, formulas, relevant dates, merged ranges, styles, tables, named ranges, comments, links, and print settings only where required. Validate every required unsupported feature before declaring a path migrated.
3. Check failure behavior, missing/empty cells, formula-like text, and temporal precision when exercised. Keep the source file unchanged, including on failure; hashing that file before/after is useful evidence.
4. If calculation results matter, use a separately authorized spreadsheet application/recalculation step and compare results, not merely expression text. Do not reinterpret missing cached values as successful parity.
5. When performance motivates migration, measure equivalent end-to-end operations and peak memory with representative workbooks. Compare eager versus eager, not unlike access modes.

Report exact passed checks and remaining limitations. Keep blocked workflows on openpyxl rather than adding an implicit fallback that silently changes behavior.
