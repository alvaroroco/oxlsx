# openpyxl to oxlsx Python mapping

Target: oxlsx 0.2.0. Check versions with `importlib.metadata.version` and inspect the target wheel's stubs. This skill is self-contained; the separate general oxlsx skill is optional.

## Adapt supported calls

| Existing openpyxl use | oxlsx Python replacement / constraint |
|---|---|
| `from openpyxl import Workbook, load_workbook` | `from oxlsx import Workbook, load_workbook`, only after auditing calls and options |
| `Workbook()` / `load_workbook("in.xlsx")` | Same names for the simple path-based case |
| `wb["Data"]`, `wb.active`, `wb.sheetnames`, `wb.worksheets` | Same access shape; handle a missing active sheet |
| `ws["B2"] = value` | Same assignment shape |
| `ws.cell(2, 2, value=42)` / `cell.value = 42` | `ws["B2"] = 42`; coordinate access is `ws.cell(row=2, column=2)` with no `value` argument |
| `ws.append(row)` | Supported; check iterable or dictionary column keys in the actual workload |
| `ws.iter_rows(..., values_only=True)` | Similar arguments, but returns a list, not a generator; `next(...)` and streaming assumptions need adaptation |
| `ws["A1"].fill = PatternFill(...)` | `ws.set_cell_fill("A1", oxlsx.PatternFill(...))`; import the oxlsx constructor |
| `ws.freeze_panes = "A2"`, `ws.auto_filter.ref = "A1:C10"` | Same basic property shapes; validate saved metadata |
| `ws.column_dimensions["B"].width = 20` | `ws.column_width(2, 20)`; one-based column index |
| `ws.row_dimensions[1].height = 24` | `ws.row_height(1, 24)` |
| `ws.merge_cells("A1:C1")` | Same range-string call; do not assume the numeric overload exists |
| `cell.comment = Comment(text, author)` | `ws.set_cell_comment("A1", text, author)` for plain-text comments |
| `cell.hyperlink = url` | `ws.set_cell_hyperlink("A1", url)` |
| `from openpyxl.utils import get_column_letter, column_index_from_string` | Import these functions directly from `oxlsx` |
| `from openpyxl.utils.dataframe import dataframe_to_rows` | `oxlsx.dataframe_to_rows(df, index=False, header=True)`; materialized values only, no MultiIndex |
| `wb.save(path)` | Supported with a path string; normalize `Path` objects explicitly with `str(path)` |

## Stop rather than erase requirements

- `load_workbook(read_only=True, data_only=True, keep_vba=True, keep_links=..., rich_text=...)`: these keyword options are not exposed in oxlsx Python. Cached formula values are not equivalent to formula strings. File-like input is not part of the path-string signature.
- `Workbook(write_only=True)` and `create_sheet`: not exposed in this Python API. Rust sheet creation/streaming support does not make those Python calls available.
- `cell.font = Font(...)`, `cell.number_format = ...`, named-style assignment, and general style copying: do not assume equivalent setters. Reading a style is not proof it can be assigned.
- `ws.move_range(..., translate=True)`: unsupported; do not drop `translate` silently.
- `pd.ExcelWriter(engine="openpyxl")`, `DataFrame.to_excel`, pandas Styler: do not replace the engine string with `oxlsx`; no oxlsx writer engine or Styler integration exists in this version.
- Charts, images, pivots, VBA, or specialized OOXML: treat preservation as unverified unless demonstrated with representative inputs. Do not rebuild an existing workbook from values to bypass a missing feature.

## Semantic differences to test

- Python leading-`=` strings represent formulas. Neither storing a formula nor comparing its text proves calculation. Preserve the existing handling of untrusted text.
- Python iteration is materialized; retain bounded ranges and measure memory for large files. Do not advertise streaming or a speedup without measurements of this workload.
- Date, datetime, time, and duration values are typed; temporal serialization may round to milliseconds. Timezone-aware datetime/time inputs are rejected. Test the application's precision policy rather than coercing silently.
- Save failures remain failures. oxlsx 0.2.0 rejects symbolic-link destinations; do not remove that guard or redirect outputs without agreement.

## Sources for refreshing the mapping

- [oxlsx 0.2.0 Python signatures](https://github.com/alvaroroco/oxlsx/blob/v0.2.0/oxlsx.pyi)
- [openpyxl workbook loading](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.reader.excel.html)
- [openpyxl worksheet API](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.worksheet.worksheet.html)
- [openpyxl styles](https://openpyxl.readthedocs.io/en/stable/styles.html)
