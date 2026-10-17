---
name: oxlsx
description: "Trigger: oxlsx, openpyxl migration, Excel .xlsx in Rust or Python. Read, create, edit, and style workbooks with verified APIs."
license: MIT OR Apache-2.0
metadata:
  author: alvaroroco
  version: "1.0"
---

# Use oxlsx

## Activation Contract

Use for applications consuming oxlsx in Python or Rust, including workbook round trips, formatting, tabular exports, and openpyxl migrations. Do not treat generic spreadsheet requests as a requirement to replace an existing library.

## Hard Rules

- Check the installed oxlsx version and language before choosing APIs. These references target 0.2.0; Python and Rust are not interchangeable or fully openpyxl-compatible.
- Use only exposed APIs. Do not invent Python `create_sheet`, `read_only`, `data_only`, `keep_vba`, writable `Cell.value`, or writable `Cell.font`.
- Preserve the original workbook by saving to a new destination unless replacement is requested. Do not promise lossless preservation of arbitrary Excel features.
- Store formulas without claiming evaluation. Treat external text beginning with `=` as potentially executable spreadsheet content in Python.

## Decision Gates

| Need | Load / choose |
|---|---|
| Python usage or migration | [Python guide](references/python.md) |
| Rust usage or large-file streaming | [Rust guide](references/rust.md) |
| Compatibility, formulas, dates, preservation, pandas | [Best practices](references/best-practices.md) |

## Execution Steps

1. Identify the input, output, language, package version, and features the workbook must retain. Inspect existing dependencies before suggesting installation.
2. Load the relevant language guide; consult best practices for compatibility-sensitive work. If a required API is absent, explain the gap rather than silently dropping the feature.
3. Implement the smallest supported flow. Use bounded row/column ranges and avoid unnecessary workbook copies.
4. Test with representative data. Save, reopen, and assert cell values and relevant styles/metadata; formula text round trips are not calculation tests.

## Output Contract

Return runnable code in the requested language, required dependencies, and relevant limitations. Distinguish checks actually executed from suggested checks. Do not claim a workbook is Excel-compatible solely because saving succeeded.

## References

- [Python API and examples](references/python.md)
- [Rust API and examples](references/rust.md)
- [Compatibility and best practices](references/best-practices.md)
