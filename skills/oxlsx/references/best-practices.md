# Compatibility and best practices

## Version and evidence

These examples target the repository's 0.2.0 API, not every future release. Check the installed package, its type stubs/Rust docs, and representative files before migrating production workflows. Do not generalize README benchmark ratios into guarantees for a user's workload.

## Preserve workbook intent

- Edit a loaded workbook rather than reconstructing it when existing sheets, formatting, or metadata matter. Save to a separate destination first and reopen it.
- Validate the features the user actually relies on: styles, dates, comments, links, tables, named ranges, or print settings. A successful save alone is insufficient.
- Do not claim universal lossless OOXML round trips, macro preservation, or `.xls` / `.xlsb` support. Test specialized features using genuine external-authoring fixtures and, where needed, Excel or LibreOffice. Keep openpyxl-authored regression fixtures independent of oxlsx output.
- If preservation requirements exceed verified support, explain the limitation and retain the original tool or propose a separately validated conversion; do not silently discard features.

## Values, dates, and formulas

- Python strings starting with `=` become formulas; Rust uses explicit `Value::Formula` with no leading `=`. Do not pass untrusted external text through Python assignment without deciding whether formulas are allowed. Do not claim an unverified escape strategy preserves text exactly.
- oxlsx stores formula expressions; it is not an Excel calculation engine. Test preserved expressions separately from externally recalculated results.
- `move_range(..., translate=True)` is unsupported. Moving cells is not a promise to update every dependent formula or range reference.
- Use native date/datetime/time/timedelta values in Python and corresponding `Value` variants in Rust. Reject or deliberately normalize timezone-aware values before writing; oxlsx does not accept timezone-aware datetime/time inputs.
- Temporal round trips can round to milliseconds; use appropriate precision in assertions. Avoid representing identifiers or precision-critical decimal amounts as binary floating-point without an explicit application policy.

## Workload and failure handling

- Use bounded rectangular iteration; sparse sheets with distant cells can create large materialized ranges. Python row iteration and `dataframe_to_rows` allocate lists.
- Prefer Rust read-only iteration for sequential trusted inputs when its error-reporting limitations are acceptable. Use eager reads when explicit parsing failures matter; see [Rust guide](rust.md).
- Do not suppress load/save exceptions or report success before checking the output. Atomic saves in 0.2.0 do not replace backups, and symbolic-link destinations are rejected.
- Importing oxlsx does not require pandas. `dataframe_to_rows` is a value adapter, not `ExcelWriter` integration; pandas Styler and MultiIndex export are not supported here.

## Maintainer sources

Use these authoritative sources to refresh this skill when the target version changes:

- [oxlsx repository and README](https://github.com/alvaroroco/oxlsx)
- [Python stubs](https://github.com/alvaroroco/oxlsx/blob/v0.2.0/oxlsx.pyi)
- [Rust API documentation](https://docs.rs/oxlsx/0.2.0/oxlsx/)
- [openpyxl API](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.html) for comparison only; an openpyxl method existing does not establish oxlsx support.
- [skills CLI](https://github.com/vercel-labs/skills) for package discovery and installation conventions.
