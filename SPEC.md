# Specification: oxlsx (Rust OpenXML Library)

## 1. Goal
Create a high-performance, memory-efficient Rust library to read and write .xlsx files, focusing on preserving and manipulating cell formatting (styles), similar to Python's openpyxl.

## 2. Architecture
- **Engine**: Use `quick-xml` (SAX-based) for high-speed parsing.
- **Storage**: Use `zip-rs` to handle the Office Open XML container.
- **Memory Model**: Implement an iterator-based reader to handle large files without loading the entire workbook into RAM.

## 3. Data Models
### Styles
- `Color`: Enum supporting ARGB Hex, Theme IDs, and Indexed colors.
- `Fill`: Struct for pattern fills (solid, gradient).
- `Font`: Struct for bold, italic, size, and family.
- `StyleSheet`: The central registry mapping `style_id` (from cell `s` attribute) to specific formatting.

### Workbook/Worksheet
- `Workbook`: Main entry point.
- `Worksheet`: Access to cells, dimensions, and merged cells.
- `Cell`: Contains `Value` (Enum: String, Number, Bool, Date) and an optional `Style`.

## 4. MVP Scope (Phase 1)
1. Unzip .xlsx container.
2. Parse `xl/styles.xml` into a `StyleSheet` struct.
3. Parse `xl/worksheets/sheet1.xml` to extract raw values and their associated `StyleId`.
4. Map `StyleId` to actual `Fill` and `Font` properties.

## 5. Performance Targets
- Read 100k rows with formatting in < 500ms.
- Memory usage < 50MB for files up to 1GB.
