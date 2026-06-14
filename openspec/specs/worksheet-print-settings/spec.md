# worksheet-print-settings Specification

## In Scope
- Worksheet-local read/write APIs for `print_area`, `print_title_rows`, and `print_title_cols`.
- Workbook `definedNames` support ONLY for reserved names `_xlnm.Print_Area` and `_xlnm.Print_Titles`.
- Independent clearing of each setting with Excel-compatible save/reload behavior.
- Round-trip preservation of unrelated defined names in read-modify-write flows when already supported safely.

## Out of Scope
- Generic named ranges, `wb.defined_names`, or user-defined name CRUD.
- Tables and table/name collision handling outside reserved print names.
- Broader print/page setup features beyond these three worksheet properties.

## Semantics Notes
- `_xlnm.Print_Area` stores the worksheet print region and is scoped by `localSheetId` to one worksheet.
- `_xlnm.Print_Titles` stores repeated rows, repeated columns, or both for one worksheet; row and column segments MUST be independently readable/writable and may coexist in one reserved definition.
- Worksheet behavior MUST remain local even though OOXML persists these settings in `workbook.xml`.

## Requirements

### Requirement: Expose worksheet print area

The system MUST expose worksheet-local `print_area` APIs aligned with openpyxl and persist them through reserved `_xlnm.Print_Area` defined names only.

#### Scenario: print_area_round_trip_existing_workbook
- GIVEN a workbook containing a worksheet-local `_xlnm.Print_Area` defined name
- WHEN Rust opens, saves, and reopens the workbook
- THEN `ws.print_area()` returns the same range and the reserved defined name remains Excel-compatible

#### Scenario: print_area_write_new_workbook
- GIVEN a new workbook without print settings
- WHEN Rust sets `print_area`, saves, and reopens the workbook
- THEN the worksheet reports the same print area and `workbook.xml` contains the matching `_xlnm.Print_Area`

#### Scenario: print_area_clear_removes_defined_name
- GIVEN a worksheet with `_xlnm.Print_Area` and unrelated defined names
- WHEN Rust clears `print_area`, saves, and reopens the workbook
- THEN only `_xlnm.Print_Area` is removed for that sheet and unrelated defined names remain unchanged

#### Scenario: test_print_area_round_trip
- GIVEN a workbook containing a worksheet-local `_xlnm.Print_Area` defined name
- WHEN Python opens the workbook and reads `ws.print_area`
- THEN the property matches the workbook contents after save/reload

#### Scenario: test_print_area_write_new_workbook
- GIVEN a new workbook without print settings
- WHEN Python assigns `ws.print_area`, saves, and reloads the workbook
- THEN the property round-trips and the saved workbook contains Excel-compatible `_xlnm.Print_Area`

#### Scenario: test_print_area_clear_removes_defined_name
- GIVEN a worksheet with `_xlnm.Print_Area` and unrelated defined names
- WHEN Python clears `ws.print_area`, saves, and reloads the workbook
- THEN only the reserved print-area effect is removed for that sheet

### Requirement: Expose worksheet print title rows and columns

The system MUST expose worksheet-local `print_title_rows` and `print_title_cols` APIs aligned with openpyxl and persist them through reserved `_xlnm.Print_Titles` defined names only.

#### Scenario: print_title_rows_and_cols_round_trip_existing_workbook
- GIVEN a workbook containing worksheet-local `_xlnm.Print_Titles` with row and column segments
- WHEN Rust opens, saves, and reopens the workbook
- THEN `ws.print_title_rows()` and `ws.print_title_cols()` return the same values and the reserved defined name remains Excel-compatible

#### Scenario: print_title_rows_and_cols_write_new_workbook
- GIVEN a new workbook without print titles
- WHEN Rust sets `print_title_rows` and `print_title_cols`, saves, and reopens the workbook
- THEN both properties round-trip and `workbook.xml` contains the matching `_xlnm.Print_Titles`

#### Scenario: print_title_rows_and_cols_clear_removes_defined_name
- GIVEN a worksheet with `_xlnm.Print_Titles` and unrelated defined names
- WHEN Rust clears title rows and/or title columns, saves, and reopens the workbook
- THEN the reserved print-titles effect is removed or reduced only for the cleared parts without mutating unrelated defined names

#### Scenario: test_print_title_rows_and_cols_round_trip
- GIVEN a workbook containing worksheet-local `_xlnm.Print_Titles` with row and column segments
- WHEN Python opens the workbook and reads `ws.print_title_rows` and `ws.print_title_cols`
- THEN both properties match the workbook contents after save/reload

#### Scenario: test_print_title_rows_and_cols_write_new_workbook
- GIVEN a new workbook without print titles
- WHEN Python assigns `ws.print_title_rows` and `ws.print_title_cols`, saves, and reloads the workbook
- THEN both properties round-trip and the saved workbook contains Excel-compatible `_xlnm.Print_Titles`

#### Scenario: test_print_title_rows_and_cols_clear_removes_defined_name
- GIVEN a worksheet with `_xlnm.Print_Titles` and unrelated defined names
- WHEN Python clears title rows and/or title columns, saves, and reloads the workbook
- THEN only the reserved print-titles effect for the cleared parts is removed

### Requirement: Preserve reserved print definitions safely in workbook serialization

The system SHALL round-trip reserved print-setting defined names through `workbook.xml` without expanding into generic named-range support, and SHOULD preserve unrelated defined definitions unchanged in read-modify-write flows when the codebase supports that safely.

#### Scenario: reserved_defined_names_round_trip_only
- GIVEN a workbook containing reserved print definitions and unsupported generic defined names
- WHEN the workbook is opened, modified only through worksheet print settings, and saved
- THEN reserved print definitions reflect worksheet state while unsupported generic definitions are not promoted into public API surface

#### Scenario: clear_one_setting_preserves_other_reserved_effects
- GIVEN a worksheet where `_xlnm.Print_Titles` encodes both repeated rows and repeated columns
- WHEN one of `print_title_rows` or `print_title_cols` is cleared
- THEN the remaining repeated title setting still persists through the same reserved definition if still present
