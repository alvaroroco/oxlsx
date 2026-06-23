//! Worksheet tables (`<table>` parts under `xl/tables/`).
//!
//! Mirrors openpyxl's `openpyxl.worksheet.table` model. This module owns the
//! `Table` data model, the SAX parser for `xl/tables/tableN.xml`, the
//! serializer, and header-row column derivation.
//!
//! ## Read-Modify-Write fidelity
//!
//! The thin [`Table`] model covers `displayName`, `ref`, `columns`, and
//! `tableStyleInfo`. Features oxlsx does NOT model — totals row
//! (`totalsRowCount`/`totalsRowShown`), per-column `totalsRowFunction`/
//! `calculatedColumnFormula`, `dxf` overrides, autoFilter criteria, and
//! `sortState` — are preserved via **verbatim retention**: the original
//! `tableN.xml` bytes are kept on read (see [`RawTable`]) and re-emitted
//! byte-for-byte on save as long as the table is not mutated through the API.
//!
//! Unmodeled features are dropped ONLY when a table is mutated (via
//! [`Worksheet::table_mut`](crate::Worksheet::table_mut)), which forces the
//! model serialization path. Tables created with [`add_table`](crate::Worksheet::add_table)
//! have no retained bytes and always use the model path.

use crate::OxlsxError;
use crate::coords::{parse_range, rc_to_a1};
use crate::reader::{Value, Worksheet};
use crate::xml::attr_value;
use quick_xml::Reader;
use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use std::io::BufRead;

/// An Excel table (structured range) attached to a worksheet.
///
/// PyO3-ready: no borrowed fields, no lifetime parameters. The `ref` range is
/// stored as `ref_range` because `ref` is a reserved word in Rust.
///
/// ## Read-Modify-Write fidelity
///
/// Unmodeled table features (totals row, DXF styles, autoFilter criteria,
/// `sortState`, `calculatedColumnFormula`) survive read-modify-write verbatim
/// while the table is not mutated, and are dropped only when the table is
/// changed through the API. See the module-level doc for details.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Table {
    /// Workbook-wide unique table id (`<table id="...">`).
    pub id: u32,
    /// Internal table name (`name` attribute); usually equals `display_name`.
    pub name: String,
    /// User-visible, workbook-wide unique name (`displayName` attribute).
    pub display_name: String,
    /// The table range in A1 notation, e.g. `"A1:C3"` (`ref` attribute).
    pub ref_range: String,
    /// Number of header rows (`headerRowCount`); defaults to `1`.
    pub header_row_count: u32,
    /// Column definitions in left-to-right order.
    pub columns: Vec<TableColumn>,
    /// Optional banded-style descriptor (`<tableStyleInfo>`).
    pub style: Option<TableStyleInfo>,
}

impl Table {
    /// Create a table from a display name and an A1 range, mirroring
    /// openpyxl's `Table(displayName=..., ref=...)`.
    ///
    /// `name` defaults to `display_name`, `header_row_count` to `1`, and `id`
    /// to `0` (the writer assigns a workbook-wide unique id at save time).
    /// Columns are left empty so they can be derived from the header row when
    /// the table is added to a worksheet.
    pub fn new(display_name: &str, ref_range: &str) -> Self {
        Self {
            id: 0,
            name: display_name.to_string(),
            display_name: display_name.to_string(),
            ref_range: ref_range.to_string(),
            header_row_count: 1,
            columns: Vec::new(),
            style: None,
        }
    }
}

/// Retained raw bytes of an `xl/tables/tableN.xml` part read from a source
/// workbook, paired with a dirty flag.
///
/// Stored in a parallel `Worksheet::table_retention` vector (NOT as fields on
/// the public [`Table`] struct, so external `Table { .. }` construction is
/// unaffected and PyO3 stays clean). When a retained table is clean
/// (`dirty == false`) the writer re-emits `bytes` verbatim, preserving every
/// unmodeled feature (totals row, `calculatedColumnFormula`, `sortState`,
/// `dxf`, autoFilter criteria). Acquiring mutable access through
/// [`Worksheet::table_mut`](crate::Worksheet::table_mut) flips `dirty` to
/// `true`, switching that table to the model serialization path.
#[derive(Debug, Clone)]
pub(crate) struct RawTable {
    /// The exact `tableN.xml` bytes as read from the source ZIP.
    pub bytes: Vec<u8>,
    /// Whether the table model was mutated through the API since it was read.
    pub dirty: bool,
}

/// A single column definition within a [`Table`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableColumn {
    /// Column id (`<tableColumn id="...">`), 1-based.
    pub id: u32,
    /// Column header name (`name` attribute).
    pub name: String,
}

/// Banded-style descriptor (`<tableStyleInfo>`), the full openpyxl struct.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableStyleInfo {
    /// Built-in style name, e.g. `"TableStyleMedium9"`.
    pub name: String,
    /// Emphasize the first column.
    pub show_first_column: bool,
    /// Emphasize the last column.
    pub show_last_column: bool,
    /// Apply banded rows.
    pub show_row_stripes: bool,
    /// Apply banded columns.
    pub show_column_stripes: bool,
}

/// Parse an `xl/tables/tableN.xml` part into a [`Table`].
///
/// Unmodeled features (totals row, dxf overrides, calculated columns,
/// autoFilter criteria, sortState) are ignored; this is the conservative
/// minimal-valid read path.
pub(crate) fn parse_table_xml<R: BufRead>(reader: R) -> Result<Table, OxlsxError> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut table = Table {
        id: 0,
        name: String::new(),
        display_name: String::new(),
        ref_range: String::new(),
        header_row_count: 1,
        columns: Vec::new(),
        style: None,
    };

    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) => match e.local_name().as_ref() {
                b"table" => {
                    if let Some(value) = attr_value(&e, b"id")? {
                        table.id = value.parse().unwrap_or(0);
                    }
                    if let Some(value) = attr_value(&e, b"name")? {
                        table.name = value;
                    }
                    if let Some(value) = attr_value(&e, b"displayName")? {
                        table.display_name = value;
                    }
                    if let Some(value) = attr_value(&e, b"ref")? {
                        table.ref_range = value;
                    }
                    if let Some(value) = attr_value(&e, b"headerRowCount")? {
                        table.header_row_count = value.parse().unwrap_or(1);
                    }
                }
                b"tableColumn" => {
                    let id = attr_value(&e, b"id")?
                        .and_then(|value| value.parse().ok())
                        .unwrap_or(0);
                    let name = attr_value(&e, b"name")?.unwrap_or_default();
                    table.columns.push(TableColumn { id, name });
                }
                b"tableStyleInfo" => {
                    table.style = Some(TableStyleInfo {
                        name: attr_value(&e, b"name")?.unwrap_or_default(),
                        show_first_column: parse_xml_bool(attr_value(&e, b"showFirstColumn")?),
                        show_last_column: parse_xml_bool(attr_value(&e, b"showLastColumn")?),
                        show_row_stripes: parse_xml_bool(attr_value(&e, b"showRowStripes")?),
                        show_column_stripes: parse_xml_bool(attr_value(&e, b"showColumnStripes")?),
                    });
                }
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(table)
}

/// Serialize a [`Table`] into a minimal Excel-valid `xl/tables/tableN.xml` part.
///
/// Emits `<table>` with `autoFilter`, `tableColumns`, and (when present)
/// `tableStyleInfo`. Boolean style flags are emitted as `"1"` and omitted when
/// `false`, matching openpyxl. Deferred features (totals row, dxf overrides,
/// calculated columns, autoFilter criteria, sortState) are not emitted; the
/// result still opens in Excel without repair.
///
/// Returns [`OxlsxError::InvalidFormat`] if the table has zero columns (defensive).
pub(crate) fn serialize_table_xml(table: &Table) -> Result<Vec<u8>, OxlsxError> {
    // FIX 2: Defensive guard against 0-column tables.
    if table.columns.is_empty() {
        return Err(OxlsxError::InvalidFormat(format!(
            "table {:?} has zero columns, cannot serialize",
            table.display_name
        )));
    }

    let mut xml = Writer::new(Vec::new());
    xml.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let id = table.id.to_string();
    let header_row_count = table.header_row_count.to_string();
    let mut root = BytesStart::new("table");
    root.push_attribute((
        "xmlns",
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    ));
    root.push_attribute(("id", id.as_str()));
    root.push_attribute(("name", table.name.as_str()));
    root.push_attribute(("displayName", table.display_name.as_str()));
    root.push_attribute(("ref", table.ref_range.as_str()));
    root.push_attribute(("headerRowCount", header_row_count.as_str()));
    xml.write_event(Event::Start(root))?;

    let mut auto_filter = BytesStart::new("autoFilter");
    auto_filter.push_attribute(("ref", table.ref_range.as_str()));
    xml.write_event(Event::Empty(auto_filter))?;

    let column_count = table.columns.len().to_string();
    let mut columns = BytesStart::new("tableColumns");
    columns.push_attribute(("count", column_count.as_str()));
    xml.write_event(Event::Start(columns))?;
    for column in &table.columns {
        let column_id = column.id.to_string();
        let mut element = BytesStart::new("tableColumn");
        element.push_attribute(("id", column_id.as_str()));
        element.push_attribute(("name", column.name.as_str()));
        xml.write_event(Event::Empty(element))?;
    }
    xml.write_event(Event::End(BytesEnd::new("tableColumns")))?;

    if let Some(style) = &table.style {
        let mut element = BytesStart::new("tableStyleInfo");
        element.push_attribute(("name", style.name.as_str()));
        if style.show_first_column {
            element.push_attribute(("showFirstColumn", "1"));
        }
        if style.show_last_column {
            element.push_attribute(("showLastColumn", "1"));
        }
        if style.show_row_stripes {
            element.push_attribute(("showRowStripes", "1"));
        }
        if style.show_column_stripes {
            element.push_attribute(("showColumnStripes", "1"));
        }
        xml.write_event(Event::Empty(element))?;
    }

    xml.write_event(Event::End(BytesEnd::new("table")))?;
    Ok(xml.into_inner())
}

/// Derive [`TableColumn`]s from the header row of a range like `"A1:C3"`.
///
/// Column ids are assigned `1..=N`. Each name is taken from the matching cell
/// in the first row of the range; empty or non-string cells fall back to
/// `"Column{id}"` to guarantee a non-empty name (openpyxl parity for the
/// common string-header case).
///
/// Duplicate names (case-insensitive) are deduplicated by appending a numeric
/// suffix ("Total" → "Total", "Total1", "Total2", ...) to ensure Excel's
/// requirement of unique case-insensitive column names within a table.
pub(crate) fn derive_columns_from_header(ws: &Worksheet, ref_range: &str) -> Vec<TableColumn> {
    let Some(((start_col, header_row), (end_col, _))) = parse_range(ref_range) else {
        return Vec::new();
    };

    let mut seen_names_lower: std::collections::HashMap<String, u32> =
        std::collections::HashMap::new();

    (start_col..=end_col)
        .enumerate()
        .map(|(offset, col)| {
            let id = offset as u32 + 1;
            let reference = rc_to_a1(header_row, col);
            let raw_name = ws
                .cell(&reference)
                .map(|cell| value_to_header(cell.value()))
                .filter(|name| !name.is_empty())
                .unwrap_or_else(|| format!("Column{id}"));

            // FIX 5 CRITICAL 2: Dedup case-insensitively. If there's a collision, append a suffix.
            // Keep incrementing the suffix until we find a unique name (re-collision guard).
            let mut unique_name = raw_name.clone();
            let mut lower = unique_name.to_ascii_lowercase();

            while let Some(count) = seen_names_lower.get_mut(&lower) {
                *count += 1;
                unique_name = format!("{raw_name}{count}");
                lower = unique_name.to_ascii_lowercase();
            }

            seen_names_lower.insert(lower, 0);

            TableColumn {
                id,
                name: unique_name,
            }
        })
        .collect()
}

/// Convert a header cell value into a column name string.
fn value_to_header(value: &Value) -> String {
    match value {
        Value::String(text) | Value::Formula(text) => text.clone(),
        Value::Number(number) => number.to_string(),
        Value::Bool(flag) => {
            if *flag {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        Value::Date(date) => date.to_string(),
        Value::Empty => String::new(),
    }
}

/// Validate that a table `ref` range has at least one column.
/// Returns `true` if the range is valid (>=1 column), `false` otherwise.
pub(crate) fn validate_table_ref(ref_range: &str) -> bool {
    if let Some(((start_col, _), (end_col, _))) = parse_range(ref_range) {
        start_col <= end_col
    } else {
        false
    }
}

/// Interpret an OOXML boolean attribute: present and truthy (`"1"`/`"true"`)
/// means `true`; absent or any other value means `false`.
fn parse_xml_bool(value: Option<String>) -> bool {
    matches!(value.as_deref(), Some("1") | Some("true"))
}

#[cfg(test)]
mod tests {
    use super::{
        Table, TableColumn, TableStyleInfo, derive_columns_from_header, parse_table_xml,
        serialize_table_xml,
    };
    use crate::Workbook;
    use std::io::Cursor;

    const TABLE_WITH_STYLE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table id="1" name="Table1" displayName="Table1" ref="A1:C3" headerRowCount="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1:C3"/><tableColumns count="3"><tableColumn id="1" name="h1"/><tableColumn id="2" name="h2"/><tableColumn id="3" name="h3"/></tableColumns><tableStyleInfo name="TableStyleMedium9" showRowStripes="1"/></table>"#;

    const TABLE_NO_STYLE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table id="7" name="Region" displayName="Region" ref="B2:D2"><tableColumns count="1"><tableColumn id="1" name="only"/></tableColumns></table>"#;

    #[test]
    fn parse_table_xml_reads_core_attributes_and_columns() {
        let table = parse_table_xml(Cursor::new(TABLE_WITH_STYLE.as_bytes())).unwrap();

        assert_eq!(table.id, 1);
        assert_eq!(table.name, "Table1");
        assert_eq!(table.display_name, "Table1");
        assert_eq!(table.ref_range, "A1:C3");
        assert_eq!(table.header_row_count, 1);
        assert_eq!(
            table.columns,
            vec![
                TableColumn {
                    id: 1,
                    name: "h1".to_string()
                },
                TableColumn {
                    id: 2,
                    name: "h2".to_string()
                },
                TableColumn {
                    id: 3,
                    name: "h3".to_string()
                },
            ]
        );
    }

    #[test]
    fn parse_table_xml_reads_style_info_with_bool_flags() {
        let table = parse_table_xml(Cursor::new(TABLE_WITH_STYLE.as_bytes())).unwrap();

        assert_eq!(
            table.style,
            Some(TableStyleInfo {
                name: "TableStyleMedium9".to_string(),
                show_first_column: false,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: false,
            })
        );
    }

    #[test]
    fn parse_table_xml_handles_missing_style_and_defaults() {
        let table = parse_table_xml(Cursor::new(TABLE_NO_STYLE.as_bytes())).unwrap();

        assert_eq!(table.id, 7);
        assert_eq!(table.display_name, "Region");
        assert_eq!(table.ref_range, "B2:D2");
        // headerRowCount absent → defaults to 1.
        assert_eq!(table.header_row_count, 1);
        assert_eq!(table.columns.len(), 1);
        assert_eq!(table.columns[0].name, "only");
        assert!(table.style.is_none());
    }

    #[test]
    fn serialize_table_xml_emits_core_shape_without_style() {
        let table = Table {
            id: 1,
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            ref_range: "A1:C3".to_string(),
            header_row_count: 1,
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "h1".to_string(),
                },
                TableColumn {
                    id: 2,
                    name: "h2".to_string(),
                },
            ],
            style: None,
        };

        let xml = String::from_utf8(serialize_table_xml(&table).unwrap()).unwrap();

        assert!(xml.contains(r#"id="1""#));
        assert!(xml.contains(r#"displayName="Table1""#));
        assert!(xml.contains(r#"name="Table1""#));
        assert!(xml.contains(r#"ref="A1:C3""#));
        assert!(xml.contains(r#"headerRowCount="1""#));
        assert!(xml.contains(r#"<autoFilter ref="A1:C3""#));
        assert!(xml.contains(r#"<tableColumns count="2">"#));
        assert!(xml.contains(r#"<tableColumn id="1" name="h1"/>"#));
        assert!(!xml.contains("tableStyleInfo"));
    }

    #[test]
    fn serialize_table_xml_emits_only_true_style_flags() {
        let table = Table {
            id: 2,
            name: "T".to_string(),
            display_name: "T".to_string(),
            ref_range: "A1:B2".to_string(),
            header_row_count: 1,
            columns: vec![TableColumn {
                id: 1,
                name: "c".to_string(),
            }],
            style: Some(TableStyleInfo {
                name: "TableStyleMedium9".to_string(),
                show_first_column: false,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: false,
            }),
        };

        let xml = String::from_utf8(serialize_table_xml(&table).unwrap()).unwrap();

        assert!(xml.contains(r#"<tableStyleInfo name="TableStyleMedium9" showRowStripes="1"/>"#));
        assert!(!xml.contains("showFirstColumn"));
        assert!(!xml.contains("showColumnStripes"));
    }

    #[test]
    fn derive_columns_from_header_uses_cells_with_empty_fallback() {
        let mut wb = Workbook::new();
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "h1");
        ws.set_cell("C1", "h3");
        // B1 left empty on purpose.

        let columns = derive_columns_from_header(ws, "A1:C3");

        assert_eq!(
            columns,
            vec![
                TableColumn {
                    id: 1,
                    name: "h1".to_string()
                },
                TableColumn {
                    id: 2,
                    name: "Column2".to_string()
                },
                TableColumn {
                    id: 3,
                    name: "h3".to_string()
                },
            ]
        );
    }
}
