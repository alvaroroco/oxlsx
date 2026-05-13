use crate::{Fill, Font, OxlsxError, StyleSheet};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufRead, BufReader};
use std::path::Path;
use std::sync::Arc;
use zip::result::ZipError;

#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Empty,
    SharedStringIdx(u32),
    Number(f64),
}

#[derive(Debug, Clone)]
pub struct Cell {
    reference: String,
    value: Value,
    style_id: u32,
    styles: Arc<StyleSheet>,
}

impl Cell {
    pub fn reference(&self) -> &str {
        &self.reference
    }

    pub fn value(&self) -> &Value {
        &self.value
    }

    pub fn style_id(&self) -> u32 {
        self.style_id
    }

    pub fn font(&self) -> Option<&Font> {
        self.styles.font(self.style_id)
    }

    pub fn fill(&self) -> Option<&Fill> {
        self.styles.fill(self.style_id)
    }
}

#[derive(Debug, Clone, Default)]
pub struct Worksheet {
    cells: HashMap<String, Cell>,
}

impl Worksheet {
    pub fn cell(&self, reference: &str) -> Option<&Cell> {
        self.cells.get(reference)
    }
}

#[derive(Debug, Clone, Default)]
pub struct Workbook {
    sheets: Vec<Worksheet>,
}

impl Workbook {
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self, OxlsxError> {
        let file = File::open(path)?;
        let mut archive = zip::ZipArchive::new(file).map_err(|err| match err {
            ZipError::InvalidArchive(_) | ZipError::UnsupportedArchive(_) => {
                OxlsxError::InvalidFormat(err.to_string())
            }
            _ => OxlsxError::Zip(err),
        })?;

        let styles = {
            let styles_file = archive
                .by_name("xl/styles.xml")
                .map_err(|_| OxlsxError::MissingPart("xl/styles.xml".to_string()))?;
            Arc::new(StyleSheet::parse(BufReader::new(styles_file))?)
        };

        let worksheet = {
            let sheet_file = archive
                .by_name("xl/worksheets/sheet1.xml")
                .map_err(|_| OxlsxError::MissingPart("xl/worksheets/sheet1.xml".to_string()))?;
            parse_sheet(BufReader::new(sheet_file), styles.clone())?
        };

        Ok(Self {
            sheets: vec![worksheet],
        })
    }

    pub fn sheet(&self, index: usize) -> Option<&Worksheet> {
        self.sheets.get(index)
    }
}

fn parse_sheet<R: BufRead>(reader: R, styles: Arc<StyleSheet>) -> Result<Worksheet, OxlsxError> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut text = String::new();
    let mut in_value = false;
    let mut current: Option<CellBuilder> = None;
    let mut cells: HashMap<String, Cell> = HashMap::new();

    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"c" => {
                let reference = attr_value(&e, b"r")?.unwrap_or_default();
                let style_id = attr_value(&e, b"s")?
                    .and_then(|s| s.parse::<u32>().ok())
                    .unwrap_or(0);
                let cell_type = attr_value(&e, b"t")?.unwrap_or_else(|| "n".to_string());

                current = Some(CellBuilder {
                    reference,
                    style_id,
                    cell_type,
                    value: Value::Empty,
                });
            }
            Event::Start(e) if e.local_name().as_ref() == b"v" => {
                in_value = true;
                text.clear();
            }
            Event::Text(t) if in_value => {
                text.push_str(&t.unescape()?.into_owned());
            }
            Event::End(e) if e.local_name().as_ref() == b"v" => {
                in_value = false;
                if let Some(builder) = current.as_mut() {
                    builder.value = match builder.cell_type.as_str() {
                        "s" => text
                            .parse::<u32>()
                            .map(Value::SharedStringIdx)
                            .unwrap_or(Value::Empty),
                        _ => text
                            .parse::<f64>()
                            .map(Value::Number)
                            .unwrap_or(Value::Empty),
                    };
                }
            }
            Event::End(e) if e.local_name().as_ref() == b"c" => {
                if let Some(builder) = current.take() {
                    if !builder.reference.is_empty() {
                        let reference = builder.reference.clone();
                        cells.insert(
                            reference,
                            Cell {
                                reference: builder.reference,
                                value: builder.value,
                                style_id: builder.style_id,
                                styles: styles.clone(),
                            },
                        );
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(Worksheet { cells })
}

#[derive(Debug)]
struct CellBuilder {
    reference: String,
    style_id: u32,
    cell_type: String,
    value: Value,
}

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Result<Option<String>, OxlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|err| OxlsxError::Xml(err.into()))?;
        if attr.key.as_ref() == key {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }

    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::{parse_sheet, Workbook, Value};
    use crate::{OxlsxError, StyleSheet};
    use std::fs;
    use std::io::Cursor;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};
    use std::sync::Arc;

    const SAMPLE_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="s"><v>0</v></c>
      <c r="B1" s="2" t="s"><v>0</v></c>
      <c r="C1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    const SAMPLE_STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11.000000"/><color theme="1"/><name val="Calibri"/></font>
    <font><b/><sz val="11.000000"/><color theme="1"/><name val="Calibri"/></font>
    <font><sz val="11.000000"/><color indexed="4"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor indexed="2"/><bgColor indexed="2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor indexed="5"/><bgColor indexed="5"/></patternFill></fill>
  </fills>
  <cellXfs count="3">
    <xf fontId="0" fillId="0"/>
    <xf fontId="1" fillId="2"/>
    <xf fontId="2" fillId="3"/>
  </cellXfs>
</styleSheet>"#;

    #[test]
    fn parses_sheet_cells_style_ids_and_values() {
        let styles = Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let ws = parse_sheet(Cursor::new(SAMPLE_SHEET_XML.as_bytes()), styles).unwrap();

        let a1 = ws.cell("A1").unwrap();
        assert_eq!(a1.style_id(), 1);
        assert_eq!(a1.value(), &Value::SharedStringIdx(0));

        let b1 = ws.cell("B1").unwrap();
        assert_eq!(b1.style_id(), 2);
        assert_eq!(b1.value(), &Value::SharedStringIdx(0));

        let c1 = ws.cell("C1").unwrap();
        assert_eq!(c1.style_id(), 0);
        assert_eq!(c1.value(), &Value::Number(42.0));

        assert!(ws.cell("D1").is_none());
    }

    #[test]
    fn style_accessors_use_stylesheet_via_style_id() {
        let styles = Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let ws = parse_sheet(Cursor::new(SAMPLE_SHEET_XML.as_bytes()), styles).unwrap();

        let a1 = ws.cell("A1").unwrap();
        assert!(a1.font().unwrap().bold);
        assert_eq!(a1.fill().unwrap().pattern_type, "solid");
    }

    #[test]
    fn cell_value_accessor_returns_parsed_value() {
        let styles = Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let ws = parse_sheet(Cursor::new(SAMPLE_SHEET_XML.as_bytes()), styles).unwrap();

        assert_eq!(ws.cell("A1").unwrap().value(), &Value::SharedStringIdx(0));
        assert_eq!(ws.cell("C1").unwrap().value(), &Value::Number(42.0));
    }

    #[test]
    fn workbook_open_missing_file_returns_io_error() {
        let err = Workbook::open("this-file-does-not-exist.xlsx").unwrap_err();
        assert!(matches!(err, OxlsxError::Io(_)));
    }

    #[test]
    fn workbook_open_invalid_zip_returns_invalid_format() {
        let mut path = PathBuf::from(std::env::temp_dir());
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        path.push(format!("oxlsx-invalid-{nanos}.xlsx"));

        fs::write(&path, b"not-a-zip").unwrap();
        let err = Workbook::open(&path).unwrap_err();
        fs::remove_file(&path).unwrap();

        assert!(matches!(err, OxlsxError::InvalidFormat(_)));
    }
}
