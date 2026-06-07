use crate::relationships::{SheetRelationship, parse_sheet_rels, sheet_rels_path};
use crate::workbook_rels::parse_sheet_refs;
use crate::{
    Fill, Font, Hyperlink, OxlsxError, ReadonlyWorkbook, SharedStrings, StyleId, StyleSheet,
};
use chrono::{Duration, NaiveDate};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufRead, BufReader, Cursor, Read};
use std::mem;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use zip::result::ZipError;

#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Empty,
    String(String),
    Number(f64),
    Bool(bool),
    Date(NaiveDate),
    Formula(String),
}

impl From<&str> for Value {
    fn from(s: &str) -> Self {
        Value::String(s.to_string())
    }
}

impl From<String> for Value {
    fn from(s: String) -> Self {
        Value::String(s)
    }
}

impl From<f64> for Value {
    fn from(n: f64) -> Self {
        Value::Number(n)
    }
}

impl From<i64> for Value {
    fn from(n: i64) -> Self {
        Value::Number(n as f64)
    }
}

impl From<bool> for Value {
    fn from(b: bool) -> Self {
        Value::Bool(b)
    }
}

impl From<NaiveDate> for Value {
    fn from(d: NaiveDate) -> Self {
        Value::Date(d)
    }
}

#[derive(Debug, Clone)]
pub struct Cell {
    reference: String,
    value: Value,
    style_id: StyleId,
    styles: Arc<StyleSheet>,
    pub(crate) write_font: Option<Font>,
    pub(crate) write_fill: Option<Fill>,
    pub(crate) write_hyperlink: Option<Hyperlink>,
}

impl Cell {
    pub fn reference(&self) -> &str {
        &self.reference
    }

    pub fn value(&self) -> &Value {
        &self.value
    }

    pub fn style_id(&self) -> StyleId {
        self.style_id
    }

    pub fn font(&self) -> Option<&Font> {
        self.styles.font(self.style_id)
    }

    pub fn fill(&self) -> Option<&Fill> {
        self.styles.fill(self.style_id)
    }

    pub fn text(&self) -> Option<&str> {
        match &self.value {
            Value::String(s) => Some(s.as_str()),
            _ => None,
        }
    }

    pub fn hyperlink(&self) -> Option<&Hyperlink> {
        self.write_hyperlink.as_ref()
    }

    pub(crate) fn parsed(
        reference: String,
        value: Value,
        style_id: StyleId,
        styles: Arc<StyleSheet>,
    ) -> Self {
        Self {
            reference,
            value,
            style_id,
            styles,
            write_font: None,
            write_fill: None,
            write_hyperlink: None,
        }
    }

    pub(crate) fn empty(reference: String, styles: Arc<StyleSheet>) -> Self {
        Self::parsed(reference, Value::Empty, StyleId::default(), styles)
    }
}

#[derive(Debug, Clone, Default)]
pub struct Worksheet {
    name: String,
    cells: HashMap<String, Cell>,
    pub(crate) merged_cells: Vec<String>,
    pub(crate) column_widths: HashMap<u32, f64>,
    pub(crate) row_heights: HashMap<u32, f64>,
    pub(crate) sheet_rels: Vec<SheetRelationship>,
}

impl Worksheet {
    pub(crate) fn new(name: &str) -> Self {
        Self {
            name: name.to_string(),
            cells: HashMap::new(),
            merged_cells: Vec::new(),
            column_widths: HashMap::new(),
            row_heights: HashMap::new(),
            sheet_rels: Vec::new(),
        }
    }

    pub fn title(&self) -> &str {
        &self.name
    }

    pub fn cell(&self, reference: &str) -> Option<&Cell> {
        self.cells.get(reference)
    }

    pub(crate) fn cells(&self) -> &HashMap<String, Cell> {
        &self.cells
    }

    pub fn set_cell(&mut self, reference: &str, value: impl Into<Value>) {
        let cell = self.ensure_cell_mut(reference);
        cell.value = value.into();
    }

    pub fn set_cell_font(&mut self, reference: &str, font: Font) {
        let cell = self.ensure_cell_mut(reference);
        cell.write_font = Some(font);
    }

    pub fn set_cell_fill(&mut self, reference: &str, fill: Fill) {
        let cell = self.ensure_cell_mut(reference);
        cell.write_fill = Some(fill);
    }

    pub fn set_cell_hyperlink(&mut self, reference: &str, hyperlink: impl Into<Hyperlink>) {
        let cell = self.ensure_cell_mut(reference);
        cell.write_hyperlink = Some(hyperlink.into());
    }

    pub fn merge_cells(&mut self, range: &str) {
        let range = range.to_string();
        if !self.merged_cells.contains(&range) {
            self.merged_cells.push(range);
        }
    }

    pub fn unmerge_cells(&mut self, range: &str) {
        self.merged_cells.retain(|existing| existing != range);
    }

    pub fn column_width(&mut self, col: u32, width: f64) {
        self.column_widths.insert(col, width);
    }

    pub fn row_height(&mut self, row: u32, height: f64) {
        self.row_heights.insert(row, height);
    }

    fn ensure_cell_mut(&mut self, reference: &str) -> &mut Cell {
        self.cells
            .entry(reference.to_string())
            .or_insert_with(|| Cell {
                reference: reference.to_string(),
                value: Value::Empty,
                style_id: StyleId::default(),
                styles: Arc::new(StyleSheet::default()),
                write_font: None,
                write_fill: None,
                write_hyperlink: None,
            })
    }
}

#[derive(Debug, Clone)]
pub struct Workbook {
    sheets: Vec<Worksheet>,
    active_index: usize,
    pub(crate) source_path: Option<PathBuf>,
    pub(crate) original_styles: Option<Arc<StyleSheet>>,
}

impl Workbook {
    pub fn new() -> Self {
        Self {
            sheets: vec![Worksheet::new("Sheet")],
            active_index: 0,
            source_path: None,
            original_styles: None,
        }
    }

    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self, OxlsxError> {
        let source_path = path.as_ref().to_path_buf();
        let file = File::open(&source_path)?;
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

        let shared_strings: Vec<String> = match archive.by_name("xl/sharedStrings.xml") {
            Ok(f) => SharedStrings::parse(BufReader::new(f))?.into_inner(),
            Err(ZipError::FileNotFound) => Vec::new(),
            Err(e) => return Err(OxlsxError::Zip(e)),
        };

        let (sheet_refs, active_index) = {
            let workbook_xml = {
                let mut wb_file = archive
                    .by_name("xl/workbook.xml")
                    .map_err(|_| OxlsxError::MissingPart("xl/workbook.xml".to_string()))?;
                let mut bytes = Vec::new();
                wb_file.read_to_end(&mut bytes)?;
                bytes
            };

            let rels_xml = {
                let mut rels_file =
                    archive.by_name("xl/_rels/workbook.xml.rels").map_err(|_| {
                        OxlsxError::MissingPart("xl/_rels/workbook.xml.rels".to_string())
                    })?;
                let mut bytes = Vec::new();
                rels_file.read_to_end(&mut bytes)?;
                bytes
            };

            parse_sheet_refs(
                BufReader::new(Cursor::new(workbook_xml)),
                BufReader::new(Cursor::new(rels_xml)),
            )?
        };

        let mut sheets = Vec::with_capacity(sheet_refs.len());
        for sheet_ref in &sheet_refs {
            let sheet_file = archive
                .by_name(&sheet_ref.zip_path)
                .map_err(|_| OxlsxError::MissingPart(sheet_ref.zip_path.clone()))?;

            let ParsedSheet {
                mut worksheet,
                hyperlink_refs,
            } = parse_sheet_with_hyperlinks(
                BufReader::new(sheet_file),
                styles.clone(),
                &shared_strings,
            )?;

            let rels = match archive.by_name(&sheet_rels_path(&sheet_ref.zip_path)) {
                Ok(rels_file) => parse_sheet_rels(BufReader::new(rels_file))?,
                Err(ZipError::FileNotFound) => Vec::new(),
                Err(err) => return Err(OxlsxError::Zip(err)),
            };

            for (cell_ref, rel_id) in hyperlink_refs {
                if let Some(target) = SheetRelationship::hyperlink_target_by_id(&rels, &rel_id) {
                    worksheet.set_cell_hyperlink(&cell_ref, Hyperlink::new(target));
                }
            }

            worksheet.sheet_rels = rels
                .into_iter()
                .filter(|rel| matches!(rel, SheetRelationship::Unknown { .. }))
                .collect();

            worksheet.name = sheet_ref.name.clone();
            sheets.push(worksheet);
        }

        Ok(Self {
            sheets,
            active_index,
            source_path: Some(source_path),
            original_styles: Some(styles),
        })
    }

    pub fn open_readonly<P: AsRef<Path>>(path: P) -> Result<ReadonlyWorkbook, OxlsxError> {
        crate::readonly::open_readonly(path)
    }

    pub fn worksheets(&self) -> &[Worksheet] {
        &self.sheets
    }

    pub fn sheetnames(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.title()).collect()
    }

    pub fn sheet_by_name(&self, name: &str) -> Option<&Worksheet> {
        self.sheets.iter().find(|s| s.title() == name)
    }

    pub fn active(&self) -> Option<&Worksheet> {
        self.sheets.get(self.active_index)
    }

    pub fn active_mut(&mut self) -> Option<&mut Worksheet> {
        self.sheets.get_mut(self.active_index)
    }

    pub fn sheet(&self, index: usize) -> Option<&Worksheet> {
        self.sheets.get(index)
    }

    pub fn sheet_mut(&mut self, index: usize) -> Option<&mut Worksheet> {
        self.sheets.get_mut(index)
    }

    pub fn create_sheet(&mut self, name: &str) -> &mut Worksheet {
        self.sheets.push(Worksheet::new(name));
        self.sheets
            .last_mut()
            .expect("create_sheet always pushes one worksheet")
    }

    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<(), OxlsxError> {
        crate::writer::write_xlsx(self, path.as_ref(), self.source_path.as_deref())
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
fn parse_sheet<R: BufRead>(
    reader: R,
    styles: Arc<StyleSheet>,
    shared_strings: &[String],
) -> Result<Worksheet, OxlsxError> {
    Ok(parse_sheet_with_hyperlinks(reader, styles, shared_strings)?.worksheet)
}

struct ParsedSheet {
    worksheet: Worksheet,
    hyperlink_refs: HashMap<String, String>,
}

fn parse_sheet_with_hyperlinks<R: BufRead>(
    reader: R,
    styles: Arc<StyleSheet>,
    shared_strings: &[String],
) -> Result<ParsedSheet, OxlsxError> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut decoder = CellDecoder::new(shared_strings);
    let mut cells: HashMap<String, Cell> = HashMap::new();
    let mut merged_cells = Vec::new();
    let mut column_widths = HashMap::new();
    let mut row_heights = HashMap::new();
    let mut hyperlink_refs = HashMap::new();

    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"c" => {
                decoder.start_cell(&e)?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"row" => {
                parse_row_height(&e, &mut row_heights)?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"col" => {
                parse_column_widths(&e, &mut column_widths)?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"mergeCell" => {
                if let Some(range_ref) = attr_value(&e, b"ref")? {
                    merged_cells.push(range_ref);
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"hyperlink" => {
                collect_hyperlink_ref(&e, &mut hyperlink_refs)?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"v" => {
                decoder.start_value();
            }
            Event::Start(e) if e.local_name().as_ref() == b"f" => {
                decoder.start_formula();
            }
            Event::Start(e) if decoder.in_inline_str() && e.local_name().as_ref() == b"is" => {
                decoder.start_inline_str();
            }
            Event::Start(e) if decoder.in_inline_str() && e.local_name().as_ref() == b"t" => {
                decoder.start_inline_text();
            }
            Event::Empty(e) if decoder.in_inline_str() && e.local_name().as_ref() == b"t" => {
                decoder.end_inline_text();
            }
            Event::Empty(e) if e.local_name().as_ref() == b"row" => {
                parse_row_height(&e, &mut row_heights)?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"col" => {
                parse_column_widths(&e, &mut column_widths)?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"mergeCell" => {
                if let Some(range_ref) = attr_value(&e, b"ref")? {
                    merged_cells.push(range_ref);
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"hyperlink" => {
                collect_hyperlink_ref(&e, &mut hyperlink_refs)?;
            }
            Event::Text(t) if decoder.in_value() => {
                decoder.push_value_text(&t.unescape()?);
            }
            Event::Text(t) if decoder.in_formula() => {
                decoder.push_formula_text(&t.unescape()?);
            }
            Event::Text(t) if decoder.in_inline_text() => {
                decoder.push_inline_text(&t.unescape()?);
            }
            Event::End(e) if e.local_name().as_ref() == b"v" => {
                decoder.end_value(&styles);
            }
            Event::End(e) if e.local_name().as_ref() == b"f" => {
                decoder.end_formula();
            }
            Event::End(e) if e.local_name().as_ref() == b"t" => {
                decoder.end_inline_text();
            }
            Event::End(e) if e.local_name().as_ref() == b"is" => {
                decoder.end_inline_str();
            }
            Event::End(e) if e.local_name().as_ref() == b"c" => {
                if let Some(cell) = decoder.finish_cell(styles.clone()) {
                    cells.insert(cell.reference().to_string(), cell);
                }
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(ParsedSheet {
        worksheet: Worksheet {
            name: String::new(),
            cells,
            merged_cells,
            column_widths,
            row_heights,
            sheet_rels: Vec::new(),
        },
        hyperlink_refs,
    })
}

fn collect_hyperlink_ref(
    e: &BytesStart<'_>,
    hyperlink_refs: &mut HashMap<String, String>,
) -> Result<(), OxlsxError> {
    let cell_ref = attr_value(e, b"ref")?;
    let rel_id = attr_value(e, b"r:id")?.or(attr_value(e, b"id")?);
    if let (Some(cell_ref), Some(rel_id)) = (cell_ref, rel_id) {
        hyperlink_refs.insert(cell_ref, rel_id);
    }

    Ok(())
}

#[derive(Debug, Clone)]
struct CellBuilder {
    reference: String,
    style_id: StyleId,
    cell_type: String,
    value: Value,
    formula_text: String,
}

#[derive(Debug, Clone)]
pub(crate) struct CellDecoder {
    shared_strings: Arc<Vec<String>>,
    value_text: String,
    inline_text: String,
    in_value: bool,
    in_formula: bool,
    in_inline_str: bool,
    in_inline_text: bool,
    current: Option<CellBuilder>,
}

impl CellDecoder {
    pub(crate) fn new(shared_strings: &[String]) -> Self {
        Self {
            shared_strings: Arc::new(shared_strings.to_vec()),
            value_text: String::new(),
            inline_text: String::new(),
            in_value: false,
            in_formula: false,
            in_inline_str: false,
            in_inline_text: false,
            current: None,
        }
    }

    pub(crate) fn new_arc(shared_strings: Arc<Vec<String>>) -> Self {
        Self {
            shared_strings,
            value_text: String::new(),
            inline_text: String::new(),
            in_value: false,
            in_formula: false,
            in_inline_str: false,
            in_inline_text: false,
            current: None,
        }
    }

    pub(crate) fn start_cell(&mut self, cell_start: &BytesStart<'_>) -> Result<(), OxlsxError> {
        let reference = attr_value(cell_start, b"r")?.unwrap_or_default();
        let style_id = attr_value(cell_start, b"s")?
            .and_then(|s| s.parse::<usize>().ok())
            .map(StyleId)
            .unwrap_or_default();
        let cell_type = attr_value(cell_start, b"t")?.unwrap_or_else(|| "n".to_string());

        self.current = Some(CellBuilder {
            reference,
            style_id,
            cell_type: cell_type.clone(),
            value: Value::Empty,
            formula_text: String::new(),
        });
        self.in_inline_str = cell_type == "inlineStr";
        self.in_inline_text = false;

        Ok(())
    }

    pub(crate) fn start_value(&mut self) {
        self.in_value = true;
        self.value_text.clear();
    }

    pub(crate) fn push_value_text(&mut self, text: &str) {
        if self.in_value {
            self.value_text.push_str(text);
        }
    }

    pub(crate) fn end_value(&mut self, styles: &StyleSheet) {
        self.in_value = false;

        if let Some(builder) = self.current.as_mut() {
            if matches!(builder.value, Value::Formula(_)) {
                return;
            }

            let mut value = match builder.cell_type.as_str() {
                "s" => self
                    .value_text
                    .parse::<usize>()
                    .ok()
                    .and_then(|i| self.shared_strings.get(i))
                    .map(|s| Value::String(s.to_string()))
                    .unwrap_or(Value::Empty),
                "b" => Value::Bool(self.value_text == "1"),
                _ => self
                    .value_text
                    .parse::<f64>()
                    .map(Value::Number)
                    .unwrap_or(Value::Empty),
            };

            if let Value::Number(n) = value
                && styles.is_date_format(builder.style_id)
            {
                value = excel_serial_to_date(n)
                    .map(Value::Date)
                    .unwrap_or(Value::Number(n));
            }

            builder.value = value;
        }
    }

    pub(crate) fn start_inline_str(&mut self) {
        if self.in_inline_str {
            self.inline_text.clear();
        }
    }

    pub(crate) fn start_formula(&mut self) {
        if let Some(builder) = self.current.as_mut() {
            builder.formula_text.clear();
            self.in_formula = true;
        }
    }

    pub(crate) fn push_formula_text(&mut self, text: &str) {
        if self.in_formula
            && let Some(builder) = self.current.as_mut()
        {
            builder.formula_text.push_str(text);
        }
    }

    pub(crate) fn end_formula(&mut self) {
        self.in_formula = false;
        if let Some(builder) = self.current.as_mut()
            && !builder.formula_text.is_empty()
        {
            builder.value = Value::Formula(mem::take(&mut builder.formula_text));
        }
    }

    pub(crate) fn start_inline_text(&mut self) {
        if self.in_inline_str {
            self.in_inline_text = true;
        }
    }

    pub(crate) fn push_inline_text(&mut self, text: &str) {
        if self.in_inline_text {
            self.inline_text.push_str(text);
        }
    }

    pub(crate) fn end_inline_text(&mut self) {
        self.in_inline_text = false;
    }

    pub(crate) fn end_inline_str(&mut self) {
        if let Some(builder) = self.current.as_mut() {
            builder.value = Value::String(mem::take(&mut self.inline_text));
        }
    }

    pub(crate) fn finish_cell(&mut self, styles: Arc<StyleSheet>) -> Option<Cell> {
        self.in_inline_str = false;
        self.in_inline_text = false;
        self.in_formula = false;

        let builder = self.current.take()?;
        if builder.reference.is_empty() {
            return None;
        }

        Some(Cell::parsed(
            builder.reference,
            builder.value,
            builder.style_id,
            styles,
        ))
    }

    pub(crate) fn in_value(&self) -> bool {
        self.in_value
    }

    pub(crate) fn in_inline_str(&self) -> bool {
        self.in_inline_str
    }

    pub(crate) fn in_formula(&self) -> bool {
        self.in_formula
    }

    pub(crate) fn in_inline_text(&self) -> bool {
        self.in_inline_text
    }
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

fn parse_column_widths(
    e: &BytesStart<'_>,
    column_widths: &mut HashMap<u32, f64>,
) -> Result<(), OxlsxError> {
    let width = attr_value(e, b"width")?.and_then(|v| v.parse::<f64>().ok());
    let min = attr_value(e, b"min")?.and_then(|v| v.parse::<u32>().ok());
    let max = attr_value(e, b"max")?.and_then(|v| v.parse::<u32>().ok());

    if let (Some(width), Some(min), Some(max)) = (width, min, max) {
        for col in min..=max {
            column_widths.insert(col, width);
        }
    }

    Ok(())
}

fn parse_row_height(
    e: &BytesStart<'_>,
    row_heights: &mut HashMap<u32, f64>,
) -> Result<(), OxlsxError> {
    let row = attr_value(e, b"r")?.and_then(|v| v.parse::<u32>().ok());
    let height = attr_value(e, b"ht")?.and_then(|v| v.parse::<f64>().ok());

    if let (Some(row), Some(height)) = (row, height) {
        row_heights.insert(row, height);
    }

    Ok(())
}

fn excel_serial_to_date(serial: f64) -> Option<NaiveDate> {
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?;
    let days = serial as i64;
    epoch.checked_add_signed(Duration::days(days))
}

#[cfg(test)]
mod tests {
    use super::{CellDecoder, Value, Workbook, Worksheet, parse_sheet};
    use crate::{Fill, Font, OxlsxError, StyleId, StyleSheet};
    use chrono::NaiveDate;
    use quick_xml::events::BytesStart;
    use std::collections::HashMap;
    use std::fs;
    use std::io::Cursor;
    use std::path::PathBuf;
    use std::sync::Arc;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn tmp_path(name: &str) -> PathBuf {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system clock must be after unix epoch")
            .as_nanos();
        std::env::temp_dir().join(format!("{stamp}_{name}"))
    }

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

    const INLINE_STR_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>hello</t></is></c>
      <c r="B1" t="b"><v>1</v></c>
      <c r="C1" t="b"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    const DATE_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="B1" s="2"><v>45678</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    const NON_DATE_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="C1" s="0"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    const HYPERLINK_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="A1" r:id="rId9"/>
  </hyperlinks>
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
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>
  </numFmts>
  <cellXfs count="3">
    <xf fontId="0" fillId="0" numFmtId="0"/>
    <xf fontId="1" fillId="2" numFmtId="14"/>
    <xf fontId="2" fillId="3" numFmtId="164"/>
  </cellXfs>
</styleSheet>"#;

    #[test]
    fn parses_sheet_cells_style_ids_and_values() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let ws = parse_sheet(
            Cursor::new(SAMPLE_SHEET_XML.as_bytes()),
            styles,
            &["Hola".to_string()],
        )
        .unwrap();

        let a1 = ws.cell("A1").unwrap();
        assert_eq!(a1.style_id(), StyleId(1));
        assert_eq!(a1.value(), &Value::String("Hola".to_string()));

        let b1 = ws.cell("B1").unwrap();
        assert_eq!(b1.style_id(), StyleId(2));
        assert_eq!(b1.value(), &Value::String("Hola".to_string()));

        let c1 = ws.cell("C1").unwrap();
        assert_eq!(c1.style_id(), StyleId(0));
        assert_eq!(c1.value(), &Value::Number(42.0));

        assert!(ws.cell("D1").is_none());
    }

    #[test]
    fn style_accessors_use_stylesheet_via_style_id() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let ws = parse_sheet(
            Cursor::new(SAMPLE_SHEET_XML.as_bytes()),
            styles,
            &["Hola".to_string()],
        )
        .unwrap();

        let a1 = ws.cell("A1").unwrap();
        assert!(a1.font().unwrap().bold);
        assert_eq!(a1.fill().unwrap().pattern_type, "solid");
    }

    #[test]
    fn cell_value_accessor_returns_parsed_value() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let ws = parse_sheet(
            Cursor::new(SAMPLE_SHEET_XML.as_bytes()),
            styles,
            &["Hola".to_string()],
        )
        .unwrap();

        assert_eq!(
            ws.cell("A1").unwrap().value(),
            &Value::String("Hola".to_string())
        );
        assert_eq!(ws.cell("C1").unwrap().value(), &Value::Number(42.0));
    }

    #[test]
    fn value_supports_string_and_bool_variants() {
        let text = Value::String("Hola".to_string());
        let boolean = Value::Bool(true);

        assert_eq!(text, Value::String("Hola".to_string()));
        assert_eq!(boolean, Value::Bool(true));
    }

    #[test]
    fn value_from_impls_convert_supported_types() {
        assert_eq!(Value::from("Hola"), Value::String("Hola".to_string()));
        assert_eq!(
            Value::from(String::from("Chau")),
            Value::String("Chau".to_string())
        );
        assert_eq!(Value::from(42.5_f64), Value::Number(42.5));
        assert_eq!(Value::from(7_i64), Value::Number(7.0));
        assert_eq!(Value::from(true), Value::Bool(true));
        assert_eq!(
            Value::from(NaiveDate::from_ymd_opt(2025, 1, 21).unwrap()),
            Value::Date(NaiveDate::from_ymd_opt(2025, 1, 21).unwrap())
        );
    }

    #[test]
    fn parses_inline_string_and_boolean_cells() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let ws = parse_sheet(
            Cursor::new(INLINE_STR_SHEET_XML.as_bytes()),
            styles,
            &["Hola".to_string()],
        )
        .unwrap();

        assert_eq!(
            ws.cell("A1").unwrap().value(),
            &Value::String("hello".into())
        );
        assert_eq!(ws.cell("B1").unwrap().value(), &Value::Bool(true));
        assert_eq!(ws.cell("C1").unwrap().value(), &Value::Bool(false));
    }

    #[test]
    fn cell_text_returns_only_string_values() {
        let styles = Arc::new(StyleSheet::default());

        let string_cell = super::Cell {
            reference: "A1".to_string(),
            value: Value::String("Hola".to_string()),
            style_id: StyleId(0),
            styles: styles.clone(),
            write_font: None,
            write_fill: None,
            write_hyperlink: None,
        };
        assert_eq!(string_cell.text(), Some("Hola"));

        let number_cell = super::Cell {
            reference: "B1".to_string(),
            value: Value::Number(42.0),
            style_id: StyleId(0),
            styles: styles.clone(),
            write_font: None,
            write_fill: None,
            write_hyperlink: None,
        };
        assert_eq!(number_cell.text(), None);

        let bool_cell = super::Cell {
            reference: "C1".to_string(),
            value: Value::Bool(true),
            style_id: StyleId(0),
            styles: styles.clone(),
            write_font: None,
            write_fill: None,
            write_hyperlink: None,
        };
        assert_eq!(bool_cell.text(), None);

        let empty_cell = super::Cell {
            reference: "D1".to_string(),
            value: Value::Empty,
            style_id: StyleId(0),
            styles,
            write_font: None,
            write_fill: None,
            write_hyperlink: None,
        };
        assert_eq!(empty_cell.text(), None);
    }

    #[test]
    fn workbook_open_missing_file_returns_io_error() {
        let err = Workbook::open("this-file-does-not-exist.xlsx").unwrap_err();
        assert!(matches!(err, OxlsxError::Io(_)));
    }

    #[test]
    fn workbook_open_invalid_zip_returns_invalid_format() {
        let mut path = std::env::temp_dir();
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

    #[test]
    fn excel_serial_to_date_maps_known_serials() {
        assert_eq!(
            super::excel_serial_to_date(45678.0),
            NaiveDate::from_ymd_opt(2025, 1, 21)
        );
        assert_eq!(
            super::excel_serial_to_date(1.0),
            NaiveDate::from_ymd_opt(1899, 12, 31)
        );
    }

    #[test]
    fn parse_sheet_converts_numeric_date_by_style() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let ws = parse_sheet(Cursor::new(DATE_SHEET_XML.as_bytes()), styles, &[]).unwrap();

        assert_eq!(
            ws.cell("B1").unwrap().value(),
            &Value::Date(NaiveDate::from_ymd_opt(2025, 1, 21).unwrap())
        );
    }

    #[test]
    fn parse_sheet_keeps_non_date_numbers_unchanged() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let ws = parse_sheet(Cursor::new(NON_DATE_SHEET_XML.as_bytes()), styles, &[]).unwrap();

        assert_eq!(ws.cell("C1").unwrap().value(), &Value::Number(42.0));
    }

    #[test]
    fn parse_sheet_collects_hyperlink_relationship_references() {
        let styles = Arc::new(StyleSheet::default());
        let parsed = super::parse_sheet_with_hyperlinks(
            Cursor::new(HYPERLINK_SHEET_XML.as_bytes()),
            styles,
            &["docs".to_string()],
        )
        .unwrap();

        assert_eq!(
            parsed.hyperlink_refs.get("A1").map(String::as_str),
            Some("rId9")
        );
    }

    #[test]
    fn worksheet_title_and_workbook_multi_sheet_accessors() {
        let wb = Workbook {
            sheets: vec![
                Worksheet {
                    name: "Empleados".to_string(),
                    cells: HashMap::new(),
                    merged_cells: Vec::new(),
                    column_widths: HashMap::new(),
                    row_heights: HashMap::new(),
                    sheet_rels: Vec::new(),
                },
                Worksheet {
                    name: "Productos".to_string(),
                    cells: HashMap::new(),
                    merged_cells: Vec::new(),
                    column_widths: HashMap::new(),
                    row_heights: HashMap::new(),
                    sheet_rels: Vec::new(),
                },
            ],
            active_index: 1,
            source_path: None,
            original_styles: None,
        };

        assert_eq!(wb.sheet(0).unwrap().title(), "Empleados");
        assert_eq!(wb.sheet(1).unwrap().title(), "Productos");
        assert_eq!(wb.sheetnames(), vec!["Empleados", "Productos"]);
        assert!(wb.sheet_by_name("Empleados").is_some());
        assert!(wb.sheet_by_name("NoExiste").is_none());
        assert_eq!(wb.active().unwrap().title(), "Productos");
    }

    #[test]
    fn workbook_new_creates_default_sheet() {
        let wb = Workbook::new();

        assert_eq!(wb.sheet(0).unwrap().title(), "Sheet");
    }

    #[test]
    fn workbook_new_leaves_source_path_empty() {
        let wb = Workbook::new();

        assert!(wb.source_path.is_none());
    }

    #[test]
    fn workbook_open_tracks_source_path() {
        let wb = Workbook::open("samples/test.xlsx").unwrap();

        assert_eq!(
            wb.source_path.as_deref(),
            Some(std::path::Path::new("samples/test.xlsx"))
        );
    }

    #[test]
    fn workbook_create_sheet_appends_new_sheet() {
        let mut wb = Workbook::new();

        wb.create_sheet("Data");

        assert_eq!(wb.sheetnames(), vec!["Sheet", "Data"]);
    }

    #[test]
    fn workbook_active_mut_returns_some_for_default_workbook() {
        let mut wb = Workbook::new();

        assert!(wb.active_mut().is_some());
    }

    #[test]
    fn workbook_save_writes_xlsx_file() {
        let mut wb = Workbook::new();
        wb.active_mut().unwrap().set_cell("A1", "saved");
        let path = tmp_path("workbook_save_writes_xlsx_file.xlsx");

        wb.save(&path).unwrap();

        assert!(path.exists());
        fs::remove_file(path).ok();
    }

    #[test]
    fn worksheet_set_cell_sets_string_number_and_bool_values() {
        let mut ws = Worksheet::new("Data");

        ws.set_cell("A1", "Hello");
        assert_eq!(
            ws.cell("A1").unwrap().value(),
            &Value::String("Hello".to_string())
        );

        ws.set_cell("A1", 42.0_f64);
        assert_eq!(ws.cell("A1").unwrap().value(), &Value::Number(42.0));

        ws.set_cell("A1", true);
        assert_eq!(ws.cell("A1").unwrap().value(), &Value::Bool(true));
    }

    #[test]
    fn worksheet_set_cell_font_stores_write_font_and_creates_missing_cell() {
        let mut ws = Worksheet::new("Data");

        ws.set_cell_font(
            "A1",
            Font {
                bold: true,
                ..Default::default()
            },
        );

        let cell = ws.cell("A1").unwrap();
        assert!(cell.write_font.is_some());
    }

    #[test]
    fn worksheet_set_cell_fill_stores_write_fill() {
        let mut ws = Worksheet::new("Data");

        ws.set_cell_fill(
            "A1",
            Fill {
                pattern_type: "solid".to_string(),
                ..Default::default()
            },
        );

        let cell = ws.cell("A1").unwrap();
        assert!(cell.write_fill.is_some());
    }

    #[test]
    fn workbook_open_loads_sheet_refs_from_workbook_metadata() {
        let wb = Workbook::open("samples/multi_sheet.xlsx").unwrap();

        assert_eq!(wb.worksheets().len(), 2);
        assert_eq!(wb.sheetnames(), vec!["Empleados", "Productos"]);
        assert_eq!(wb.active().unwrap().title(), "Empleados");
    }

    #[test]
    fn cell_decoder_decodes_shared_string_cells() {
        let shared_strings = ["Hola".to_string()];
        let styles = Arc::new(StyleSheet::default());
        let mut decoder = CellDecoder::new(&shared_strings);

        let mut cell_start = BytesStart::new("c");
        cell_start.push_attribute(("r", "A1"));
        cell_start.push_attribute(("s", "0"));
        cell_start.push_attribute(("t", "s"));

        decoder.start_cell(&cell_start).unwrap();
        decoder.start_value();
        decoder.push_value_text("0");
        decoder.end_value(&styles);

        let cell = decoder.finish_cell(styles).expect("cell should be decoded");
        assert_eq!(cell.reference(), "A1");
        assert_eq!(cell.style_id(), StyleId(0));
        assert_eq!(cell.value(), &Value::String("Hola".to_string()));
    }
}
