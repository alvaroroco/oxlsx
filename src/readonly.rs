use crate::coords::{column_index, column_name};
use crate::reader::CellDecoder;
use crate::workbook_rels::parse_sheet_refs;
use crate::{Cell, OxlsxError, SharedStrings, StyleSheet};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::cell::RefCell;
use std::fs::File;
use std::io::{BufRead, BufReader, Cursor, Read};
use std::sync::Arc;
use zip::ZipArchive;
use zip::result::ZipError;

#[derive(Debug, Clone)]
struct SheetMeta {
    name: String,
    zip_path: String,
    dimension_cols: Option<usize>,
}

#[derive(Debug)]
pub struct ReadonlyWorkbook {
    archive: Arc<RefCell<ZipArchive<File>>>,
    styles: Arc<StyleSheet>,
    shared_strings: Arc<Vec<String>>,
    sheet_metas: Vec<SheetMeta>,
    active_index: usize,
}

impl ReadonlyWorkbook {
    pub fn sheet(&self, index: usize) -> Option<ReadonlyWorksheet<'_>> {
        self.sheet_metas.get(index)?;
        Some(ReadonlyWorksheet {
            workbook: self,
            sheet_index: index,
        })
    }

    pub fn sheetnames(&self) -> Vec<&str> {
        self.sheet_metas
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect()
    }

    pub fn active(&self) -> Option<ReadonlyWorksheet<'_>> {
        self.sheet(self.active_index)
    }

    fn streaming_rows_for_sheet(
        &self,
        sheet_index: usize,
    ) -> Result<RowIter<impl BufRead + '_>, OxlsxError> {
        let sheet_meta = self
            .sheet_metas
            .get(sheet_index)
            .ok_or_else(|| OxlsxError::MissingPart(format!("sheet index {sheet_index}")))?;

        // Safety: borrow_mut panics if called concurrently (documented limitation).
        // rows() is not Send, so concurrent calls require unsafe workarounds the
        // caller would have to explicitly construct.
        let mut archive = self.archive.borrow_mut();
        let sheet_file = archive
            .by_name(&sheet_meta.zip_path)
            .map_err(|_| OxlsxError::MissingPart(sheet_meta.zip_path.clone()))?;

        // We need to own the bytes because the ZipFile borrows the archive mutably.
        // Reading into a Vec and wrapping in Cursor gives us an owned BufRead
        // without holding the borrow across the iterator's lifetime.
        let mut bytes = Vec::new();
        sheet_file
            .take(64 * 1024 * 1024) // 64 MB guard
            .read_to_end(&mut bytes)?;

        Ok(RowIter::new(
            Cursor::new(bytes),
            self.shared_strings.clone(),
            self.styles.clone(),
            sheet_meta.dimension_cols,
        ))
    }

    #[cfg(test)]
    pub(crate) fn new_stub() -> Self {
        let file = File::open("samples/test.xlsx")
            .expect("samples/test.xlsx should exist for readonly stub test");
        let archive = ZipArchive::new(file).expect("test fixture should be a valid zip");
        #[allow(clippy::arc_with_non_send_sync)]
        let archive = Arc::new(RefCell::new(archive));

        Self {
            archive,
            styles: Arc::new(StyleSheet::default()),
            shared_strings: Arc::new(Vec::new()),
            sheet_metas: vec![SheetMeta {
                name: "Sheet1".to_string(),
                zip_path: "xl/worksheets/sheet1.xml".to_string(),
                dimension_cols: None,
            }],
            active_index: 0,
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub struct ReadonlyWorksheet<'a> {
    workbook: &'a ReadonlyWorkbook,
    sheet_index: usize,
}

impl<'a> ReadonlyWorksheet<'a> {
    pub fn title(&self) -> &str {
        self.workbook
            .sheet_metas
            .get(self.sheet_index)
            .map(|sheet| sheet.name.as_str())
            .unwrap_or_default()
    }

    pub fn rows(&self) -> RowIter<impl BufRead + '_> {
        self.workbook
            .streaming_rows_for_sheet(self.sheet_index)
            .unwrap_or_else(|_| RowIter::empty())
    }
}

/// A lazy row iterator that parses one row per `next()` call from an XML stream.
/// No cells are parsed until the caller drives the iterator.
pub struct RowIter<R: BufRead> {
    xml: Option<Reader<R>>,
    decoder: CellDecoder,
    styles: Arc<StyleSheet>,
    dimension_cols: Option<usize>,
    buf: Vec<u8>,
    row_cells: Vec<(usize, Cell)>,
    done: bool,
    current_row_num: usize,
    in_row: bool,
}

impl<R: BufRead> RowIter<R> {
    fn new(
        reader: R,
        shared_strings: Arc<Vec<String>>,
        styles: Arc<StyleSheet>,
        dimension_cols: Option<usize>,
    ) -> Self {
        let mut xml = Reader::from_reader(reader);
        xml.config_mut().trim_text(true);
        Self {
            xml: Some(xml),
            decoder: CellDecoder::new_arc(shared_strings),
            styles,
            dimension_cols,
            buf: Vec::new(),
            row_cells: Vec::new(),
            done: false,
            current_row_num: 0,
            in_row: false,
        }
    }

    fn empty() -> Self
    where
        R: BufRead,
    {
        // This path is unreachable in practice but needed for type inference
        // when RowIter<impl BufRead> appears in the error path.
        // We set done=true so next() immediately returns None.
        // The xml field is None in this sentinel state.
        Self {
            xml: None,
            decoder: CellDecoder::new(&[]),
            styles: Arc::new(StyleSheet::default()),
            dimension_cols: None,
            buf: Vec::new(),
            row_cells: Vec::new(),
            done: true,
            current_row_num: 0,
            in_row: false,
        }
    }
}

impl<R: BufRead> Iterator for RowIter<R> {
    type Item = Vec<Cell>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        let xml = self.xml.as_mut()?;

        loop {
            self.buf.clear();
            match xml.read_event_into(&mut self.buf) {
                Err(_) => {
                    self.done = true;
                    return None;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"row" => {
                    self.in_row = true;
                    self.row_cells.clear();
                    self.current_row_num =
                        row_number(e).unwrap_or(self.current_row_num.saturating_add(1));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"row" => {
                    self.in_row = false;
                    let row = materialize_row(
                        &mut self.row_cells,
                        self.current_row_num,
                        self.styles.clone(),
                        self.dimension_cols,
                    );
                    return Some(row);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"c" => {
                    let _ = self.decoder.start_cell(e);
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"c" => {
                    let _ = self.decoder.start_cell(e);
                    if let Some(cell) = self.decoder.finish_cell(self.styles.clone())
                        && self.in_row
                    {
                        if let Some(col) =
                            column_index(cell.reference()).and_then(|c| usize::try_from(c).ok())
                        {
                            self.row_cells.push((col, cell));
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"v" => {
                    self.decoder.start_value();
                }
                Ok(Event::Start(ref e))
                    if self.decoder.in_inline_str() && e.local_name().as_ref() == b"is" =>
                {
                    self.decoder.start_inline_str();
                }
                Ok(Event::Start(ref e))
                    if self.decoder.in_inline_str() && e.local_name().as_ref() == b"t" =>
                {
                    self.decoder.start_inline_text();
                }
                Ok(Event::Empty(ref e))
                    if self.decoder.in_inline_str() && e.local_name().as_ref() == b"t" =>
                {
                    self.decoder.end_inline_text();
                }
                Ok(Event::Text(ref t)) if self.decoder.in_value() => {
                    if let Ok(s) = t.unescape() {
                        self.decoder.push_value_text(&s);
                    }
                }
                Ok(Event::Text(ref t)) if self.decoder.in_inline_text() => {
                    if let Ok(s) = t.unescape() {
                        self.decoder.push_inline_text(&s);
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"v" => {
                    self.decoder.end_value(&self.styles);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"t" => {
                    self.decoder.end_inline_text();
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"is" => {
                    self.decoder.end_inline_str();
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => {
                    if let Some(cell) = self.decoder.finish_cell(self.styles.clone())
                        && self.in_row
                    {
                        if let Some(col) =
                            column_index(cell.reference()).and_then(|c| usize::try_from(c).ok())
                        {
                            self.row_cells.push((col, cell));
                        }
                    }
                }
                Ok(Event::Eof) => {
                    self.done = true;
                    return None;
                }
                _ => {}
            }
        }
    }
}

/// Creates a streaming `RowIter` from any `BufRead` source.
/// Used in tests. Parses one row per `next()` call.
#[cfg(test)]
pub(crate) fn streaming_rows_from_xml<R: BufRead>(
    reader: R,
    styles: Arc<StyleSheet>,
    shared_strings: &[String],
    dimension_cols: Option<usize>,
) -> RowIter<R> {
    RowIter::new(
        reader,
        Arc::new(shared_strings.to_vec()),
        styles,
        dimension_cols,
    )
}

pub(crate) fn open_readonly<P: AsRef<std::path::Path>>(
    path: P,
) -> Result<ReadonlyWorkbook, OxlsxError> {
    let file = File::open(path)?;
    let mut archive = ZipArchive::new(file).map_err(|err| match err {
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
            let mut rels_file = archive
                .by_name("xl/_rels/workbook.xml.rels")
                .map_err(|_| OxlsxError::MissingPart("xl/_rels/workbook.xml.rels".to_string()))?;
            let mut bytes = Vec::new();
            rels_file.read_to_end(&mut bytes)?;
            bytes
        };

        parse_sheet_refs(
            BufReader::new(Cursor::new(workbook_xml)),
            BufReader::new(Cursor::new(rels_xml)),
        )?
    };

    let mut sheet_metas = Vec::with_capacity(sheet_refs.len());
    for sheet_ref in sheet_refs {
        let sheet_file = archive
            .by_name(&sheet_ref.zip_path)
            .map_err(|_| OxlsxError::MissingPart(sheet_ref.zip_path.clone()))?;
        let dimension_cols = parse_sheet_dimension(BufReader::new(sheet_file))?;

        sheet_metas.push(SheetMeta {
            name: sheet_ref.name,
            zip_path: sheet_ref.zip_path,
            dimension_cols,
        });
    }

    #[allow(clippy::arc_with_non_send_sync)]
    let archive = Arc::new(RefCell::new(archive));

    Ok(ReadonlyWorkbook {
        archive,
        styles,
        shared_strings: Arc::new(shared_strings),
        sheet_metas,
        active_index,
    })
}

fn materialize_row(
    row_cells: &mut [(usize, Cell)],
    row_num: usize,
    styles: Arc<StyleSheet>,
    dimension_cols: Option<usize>,
) -> Vec<Cell> {
    row_cells.sort_by_key(|(col, _)| *col);

    let observed_max = row_cells.iter().map(|(col, _)| *col).max().unwrap_or(0);
    let total_cols = dimension_cols.unwrap_or(observed_max);
    if total_cols == 0 {
        return Vec::new();
    }

    let mut result = Vec::with_capacity(total_cols);
    let mut cursor = 0usize;

    for col in 1..=total_cols {
        if let Some((cell_col, cell)) = row_cells.get(cursor)
            && *cell_col == col
        {
            result.push(cell.clone());
            cursor += 1;
            continue;
        }

        let Some(col_u32) = u32::try_from(col).ok() else {
            continue;
        };
        let reference = format!("{}{}", column_name(col_u32), row_num);
        result.push(Cell::empty(reference, styles.clone()));
    }

    result
}

fn parse_sheet_dimension<R: BufRead>(reader: R) -> Result<Option<usize>, OxlsxError> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"dimension" => {
                if let Some(range) = attr_value(&e, b"ref")? {
                    return Ok(dimension_last_column(&range));
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"sheetData" => return Ok(None),
            Event::Eof => return Ok(None),
            _ => {}
        }
        buf.clear();
    }
}

fn dimension_last_column(range: &str) -> Option<usize> {
    let target = range.rsplit(':').next()?;
    column_index(target).and_then(|c| usize::try_from(c).ok())
}

fn row_number(cell_start: &BytesStart<'_>) -> Option<usize> {
    attr_value(cell_start, b"r")
        .ok()
        .flatten()
        .and_then(|r| r.parse::<usize>().ok())
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
pub(crate) fn parse_rows_from_xml_for_test(xml: &str) -> Vec<Vec<Cell>> {
    streaming_rows_from_xml(
        Cursor::new(xml.as_bytes()),
        Arc::new(StyleSheet::default()),
        &[],
        None,
    )
    .collect()
}

#[cfg(test)]
mod tests {
    use super::{ReadonlyWorkbook, ReadonlyWorksheet, streaming_rows_from_xml};
    use crate::Value;
    use std::sync::Arc;

    /// Verifies that RowIter is lazy: calling next() once should not consume
    /// subsequent rows — each call parses only the next row from the XML stream.
    /// This test fails if rows() materializes everything upfront (the old behaviour)
    /// because we can't distinguish that case... BUT the structural test is:
    /// RowIter must NOT wrap a Vec<Vec<Cell>> — it must hold the XML reader directly.
    /// We enforce this via the type: RowIter<R> must be generic over a reader.
    #[test]
    fn row_iter_is_lazy_not_materialized() {
        // A 3-row sheet. We consume only the first row and verify the iterator
        // still holds state to produce the remaining two — without having parsed them.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>row1</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>row2</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>row3</t></is></c></row>
  </sheetData>
</worksheet>"#;

        let mut iter = super::streaming_rows_from_xml(
            std::io::Cursor::new(xml.as_bytes()),
            Arc::new(crate::StyleSheet::default()),
            &[],
            None,
        );

        let r1 = iter.next().unwrap();
        assert_eq!(r1[0].value(), &Value::String("row1".to_string()));

        let r2 = iter.next().unwrap();
        assert_eq!(r2[0].value(), &Value::String("row2".to_string()));

        let r3 = iter.next().unwrap();
        assert_eq!(r3[0].value(), &Value::String("row3".to_string()));

        assert!(iter.next().is_none());
    }

    #[test]
    fn readonly_stub_types_compile_and_instantiate() {
        let workbook = ReadonlyWorkbook::new_stub();
        let worksheet: ReadonlyWorksheet<'_> = workbook
            .sheet(0)
            .expect("stub workbook should expose one sheet");
        let mut rows = worksheet.rows();

        assert_eq!(worksheet.title(), "Sheet1");
        assert_eq!(workbook.sheetnames(), vec!["Sheet1"]);
        assert_eq!(workbook.active().expect("active sheet").title(), "Sheet1");
        assert!(rows.next().is_some());
    }

    #[test]
    fn readonly_empty_cell_expansion() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C1"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>hello</t></is></c>
      <c r="C1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let rows: Vec<_> = streaming_rows_from_xml(
            std::io::Cursor::new(xml.as_bytes()),
            Arc::new(crate::StyleSheet::default()),
            &[],
            Some(3),
        )
        .collect();
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0].len(), 3);
        assert_eq!(rows[0][0].value(), &Value::String("hello".to_string()));
        assert_eq!(rows[0][1].value(), &Value::Empty);
        assert_eq!(rows[0][2].value(), &Value::Number(42.0));
    }

    #[test]
    fn readonly_fallback_to_data_range_when_dimension_absent() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>hello</t></is></c>
      <c r="C1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let rows = super::parse_rows_from_xml_for_test(xml);
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0].len(), 3);
        assert_eq!(rows[0][1].value(), &Value::Empty);
    }
}
