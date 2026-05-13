use crate::{Color, Fill, Font, OxlsxError, Value, Workbook, Worksheet};
use chrono::NaiveDate;
use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use std::collections::HashMap;
use std::fs::File;
use std::io::Write;
use std::path::Path;
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

pub(crate) struct CollectedStyles {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub cell_xfs: Vec<(u32, u32)>,
    pub cell_style_ids: HashMap<(usize, String), u32>,
}

// Returns (row, col) for sorting cells in sheet XML.
pub(crate) fn cell_sort_key(reference: &str) -> (u32, u32) {
    let mut col = 0_u32;
    let mut row_text = String::new();

    for ch in reference.chars() {
        if ch.is_ascii_alphabetic() {
            let upper = ch.to_ascii_uppercase() as u32;
            col = col * 26 + (upper - 'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            row_text.push(ch);
        }
    }

    let row = row_text.parse::<u32>().unwrap_or(0);
    (row, col)
}

pub(crate) fn collect_styles(sheets: &[Worksheet]) -> CollectedStyles {
    let styled_cells = sheets
        .iter()
        .map(|sheet| sheet.cells().len())
        .sum::<usize>();

    let mut fonts = Vec::with_capacity(1 + styled_cells);
    fonts.push(Font::default());

    let mut fills = Vec::with_capacity(2 + styled_cells);
    fills.push(Fill::none());
    fills.push(Fill::gray125());

    let mut cell_xfs = Vec::with_capacity(1 + styled_cells);
    cell_xfs.push((0_u32, 0_u32));

    let mut cell_style_ids = HashMap::with_capacity(styled_cells);

    for (sheet_idx, sheet) in sheets.iter().enumerate() {
        for (reference, cell) in sheet.cells() {
            if cell.write_font.is_none() && cell.write_fill.is_none() {
                continue;
            }

            let font = cell.write_font.clone().unwrap_or_default();
            let fill = cell.write_fill.clone().unwrap_or_else(Fill::none);

            let font_idx = find_or_push(&mut fonts, font) as u32;
            let fill_idx = find_or_push(&mut fills, fill) as u32;
            let style_id = find_or_push(&mut cell_xfs, (font_idx, fill_idx)) as u32;

            cell_style_ids.insert((sheet_idx, reference.clone()), style_id);
        }
    }

    CollectedStyles {
        fonts,
        fills,
        cell_xfs,
        cell_style_ids,
    }
}

fn find_or_push<T: PartialEq>(items: &mut Vec<T>, value: T) -> usize {
    if let Some(idx) = items.iter().position(|item| item == &value) {
        idx
    } else {
        items.push(value);
        items.len() - 1
    }
}

pub fn write_xlsx(wb: &Workbook, path: &Path) -> Result<(), OxlsxError> {
    let file = File::create(path)?;
    let mut zip = ZipWriter::new(file);
    let options = SimpleFileOptions::default();
    let sheets = wb.worksheets();
    let collected = collect_styles(sheets);

    write_content_types(&mut zip, options, sheets.len())?;
    write_root_rels(&mut zip, options)?;
    write_workbook_xml(&mut zip, options, wb)?;
    write_workbook_rels(&mut zip, options, sheets.len())?;
    write_styles_xml(&mut zip, options, &collected)?;

    for (idx, ws) in sheets.iter().enumerate() {
        write_sheet_xml(&mut zip, options, ws, idx, &collected.cell_style_ids)?;
    }

    zip.finish()?;
    Ok(())
}

fn write_content_types(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    sheet_count: usize,
) -> Result<(), OxlsxError> {
    let mut xml = xml_writer_with_decl()?;
    write_start(
        &mut xml,
        "Types",
        &[(
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/content-types",
        )],
    )?;

    write_empty(
        &mut xml,
        "Default",
        &[
            ("Extension", "rels"),
            (
                "ContentType",
                "application/vnd.openxmlformats-package.relationships+xml",
            ),
        ],
    )?;
    write_empty(
        &mut xml,
        "Default",
        &[("Extension", "xml"), ("ContentType", "application/xml")],
    )?;
    write_empty(
        &mut xml,
        "Override",
        &[
            ("PartName", "/xl/workbook.xml"),
            (
                "ContentType",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            ),
        ],
    )?;
    write_empty(
        &mut xml,
        "Override",
        &[
            ("PartName", "/xl/styles.xml"),
            (
                "ContentType",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
            ),
        ],
    )?;
    for i in 1..=sheet_count {
        let part_name = format!("/xl/worksheets/sheet{i}.xml");
        write_empty(
            &mut xml,
            "Override",
            &[
                ("PartName", part_name.as_str()),
                (
                    "ContentType",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                ),
            ],
        )?;
    }

    write_end(&mut xml, "Types")?;
    write_zip_entry(zip, options, "[Content_Types].xml", xml.into_inner())
}

fn write_root_rels(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
) -> Result<(), OxlsxError> {
    let mut xml = xml_writer_with_decl()?;
    write_start(
        &mut xml,
        "Relationships",
        &[(
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        )],
    )?;
    write_empty(
        &mut xml,
        "Relationship",
        &[
            ("Id", "rId1"),
            (
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            ),
            ("Target", "xl/workbook.xml"),
        ],
    )?;
    write_end(&mut xml, "Relationships")?;
    write_zip_entry(zip, options, "_rels/.rels", xml.into_inner())
}

fn write_workbook_xml(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    wb: &Workbook,
) -> Result<(), OxlsxError> {
    let mut xml = xml_writer_with_decl()?;
    write_start(
        &mut xml,
        "workbook",
        &[
            (
                "xmlns",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            ),
        ],
    )?;

    write_start(&mut xml, "bookViews", &[])?;
    write_empty(&mut xml, "workbookView", &[("activeTab", "0")])?;
    write_end(&mut xml, "bookViews")?;

    write_start(&mut xml, "sheets", &[])?;
    for (idx, ws) in wb.worksheets().iter().enumerate() {
        let sheet_id = (idx + 1).to_string();
        let rid = (idx + 2).to_string();
        write_empty(
            &mut xml,
            "sheet",
            &[
                ("name", ws.title()),
                ("sheetId", sheet_id.as_str()),
                ("r:id", format!("rId{rid}").as_str()),
            ],
        )?;
    }
    write_end(&mut xml, "sheets")?;
    write_end(&mut xml, "workbook")?;
    write_zip_entry(zip, options, "xl/workbook.xml", xml.into_inner())
}

fn write_workbook_rels(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    sheet_count: usize,
) -> Result<(), OxlsxError> {
    let mut xml = xml_writer_with_decl()?;
    write_start(
        &mut xml,
        "Relationships",
        &[(
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        )],
    )?;
    write_empty(
        &mut xml,
        "Relationship",
        &[
            ("Id", "rId1"),
            (
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            ),
            ("Target", "styles.xml"),
        ],
    )?;
    for i in 1..=sheet_count {
        let rid = format!("rId{}", i + 1);
        let target = format!("worksheets/sheet{i}.xml");
        write_empty(
            &mut xml,
            "Relationship",
            &[
                ("Id", rid.as_str()),
                (
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                ),
                ("Target", target.as_str()),
            ],
        )?;
    }
    write_end(&mut xml, "Relationships")?;
    write_zip_entry(zip, options, "xl/_rels/workbook.xml.rels", xml.into_inner())
}

fn write_styles_xml(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    collected: &CollectedStyles,
) -> Result<(), OxlsxError> {
    let mut xml = xml_writer_with_decl()?;
    write_start(
        &mut xml,
        "styleSheet",
        &[(
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        )],
    )?;

    let font_count = collected.fonts.len().to_string();
    write_start(&mut xml, "fonts", &[("count", font_count.as_str())])?;
    for font in &collected.fonts {
        write_font(&mut xml, font)?;
    }
    write_end(&mut xml, "fonts")?;

    let fill_count = collected.fills.len().to_string();
    write_start(&mut xml, "fills", &[("count", fill_count.as_str())])?;
    for fill in &collected.fills {
        write_fill(&mut xml, fill)?;
    }
    write_end(&mut xml, "fills")?;

    write_start(&mut xml, "borders", &[("count", "1")])?;
    write_start(&mut xml, "border", &[])?;
    write_empty(&mut xml, "left", &[])?;
    write_empty(&mut xml, "right", &[])?;
    write_empty(&mut xml, "top", &[])?;
    write_empty(&mut xml, "bottom", &[])?;
    write_empty(&mut xml, "diagonal", &[])?;
    write_end(&mut xml, "border")?;
    write_end(&mut xml, "borders")?;

    write_start(&mut xml, "cellStyleXfs", &[("count", "1")])?;
    write_empty(
        &mut xml,
        "xf",
        &[
            ("numFmtId", "0"),
            ("fontId", "0"),
            ("fillId", "0"),
            ("borderId", "0"),
        ],
    )?;
    write_end(&mut xml, "cellStyleXfs")?;

    let xf_count = collected.cell_xfs.len().to_string();
    write_start(&mut xml, "cellXfs", &[("count", xf_count.as_str())])?;
    for (font_id, fill_id) in &collected.cell_xfs {
        let font_id_str = font_id.to_string();
        let fill_id_str = fill_id.to_string();
        write_empty(
            &mut xml,
            "xf",
            &[
                ("numFmtId", "0"),
                ("fontId", font_id_str.as_str()),
                ("fillId", fill_id_str.as_str()),
                ("borderId", "0"),
                ("xfId", "0"),
            ],
        )?;
    }
    write_end(&mut xml, "cellXfs")?;

    write_start(&mut xml, "cellStyles", &[("count", "1")])?;
    write_empty(
        &mut xml,
        "cellStyle",
        &[("name", "Normal"), ("xfId", "0"), ("builtinId", "0")],
    )?;
    write_end(&mut xml, "cellStyles")?;

    write_end(&mut xml, "styleSheet")?;
    write_zip_entry(zip, options, "xl/styles.xml", xml.into_inner())
}

pub(crate) fn write_font(xml: &mut Writer<Vec<u8>>, font: &Font) -> Result<(), OxlsxError> {
    write_start(xml, "font", &[])?;
    if font.bold {
        write_empty(xml, "b", &[])?;
    }
    if font.italic {
        write_empty(xml, "i", &[])?;
    }

    if font.size > 0.0 {
        let size = font.size.to_string();
        write_empty(xml, "sz", &[("val", size.as_str())])?;
    } else if font.name.is_empty() && font.color.is_none() {
        write_empty(xml, "sz", &[("val", "11")])?;
    }

    if !font.name.is_empty() {
        write_empty(xml, "name", &[("val", font.name.as_str())])?;
    } else if font.size <= 0.0 && font.color.is_none() {
        write_empty(xml, "name", &[("val", "Calibri")])?;
    }

    if let Some(color) = &font.color {
        write_color(xml, "color", color)?;
    }

    write_end(xml, "font")
}

pub(crate) fn write_fill(xml: &mut Writer<Vec<u8>>, fill: &Fill) -> Result<(), OxlsxError> {
    write_start(xml, "fill", &[])?;
    write_start(
        xml,
        "patternFill",
        &[("patternType", fill.pattern_type.as_str())],
    )?;

    if let Some(color) = &fill.fg_color {
        write_color(xml, "fgColor", color)?;
    }

    write_end(xml, "patternFill")?;
    write_end(xml, "fill")
}

fn write_sheet_xml(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    ws: &Worksheet,
    sheet_idx: usize,
    style_ids: &HashMap<(usize, String), u32>,
) -> Result<(), OxlsxError> {
    let mut xml = xml_writer_with_decl()?;
    write_start(
        &mut xml,
        "worksheet",
        &[(
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        )],
    )?;
    write_start(&mut xml, "sheetData", &[])?;

    let mut ordered_cells = ws
        .cells()
        .iter()
        .map(|(reference, cell)| {
            let (row, col) = cell_sort_key(reference);
            (row, col, reference.as_str(), cell)
        })
        .collect::<Vec<_>>();
    ordered_cells.sort_by_key(|(row, col, _, _)| (*row, *col));

    let mut current_row: Option<u32> = None;
    for (row, _, reference, cell) in ordered_cells {
        if matches!(cell.value(), Value::Empty) {
            continue;
        }

        if current_row != Some(row) {
            if current_row.is_some() {
                write_end(&mut xml, "row")?;
            }
            let row_attr = row.to_string();
            write_start(&mut xml, "row", &[("r", row_attr.as_str())])?;
            current_row = Some(row);
        }

        let style_id = style_ids
            .get(&(sheet_idx, reference.to_string()))
            .copied()
            .unwrap_or(0);
        write_cell(&mut xml, reference, style_id, cell.value())?;
    }

    if current_row.is_some() {
        write_end(&mut xml, "row")?;
    }

    write_end(&mut xml, "sheetData")?;
    write_end(&mut xml, "worksheet")?;
    let part = format!("xl/worksheets/sheet{}.xml", sheet_idx + 1);
    write_zip_entry(zip, options, part.as_str(), xml.into_inner())
}

fn write_cell(
    xml: &mut Writer<Vec<u8>>,
    reference: &str,
    style_id: u32,
    value: &Value,
) -> Result<(), OxlsxError> {
    let style_str = style_id.to_string();
    let mut base = vec![("r", reference), ("s", style_str.as_str())];

    match value {
        Value::String(text) => {
            base.push(("t", "inlineStr"));
            write_start(xml, "c", &base)?;
            write_start(xml, "is", &[])?;
            write_start(xml, "t", &[])?;
            xml.write_event(Event::Text(BytesText::new(text)))?;
            write_end(xml, "t")?;
            write_end(xml, "is")?;
            write_end(xml, "c")?;
        }
        Value::Number(n) => {
            write_start(xml, "c", &base)?;
            write_start(xml, "v", &[])?;
            let value = n.to_string();
            xml.write_event(Event::Text(BytesText::new(&value)))?;
            write_end(xml, "v")?;
            write_end(xml, "c")?;
        }
        Value::Bool(b) => {
            base.push(("t", "b"));
            write_start(xml, "c", &base)?;
            write_start(xml, "v", &[])?;
            let value = if *b { "1" } else { "0" };
            xml.write_event(Event::Text(BytesText::new(value)))?;
            write_end(xml, "v")?;
            write_end(xml, "c")?;
        }
        Value::Date(date) => {
            write_start(xml, "c", &base)?;
            write_start(xml, "v", &[])?;
            let value = date_to_excel_serial(*date).to_string();
            xml.write_event(Event::Text(BytesText::new(&value)))?;
            write_end(xml, "v")?;
            write_end(xml, "c")?;
        }
        Value::Empty => {}
    }

    Ok(())
}

fn date_to_excel_serial(date: NaiveDate) -> i64 {
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30).expect("valid excel epoch");
    date.signed_duration_since(epoch).num_days()
}

fn write_color(xml: &mut Writer<Vec<u8>>, tag: &str, color: &Color) -> Result<(), OxlsxError> {
    match color {
        Color::Argb(argb) => write_empty(xml, tag, &[("rgb", argb.as_str())]),
        Color::Indexed(indexed) => {
            let indexed = indexed.to_string();
            write_empty(xml, tag, &[("indexed", indexed.as_str())])
        }
        Color::Theme(theme) => {
            let theme = theme.to_string();
            write_empty(xml, tag, &[("theme", theme.as_str())])
        }
    }
}

fn write_zip_entry(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    name: &str,
    data: Vec<u8>,
) -> Result<(), OxlsxError> {
    zip.start_file(name, options)?;
    zip.write_all(&data)?;
    Ok(())
}

fn xml_writer_with_decl() -> Result<Writer<Vec<u8>>, OxlsxError> {
    let mut xml = Writer::new(Vec::new());
    xml.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;
    Ok(xml)
}

fn write_start(
    xml: &mut Writer<Vec<u8>>,
    name: &str,
    attrs: &[(&str, &str)],
) -> Result<(), OxlsxError> {
    let mut elem = BytesStart::new(name);
    for (k, v) in attrs {
        elem.push_attribute((*k, *v));
    }
    xml.write_event(Event::Start(elem))?;
    Ok(())
}

fn write_empty(
    xml: &mut Writer<Vec<u8>>,
    name: &str,
    attrs: &[(&str, &str)],
) -> Result<(), OxlsxError> {
    let mut elem = BytesStart::new(name);
    for (k, v) in attrs {
        elem.push_attribute((*k, *v));
    }
    xml.write_event(Event::Empty(elem))?;
    Ok(())
}

fn write_end(xml: &mut Writer<Vec<u8>>, name: &str) -> Result<(), OxlsxError> {
    xml.write_event(Event::End(BytesEnd::new(name)))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use crate::{Fill, Font, Workbook, Worksheet};
    use std::env;
    use std::fs;
    use std::io::Read;
    use std::time::{SystemTime, UNIX_EPOCH};
    use zip::ZipArchive;

    fn temp_path(name: &str) -> std::path::PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        env::temp_dir().join(format!("{name}_{nanos}.xlsx"))
    }

    #[test]
    fn cell_sort_key_orders_column_and_row_coordinates() {
        assert_eq!(super::cell_sort_key("A1"), (1, 1));
        assert_eq!(super::cell_sort_key("Z1"), (1, 26));
        assert_eq!(super::cell_sort_key("AA1"), (1, 27));
        assert_eq!(super::cell_sort_key("B10"), (10, 2));
    }

    #[test]
    fn collect_styles_defaults_without_styled_cells() {
        let mut ws = Worksheet::new("Sheet");
        ws.set_cell("A1", "Hello");

        let collected = super::collect_styles(&[ws]);

        assert_eq!(collected.fonts.len(), 1);
        assert_eq!(collected.fills.len(), 2);
        assert_eq!(collected.cell_xfs.len(), 1);
        assert!(collected.cell_style_ids.is_empty());
    }

    #[test]
    fn collect_styles_dedupes_repeated_font_styles() {
        let mut ws = Worksheet::new("Sheet");

        ws.set_cell_font(
            "A1",
            Font {
                bold: true,
                ..Default::default()
            },
        );
        ws.set_cell_font(
            "A2",
            Font {
                bold: true,
                ..Default::default()
            },
        );
        ws.set_cell_font(
            "A3",
            Font {
                bold: true,
                ..Default::default()
            },
        );

        let collected = super::collect_styles(&[ws]);

        assert_eq!(collected.fonts.len(), 2);
        assert_eq!(collected.fills.len(), 2);
        assert_eq!(collected.cell_xfs.len(), 2);
        assert_eq!(collected.cell_style_ids.len(), 3);
    }

    #[test]
    fn collect_styles_adds_custom_fill_after_defaults() {
        let mut ws = Worksheet::new("Sheet");

        ws.set_cell_fill(
            "A1",
            Fill {
                pattern_type: "solid".to_string(),
                ..Default::default()
            },
        );

        let collected = super::collect_styles(&[ws]);

        assert_eq!(collected.fills.len(), 3);
    }

    #[test]
    fn write_xlsx_writes_required_ooxml_parts_and_escapes_inline_strings() {
        let path = temp_path("oxlsx_writer_parts");
        let mut wb = Workbook::new();
        wb.active_mut().unwrap().set_cell("A1", "A&B<Z>");
        wb.create_sheet("Data").set_cell("B1", 42.0_f64);

        super::write_xlsx(&wb, &path).unwrap();

        let file = fs::File::open(&path).unwrap();
        let mut zip = ZipArchive::new(file).unwrap();

        let required_parts = [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/styles.xml",
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet2.xml",
        ];
        for part in required_parts {
            assert!(zip.by_name(part).is_ok(), "missing part: {part}");
        }

        let mut sheet1_xml = String::new();
        zip.by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet1_xml)
            .unwrap();
        assert!(sheet1_xml.contains("A&amp;B&lt;Z&gt;"));

        fs::remove_file(path).ok();
    }

    #[test]
    fn write_xlsx_sorts_cells_and_skips_empty_values() {
        let path = temp_path("oxlsx_writer_sheet_order");
        let mut wb = Workbook::new();
        let ws = wb.active_mut().unwrap();
        ws.set_cell_font(
            "C3",
            Font {
                bold: true,
                ..Default::default()
            },
        );
        ws.set_cell("B2", 2.0_f64);
        ws.set_cell("A1", 1.0_f64);

        super::write_xlsx(&wb, &path).unwrap();

        let file = fs::File::open(&path).unwrap();
        let mut zip = ZipArchive::new(file).unwrap();
        let mut sheet_xml = String::new();
        zip.by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet_xml)
            .unwrap();

        let a1 = sheet_xml.find("r=\"A1\"").unwrap();
        let b2 = sheet_xml.find("r=\"B2\"").unwrap();
        assert!(a1 < b2);
        assert!(!sheet_xml.contains("r=\"C3\""));

        fs::remove_file(path).ok();
    }
}
