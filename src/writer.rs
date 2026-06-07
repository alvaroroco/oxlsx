use crate::styles::CellXf;
use crate::relationships::{SheetRelationship, allocate_next_rid, serialize_sheet_rels, sheet_rels_path};
use crate::{Color, Fill, Font, OxlsxError, StyleSheet, Value, Workbook, Worksheet};
use chrono::NaiveDate;
use indexmap::IndexMap;
use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::Write;
use std::path::Path;
use tempfile::NamedTempFile;
use zip::ZipArchive;
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

pub(crate) struct CollectedStyles {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub cell_xfs: Vec<CellXf>,
    pub num_fmts: Vec<(u32, String)>,
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

pub(crate) fn collect_styles(
    sheets: &[Worksheet],
    original: Option<&StyleSheet>,
    preserve_original: bool,
) -> CollectedStyles {
    let styled_cells = sheets
        .iter()
        .map(|sheet| sheet.cells().len())
        .sum::<usize>();

    let (mut fonts, mut fills, mut cell_xfs, mut num_fmts) =
        if preserve_original && let Some(source) = original {
            (
                source.fonts.clone(),
                source.fills.clone(),
                source.cell_xfs.clone(),
                source.num_fmts_vec(),
            )
        } else {
            (
                vec![Font::default()],
                vec![Fill::none(), Fill::gray125()],
                vec![CellXf {
                    font_id: 0,
                    fill_id: 0,
                    num_fmt_id: 0,
                }],
                Vec::new(),
            )
        };

    if fonts.capacity() < fonts.len() + styled_cells {
        fonts.reserve(styled_cells);
    }
    if fills.capacity() < fills.len() + styled_cells {
        fills.reserve(styled_cells);
    }
    if cell_xfs.capacity() < cell_xfs.len() + styled_cells {
        cell_xfs.reserve(styled_cells);
    }

    let mut cell_style_ids = HashMap::with_capacity(styled_cells);

    for (sheet_idx, sheet) in sheets.iter().enumerate() {
        for (reference, cell) in sheet.cells() {
            if preserve_original && cell.write_font.is_none() && cell.write_fill.is_none() {
                continue;
            }

            if cell.write_font.is_none() && cell.write_fill.is_none() {
                continue;
            }

            let base_xf = cell_xfs
                .get(cell.style_id().as_usize())
                .cloned()
                .unwrap_or(CellXf {
                    font_id: 0,
                    fill_id: 0,
                    num_fmt_id: 0,
                });

            let font = match &cell.write_font {
                Some(font) => font.clone(),
                None => fonts
                    .get(base_xf.font_id as usize)
                    .cloned()
                    .unwrap_or_default(),
            };
            let fill = match &cell.write_fill {
                Some(fill) => fill.clone(),
                None => fills
                    .get(base_xf.fill_id as usize)
                    .cloned()
                    .unwrap_or_else(Fill::none),
            };

            let font_idx = find_or_push(&mut fonts, font) as u32;
            let fill_idx = find_or_push(&mut fills, fill) as u32;
            let style_id = find_or_push(
                &mut cell_xfs,
                CellXf {
                    font_id: font_idx,
                    fill_id: fill_idx,
                    num_fmt_id: base_xf.num_fmt_id,
                },
            ) as u32;

            cell_style_ids.insert((sheet_idx, reference.clone()), style_id);
        }
    }

    for (sheet_idx, sheet) in sheets.iter().enumerate() {
        for (reference, cell) in sheet.cells() {
            if !matches!(cell.value(), Value::Date(_)) {
                continue;
            }

            let style_key = (sheet_idx, reference.clone());
            let existing_style_id = cell_style_ids
                .get(&style_key)
                .copied()
                .unwrap_or(cell.style_id().as_usize() as u32);

            let existing_xf = cell_xfs
                .get(existing_style_id as usize)
                .cloned()
                .unwrap_or(CellXf {
                    font_id: 0,
                    fill_id: 0,
                    num_fmt_id: 0,
                });

            if is_builtin_date_num_fmt(existing_xf.num_fmt_id) {
                continue;
            }

            let style_id = find_or_push(
                &mut cell_xfs,
                CellXf {
                    font_id: existing_xf.font_id,
                    fill_id: existing_xf.fill_id,
                    num_fmt_id: 14,
                },
            ) as u32;
            cell_style_ids.insert(style_key, style_id);
        }
    }

    num_fmts.sort_by_key(|(id, _)| *id);

    CollectedStyles {
        fonts,
        fills,
        cell_xfs,
        num_fmts,
        cell_style_ids,
    }
}

fn collect_shared_strings(sheets: &[Worksheet]) -> IndexMap<String, usize> {
    let mut shared_strings = IndexMap::new();

    for sheet in sheets {
        for cell in sheet.cells().values() {
            if let Value::String(text) = cell.value() {
                let next_index = shared_strings.len();
                shared_strings.entry(text.clone()).or_insert(next_index);
            }
        }
    }

    shared_strings
}

fn find_or_push<T: PartialEq>(items: &mut Vec<T>, value: T) -> usize {
    if let Some(idx) = items.iter().position(|item| item == &value) {
        idx
    } else {
        items.push(value);
        items.len() - 1
    }
}

fn is_builtin_date_num_fmt(num_fmt_id: u32) -> bool {
    (14..=17).contains(&num_fmt_id) || num_fmt_id == 22
}

pub fn write_xlsx(
    wb: &Workbook,
    path: &Path,
    source_path: Option<&Path>,
) -> Result<(), OxlsxError> {
    if source_path.is_some_and(|source| source == path) {
        let parent = path.parent().unwrap_or_else(|| Path::new("."));
        let temp = NamedTempFile::new_in(parent)?;
        write_xlsx_to_path(wb, temp.path(), source_path)?;
        temp.persist(path)
            .map_err(|err| OxlsxError::Io(err.error))?;
        return Ok(());
    }

    write_xlsx_to_path(wb, path, source_path)
}

fn write_xlsx_to_path(
    wb: &Workbook,
    path: &Path,
    source_path: Option<&Path>,
) -> Result<(), OxlsxError> {
    let file = File::create(path)?;
    let mut zip = ZipWriter::new(file);
    let options = SimpleFileOptions::default();
    let sheets = wb.worksheets();
    let preserve_original = source_path.is_some();
    let shared_strings = collect_shared_strings(sheets);
    let has_shared_strings = !shared_strings.is_empty();
    let collected = collect_styles(sheets, wb.original_styles.as_deref(), preserve_original);
    let mut written_entries = HashSet::new();

    if !preserve_original {
        write_content_types(&mut zip, options, sheets.len(), has_shared_strings)?;
        written_entries.insert("[Content_Types].xml".to_string());
        write_root_rels(&mut zip, options)?;
        written_entries.insert("_rels/.rels".to_string());
    }

    write_workbook_xml(&mut zip, options, wb)?;
    written_entries.insert("xl/workbook.xml".to_string());

    write_workbook_rels(&mut zip, options, sheets.len(), has_shared_strings)?;
    written_entries.insert("xl/_rels/workbook.xml.rels".to_string());

    write_styles_xml(&mut zip, options, &collected)?;
    written_entries.insert("xl/styles.xml".to_string());

    if has_shared_strings {
        write_shared_strings_xml(&mut zip, options, &shared_strings)?;
        written_entries.insert("xl/sharedStrings.xml".to_string());
    }

    for (idx, ws) in sheets.iter().enumerate() {
        write_sheet_xml(
            &mut zip,
            options,
            ws,
            idx,
            &collected.cell_style_ids,
            &shared_strings,
        )?;
        written_entries.insert(format!("xl/worksheets/sheet{}.xml", idx + 1));

        if let Some(rels_part) = write_sheet_rels_xml(&mut zip, options, ws, idx)? {
            written_entries.insert(rels_part);
        }
    }

    if let Some(source) = source_path {
        copy_unknown_entries(
            &mut zip,
            source,
            &written_entries,
            sheets.len(),
            preserve_original,
        )?;
    }

    zip.finish()?;
    Ok(())
}

fn copy_unknown_entries(
    zip: &mut ZipWriter<File>,
    source: &Path,
    written_entries: &HashSet<String>,
    sheet_count: usize,
    preserve_original: bool,
) -> Result<(), OxlsxError> {
    let source_file = File::open(source)?;
    let mut archive = ZipArchive::new(source_file)?;

    for index in 0..archive.len() {
        let source_entry = archive.by_index(index)?;
        let name = source_entry.name().to_string();
        if written_entries.contains(&name)
            || is_writer_owned_entry(&name, sheet_count, preserve_original)
        {
            continue;
        }
        zip.raw_copy_file(source_entry)?;
    }

    Ok(())
}

fn is_writer_owned_entry(name: &str, sheet_count: usize, preserve_original: bool) -> bool {
    if !preserve_original && (name == "[Content_Types].xml" || name == "_rels/.rels") {
        return true;
    }

    if name == "xl/workbook.xml"
        || name == "xl/_rels/workbook.xml.rels"
        || name == "xl/styles.xml"
        || name == "xl/sharedStrings.xml"
    {
        return true;
    }

    (1..=sheet_count).any(|sheet_idx| {
        name == format!("xl/worksheets/sheet{sheet_idx}.xml")
            || name == format!("xl/worksheets/_rels/sheet{sheet_idx}.xml.rels")
    })
}

fn write_content_types(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    sheet_count: usize,
    has_shared_strings: bool,
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
    if has_shared_strings {
        write_empty(
            &mut xml,
            "Override",
            &[
                ("PartName", "/xl/sharedStrings.xml"),
                (
                    "ContentType",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                ),
            ],
        )?;
    }
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
    has_shared_strings: bool,
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
    if has_shared_strings {
        let rid = format!("rId{}", sheet_count + 2);
        write_empty(
            &mut xml,
            "Relationship",
            &[
                ("Id", rid.as_str()),
                (
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                ),
                ("Target", "sharedStrings.xml"),
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

    if !collected.num_fmts.is_empty() {
        let count = collected.num_fmts.len().to_string();
        write_start(&mut xml, "numFmts", &[("count", count.as_str())])?;
        for (num_fmt_id, code) in &collected.num_fmts {
            let num_fmt_id = num_fmt_id.to_string();
            write_empty(
                &mut xml,
                "numFmt",
                &[
                    ("numFmtId", num_fmt_id.as_str()),
                    ("formatCode", code.as_str()),
                ],
            )?;
        }
        write_end(&mut xml, "numFmts")?;
    }

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
    for xf in &collected.cell_xfs {
        let font_id_str = xf.font_id.to_string();
        let fill_id_str = xf.fill_id.to_string();
        let num_fmt_id_str = xf.num_fmt_id.to_string();
        write_empty(
            &mut xml,
            "xf",
            &[
                ("numFmtId", num_fmt_id_str.as_str()),
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

fn write_shared_strings_xml(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    shared_strings: &IndexMap<String, usize>,
) -> Result<(), OxlsxError> {
    let mut xml = xml_writer_with_decl()?;
    let count = shared_strings.len().to_string();
    write_start(
        &mut xml,
        "sst",
        &[
            (
                "xmlns",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
            ("count", count.as_str()),
            ("uniqueCount", count.as_str()),
        ],
    )?;

    for value in shared_strings.keys() {
        write_start(&mut xml, "si", &[])?;
        write_start(&mut xml, "t", &[])?;
        xml.write_event(Event::Text(BytesText::new(value)))?;
        write_end(&mut xml, "t")?;
        write_end(&mut xml, "si")?;
    }

    write_end(&mut xml, "sst")?;
    write_zip_entry(zip, options, "xl/sharedStrings.xml", xml.into_inner())
}

fn write_sheet_xml(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    ws: &Worksheet,
    sheet_idx: usize,
    style_ids: &HashMap<(usize, String), u32>,
    shared_strings: &IndexMap<String, usize>,
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

    if !ws.column_widths.is_empty() {
        write_start(&mut xml, "cols", &[])?;

        let mut ordered_cols = ws
            .column_widths
            .iter()
            .map(|(col, width)| (*col, *width))
            .collect::<Vec<_>>();
        ordered_cols.sort_by_key(|(col, _)| *col);

        for (col, width) in ordered_cols {
            let col_str = col.to_string();
            let width_str = width.to_string();
            write_empty(
                &mut xml,
                "col",
                &[
                    ("min", col_str.as_str()),
                    ("max", col_str.as_str()),
                    ("width", width_str.as_str()),
                    ("customWidth", "1"),
                ],
            )?;
        }

        write_end(&mut xml, "cols")?;
    }

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

    let mut row_cells = HashMap::new();
    for (row, col, reference, cell) in ordered_cells {
        if matches!(cell.value(), Value::Empty) {
            continue;
        }

        row_cells
            .entry(row)
            .or_insert_with(Vec::new)
            .push((col, reference, cell));
    }

    let mut rows_to_write = row_cells.keys().copied().collect::<Vec<_>>();
    for row in ws.row_heights.keys() {
        if !rows_to_write.contains(row) {
            rows_to_write.push(*row);
        }
    }
    rows_to_write.sort_unstable();

    for row in rows_to_write {
        let row_attr = row.to_string();
        let height_value = ws.row_heights.get(&row).map(|height| height.to_string());

        if let Some(height_value) = height_value.as_deref() {
            write_start(
                &mut xml,
                "row",
                &[
                    ("r", row_attr.as_str()),
                    ("ht", height_value),
                    ("customHeight", "1"),
                ],
            )?;
        } else {
            write_start(&mut xml, "row", &[("r", row_attr.as_str())])?;
        }

        if let Some(cells) = row_cells.get_mut(&row) {
            cells.sort_by_key(|(col, _, _)| *col);
            for (_, reference, cell) in cells {
                let style_id = style_ids
                    .get(&(sheet_idx, (*reference).to_string()))
                    .copied()
                    .unwrap_or(cell.style_id().as_usize() as u32);
                write_cell(&mut xml, reference, style_id, cell.value(), shared_strings)?;
            }
        }

        write_end(&mut xml, "row")?;
    }

    write_end(&mut xml, "sheetData")?;

    let (hyperlink_entries, _) = collect_sheet_hyperlink_rels(ws);
    if !hyperlink_entries.is_empty() {
        write_start(
            &mut xml,
            "hyperlinks",
            &[(
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            )],
        )?;

        for (reference, rel_id) in hyperlink_entries {
            write_empty(
                &mut xml,
                "hyperlink",
                &[("ref", reference.as_str()), ("r:id", rel_id.as_str())],
            )?;
        }

        write_end(&mut xml, "hyperlinks")?;
    }

    if !ws.merged_cells.is_empty() {
        let count = ws.merged_cells.len().to_string();
        write_start(&mut xml, "mergeCells", &[("count", count.as_str())])?;
        for range_ref in &ws.merged_cells {
            write_empty(&mut xml, "mergeCell", &[("ref", range_ref.as_str())])?;
        }
        write_end(&mut xml, "mergeCells")?;
    }

    write_end(&mut xml, "worksheet")?;
    let part = format!("xl/worksheets/sheet{}.xml", sheet_idx + 1);
    write_zip_entry(zip, options, part.as_str(), xml.into_inner())
}

fn write_sheet_rels_xml(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    ws: &Worksheet,
    sheet_idx: usize,
) -> Result<Option<String>, OxlsxError> {
    let (_, rels) = collect_sheet_hyperlink_rels(ws);

    if rels.is_empty() {
        return Ok(None);
    }

    let xml = serialize_sheet_rels(&rels)?;
    let part = sheet_rels_path(&format!("xl/worksheets/sheet{}.xml", sheet_idx + 1));
    write_zip_entry(zip, options, part.as_str(), xml)?;

    Ok(Some(part))
}

fn collect_sheet_hyperlink_rels(ws: &Worksheet) -> (Vec<(String, String)>, Vec<SheetRelationship>) {
    let mut entries = ws
        .cells()
        .iter()
        .filter_map(|(reference, cell)| {
            cell.write_hyperlink
                .as_ref()
                .map(|hyperlink| (reference.clone(), hyperlink.target().to_string()))
        })
        .collect::<Vec<_>>();
    entries.sort_by_key(|(reference, _)| cell_sort_key(reference));

    let mut rels = ws.sheet_rels.clone();
    let mut hyperlink_entries = Vec::with_capacity(entries.len());
    for (reference, target) in entries {
        let rel_id = allocate_next_rid(&rels);
        rels.push(SheetRelationship::hyperlink(rel_id.clone(), target));
        hyperlink_entries.push((reference, rel_id));
    }

    (hyperlink_entries, rels)
}

fn write_cell(
    xml: &mut Writer<Vec<u8>>,
    reference: &str,
    style_id: u32,
    value: &Value,
    shared_strings: &IndexMap<String, usize>,
) -> Result<(), OxlsxError> {
    let style_str = style_id.to_string();
    let mut base = vec![("r", reference), ("s", style_str.as_str())];

    match value {
        Value::String(text) => {
            let Some(index) = shared_strings.get(text) else {
                return Ok(());
            };
            let index = index.to_string();

            base.push(("t", "s"));
            write_start(xml, "c", &base)?;
            write_start(xml, "v", &[])?;
            xml.write_event(Event::Text(BytesText::new(index.as_str())))?;
            write_end(xml, "v")?;
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
        Value::Formula(formula) => {
            write_start(xml, "c", &base)?;
            write_start(xml, "f", &[])?;
            xml.write_event(Event::Text(BytesText::new(formula)))?;
            write_end(xml, "f")?;
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
    use crate::{Fill, Font, StyleSheet, Workbook, Worksheet};
    use std::env;
    use std::fs;
    use std::io::Read;
    use std::io::{BufReader, Cursor};
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

        let collected = super::collect_styles(&[ws], None, false);

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

        let collected = super::collect_styles(&[ws], None, false);

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

        let collected = super::collect_styles(&[ws], None, false);

        assert_eq!(collected.fills.len(), 3);
    }

    #[test]
    fn write_xlsx_writes_required_ooxml_parts_and_escapes_shared_strings() {
        let path = temp_path("oxlsx_writer_parts");
        let mut wb = Workbook::new();
        wb.active_mut().unwrap().set_cell("A1", "A&B<Z>");
        wb.create_sheet("Data").set_cell("B1", 42.0_f64);

        super::write_xlsx(&wb, &path, None).unwrap();

        let file = fs::File::open(&path).unwrap();
        let mut zip = ZipArchive::new(file).unwrap();

        let required_parts = [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet2.xml",
        ];
        for part in required_parts {
            assert!(zip.by_name(part).is_ok(), "missing part: {part}");
        }

        let mut shared_strings_xml = String::new();
        zip.by_name("xl/sharedStrings.xml")
            .unwrap()
            .read_to_string(&mut shared_strings_xml)
            .unwrap();
        assert!(shared_strings_xml.contains("A&amp;B&lt;Z&gt;"));

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

        super::write_xlsx(&wb, &path, None).unwrap();

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

    #[test]
    fn write_xlsx_accepts_missing_source_path() {
        let path = temp_path("oxlsx_writer_no_source");
        let mut wb = Workbook::new();
        wb.active_mut().unwrap().set_cell("A1", "ok");

        super::write_xlsx(&wb, &path, None).unwrap();

        assert!(path.exists());
        fs::remove_file(path).ok();
    }

    #[test]
    fn write_xlsx_copies_unknown_entries_from_source_archive() {
        let path = temp_path("oxlsx_writer_copy_unknown");
        let wb = Workbook::open("samples/test.xlsx").unwrap();

        super::write_xlsx(&wb, &path, Some(std::path::Path::new("samples/test.xlsx"))).unwrap();

        let file = fs::File::open(&path).unwrap();
        let mut zip = ZipArchive::new(file).unwrap();

        assert!(zip.by_name("xl/theme/theme1.xml").is_ok());
        assert!(zip.by_name("docProps/app.xml").is_ok());

        fs::remove_file(path).ok();
    }

    #[test]
    fn write_xlsx_preserves_original_content_types_and_root_rels_when_source_present() {
        let path = temp_path("oxlsx_writer_preserve_root_parts");
        let wb = Workbook::open("samples/test.xlsx").unwrap();

        super::write_xlsx(&wb, &path, Some(std::path::Path::new("samples/test.xlsx"))).unwrap();

        let source_file = fs::File::open("samples/test.xlsx").unwrap();
        let mut source_zip = ZipArchive::new(source_file).unwrap();
        let mut source_content_types = Vec::new();
        source_zip
            .by_name("[Content_Types].xml")
            .unwrap()
            .read_to_end(&mut source_content_types)
            .unwrap();
        let mut source_root_rels = Vec::new();
        source_zip
            .by_name("_rels/.rels")
            .unwrap()
            .read_to_end(&mut source_root_rels)
            .unwrap();

        let output_file = fs::File::open(&path).unwrap();
        let mut output_zip = ZipArchive::new(output_file).unwrap();
        let mut output_content_types = Vec::new();
        output_zip
            .by_name("[Content_Types].xml")
            .unwrap()
            .read_to_end(&mut output_content_types)
            .unwrap();
        let mut output_root_rels = Vec::new();
        output_zip
            .by_name("_rels/.rels")
            .unwrap()
            .read_to_end(&mut output_root_rels)
            .unwrap();

        assert_eq!(output_content_types, source_content_types);
        assert_eq!(output_root_rels, source_root_rels);

        fs::remove_file(path).ok();
    }

    #[test]
    fn write_xlsx_preserves_source_styles_and_appends_new_xf() {
        let path = temp_path("oxlsx_writer_preserve_styles");
        let mut wb = Workbook::open("samples/test.xlsx").unwrap();
        let ws = wb.active_mut().unwrap();
        ws.set_cell("C1", 99.0_f64);
        ws.set_cell_font(
            "C1",
            Font {
                italic: true,
                ..Default::default()
            },
        );

        super::write_xlsx(&wb, &path, Some(std::path::Path::new("samples/test.xlsx"))).unwrap();

        let source_file = fs::File::open("samples/test.xlsx").unwrap();
        let mut source_zip = ZipArchive::new(source_file).unwrap();
        let mut source_styles_bytes = Vec::new();
        source_zip
            .by_name("xl/styles.xml")
            .unwrap()
            .read_to_end(&mut source_styles_bytes)
            .unwrap();
        let source_styles =
            StyleSheet::parse(BufReader::new(Cursor::new(source_styles_bytes))).unwrap();

        let output_file = fs::File::open(&path).unwrap();
        let mut output_zip = ZipArchive::new(output_file).unwrap();
        let mut output_styles_bytes = Vec::new();
        output_zip
            .by_name("xl/styles.xml")
            .unwrap()
            .read_to_end(&mut output_styles_bytes)
            .unwrap();
        let output_styles =
            StyleSheet::parse(BufReader::new(Cursor::new(output_styles_bytes))).unwrap();

        assert!(output_styles.cell_xfs.len() > source_styles.cell_xfs.len());
        assert_eq!(
            &output_styles.cell_xfs[..source_styles.cell_xfs.len()],
            &source_styles.cell_xfs[..]
        );

        fs::remove_file(path).ok();
    }
}
