use crate::styles::CellXf;
use crate::{
    Color, Comment, Fill, Font, Hyperlink, OxlsxError, StyleSheet, Value, Workbook, Worksheet,
    defined_names, sheet_rels,
};
use chrono::NaiveDate;
use indexmap::IndexMap;
use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{Read, Write};
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
    // Reject duplicate table displayNames anywhere in the workbook BEFORE creating
    // any file (or temp file): Excel requires workbook-wide uniqueness, and a
    // malformed table set corrupts the whole file. Doing this first guarantees no
    // partial/0-byte artifact is left behind when the guard trips. The Python layer
    // also checks early, at add_table time.
    ensure_unique_table_display_names(wb.worksheets())?;

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
    let commented_sheets = sheets
        .iter()
        .enumerate()
        .filter_map(|(idx, sheet)| has_comments(sheet).then_some(idx + 1))
        .collect::<Vec<_>>();
    let collected = collect_styles(sheets, wb.original_styles.as_deref(), preserve_original);
    let mut written_entries = HashSet::new();

    // Allocate the internal `<table id>` for every table (verbatim tables keep
    // their original id; model tables get collision-free ids). Part numbers —
    // used for the part filename (`xl/tables/tableN.xml`), the sheet-rel target,
    // the `<tableParts>` rId, and the content-type override — are a SEPARATE,
    // strictly sequential workbook-wide sequence, decoupled from the id.
    let table_ids = assign_table_ids(sheets);
    let total_tables = table_ids.iter().map(Vec::len).sum::<usize>();

    let mut next_part = 1u32;
    let part_numbers: Vec<Vec<u32>> = table_ids
        .iter()
        .map(|sheet| {
            sheet
                .iter()
                .map(|_| {
                    let number = next_part;
                    next_part += 1;
                    number
                })
                .collect()
        })
        .collect();

    if !preserve_original {
        write_content_types(
            &mut zip,
            options,
            sheets.len(),
            has_shared_strings,
            &commented_sheets,
            total_tables,
        )?;
        written_entries.insert("[Content_Types].xml".to_string());
        write_root_rels(&mut zip, options)?;
        written_entries.insert("_rels/.rels".to_string());
    } else if has_shared_strings || !commented_sheets.is_empty() || total_tables > 0 {
        write_content_types_from_source(
            &mut zip,
            options,
            source_path.expect("source path exists when preserve_original is true"),
            &commented_sheets,
            total_tables,
            has_shared_strings,
        )?;
        written_entries.insert("[Content_Types].xml".to_string());
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
        let mut sheet_rels_entries = wb.original_sheet_rels.get(idx).cloned().unwrap_or_default();
        let mut legacy_drawing_rid = None;
        let sheet_part_numbers = &part_numbers[idx];

        if has_comments(ws) {
            let comment_target = format!("../comments{}.xml", idx + 1);
            let vml_target = format!("../drawings/vmlDrawing{}.vml", idx + 1);
            sheet_rels_entries = sheet_rels::upsert_comment_rels(
                &sheet_rels_entries,
                comment_target.as_str(),
                vml_target.as_str(),
            );
            legacy_drawing_rid = sheet_rels::relation_id_for_type(
                &sheet_rels_entries,
                sheet_rels::VML_DRAWING_REL_TYPE,
            )
            .map(ToString::to_string);
        }

        // Allocate table relationships AFTER comments so rIds never collide.
        let mut table_rids: Vec<String> = Vec::new();
        if !sheet_part_numbers.is_empty() {
            let targets = sheet_part_numbers
                .iter()
                .map(|n| format!("../tables/table{n}.xml"))
                .collect::<Vec<_>>();
            let target_refs = targets.iter().map(String::as_str).collect::<Vec<_>>();
            sheet_rels_entries = sheet_rels::upsert_table_rels(&sheet_rels_entries, &target_refs);
            for target in &targets {
                if let Some(rid) = sheet_rels::relation_id_for_target(&sheet_rels_entries, target) {
                    table_rids.push(rid.to_string());
                }
            }
        }

        // Collect emittable hyperlinks in deterministic (row, col) order.
        // External links (non-empty target) allocate a rel + `r:id` AFTER
        // comments/tables (via `upsert_hyperlink_rels` → `next_rel_id`), so
        // rIds never collide; internal `location`-only links carry `location`
        // and fabricate no rId.
        let mut hyperlink_cells = ws
            .cells()
            .iter()
            .filter_map(|(reference, cell)| {
                cell.hyperlink().and_then(|hyperlink| {
                    (!hyperlink.target.is_empty() || hyperlink.location.is_some())
                        .then_some((reference.as_str(), hyperlink))
                })
            })
            .collect::<Vec<_>>();
        hyperlink_cells.sort_by_key(|(reference, _)| cell_sort_key(reference));

        let external_targets = hyperlink_cells
            .iter()
            .filter(|(_, hyperlink)| !hyperlink.target.is_empty())
            .map(|(_, hyperlink)| hyperlink.target.as_str())
            .collect::<Vec<_>>();
        let has_external_hyperlinks = !external_targets.is_empty();
        if has_external_hyperlinks {
            sheet_rels_entries =
                sheet_rels::upsert_hyperlink_rels(&sheet_rels_entries, &external_targets);
        }
        // After the upsert, the only HYPERLINK_REL_TYPE entries are the freshly
        // allocated ones, in `external_targets` order.
        let hyperlink_rids = sheet_rels_entries
            .iter()
            .filter(|entry| entry.rel_type == sheet_rels::HYPERLINK_REL_TYPE)
            .map(|entry| entry.id.clone())
            .collect::<Vec<_>>();
        let mut next_hyperlink_rid = hyperlink_rids.iter();
        let hyperlink_refs = hyperlink_cells
            .iter()
            .map(|(reference, hyperlink)| {
                let rel_id = if hyperlink.target.is_empty() {
                    None
                } else {
                    next_hyperlink_rid.next().map(String::as_str)
                };
                HyperlinkRef {
                    reference,
                    hyperlink,
                    rel_id,
                }
            })
            .collect::<Vec<_>>();

        write_sheet_xml(
            &mut zip,
            options,
            ws,
            idx,
            &collected.cell_style_ids,
            &shared_strings,
            &SheetRelRefs {
                legacy_drawing_rid: legacy_drawing_rid.as_deref(),
                table_rids: &table_rids,
                hyperlinks: &hyperlink_refs,
            },
        )?;
        written_entries.insert(format!("xl/worksheets/sheet{}.xml", idx + 1));

        if has_comments(ws) {
            write_comments_xml(&mut zip, options, ws, idx + 1)?;
            write_vml_drawing(&mut zip, options, ws, idx + 1)?;
            written_entries.insert(format!("xl/comments{}.xml", idx + 1));
            written_entries.insert(format!("xl/drawings/vmlDrawing{}.vml", idx + 1));
        }

        // Re-emit every table part this sheet owns. Table parts are
        // writer-owned (see `is_writer_owned_entry`), so the originals are NOT
        // copied by `copy_unknown_entries` — this loop is the single source of
        // truth. A clean retained table is re-emitted byte-for-byte verbatim
        // (preserving unmodeled features); a dirty/new table is serialized from
        // the model. Either way the part filename is the sequential part number,
        // keeping the `<tableParts>`/rel/content-type wiring consistent.
        let sheet_table_ids = &table_ids[idx];
        for (table_idx, (table, &part_number)) in ws
            .tables()
            .iter()
            .zip(sheet_part_numbers.iter())
            .enumerate()
        {
            let (assigned_id, is_verbatim) = sheet_table_ids[table_idx];
            let part = format!("xl/tables/table{part_number}.xml");
            if is_verbatim {
                let raw = ws.table_retention[table_idx]
                    .as_ref()
                    .expect("a verbatim table always has retained bytes");
                write_zip_entry(&mut zip, options, part.as_str(), raw.bytes.clone())?;
            } else {
                let mut table_to_write = table.clone();
                table_to_write.id = assigned_id;
                if table_to_write.name.is_empty() {
                    table_to_write.name = table_to_write.display_name.clone();
                }
                let bytes = crate::table::serialize_table_xml(&table_to_write)?;
                write_zip_entry(&mut zip, options, part.as_str(), bytes)?;
            }
            written_entries.insert(part);
        }

        if has_comments(ws) || !sheet_part_numbers.is_empty() || has_external_hyperlinks {
            let sheet_rels_xml = sheet_rels::write_sheet_rels(&sheet_rels_entries)?;
            let rels_part = format!("xl/worksheets/_rels/sheet{}.xml.rels", idx + 1);
            write_zip_entry(&mut zip, options, rels_part.as_str(), sheet_rels_xml)?;
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

    // Table parts are always re-emitted from the in-memory model, so the writer
    // owns them; never blindly copy the originals (which would duplicate or
    // orphan them after re-numbering).
    if name.starts_with("xl/tables/") {
        return true;
    }

    (1..=sheet_count).any(|sheet_idx| name == format!("xl/worksheets/sheet{sheet_idx}.xml"))
}

/// Return `Err` if two tables anywhere in the workbook share a `display_name` (case-insensitive).
fn ensure_unique_table_display_names(sheets: &[Worksheet]) -> Result<(), OxlsxError> {
    let mut seen = HashSet::new();
    for ws in sheets {
        for table in ws.tables() {
            // FIX 6: Case-insensitive comparison (Excel requirement).
            let lower = table.display_name.to_ascii_lowercase();
            if !seen.insert(lower) {
                return Err(OxlsxError::DuplicateTableName(table.display_name.clone()));
            }
        }
    }
    Ok(())
}

/// Allocate the internal `<table id>` for every table, workbook-wide.
///
/// Clean retained (verbatim) tables keep their original parsed id — the writer
/// re-emits their bytes untouched, so the id baked into those bytes must be
/// honored. Model-path tables (dirty or freshly added) receive newly allocated
/// ids that avoid every verbatim id, preserving workbook-wide uniqueness.
///
/// `result[sheet_idx][table_idx]` is `(assigned_id, is_verbatim)` in declared
/// table order. The id is independent of the part-filename sequence.
fn assign_table_ids(sheets: &[Worksheet]) -> Vec<Vec<(u32, bool)>> {
    // Verbatim ids are fixed (baked into retained bytes); reserve them all
    // before allocating ids for model-path tables.
    let mut used: HashSet<u32> = HashSet::new();
    for ws in sheets {
        for (idx, table) in ws.tables().iter().enumerate() {
            if is_clean_retained(ws, idx) {
                used.insert(table.id);
            }
        }
    }

    let mut next_free = 1u32;
    sheets
        .iter()
        .map(|ws| {
            ws.tables()
                .iter()
                .enumerate()
                .map(|(idx, table)| {
                    if is_clean_retained(ws, idx) {
                        // Verbatim: keep the original id from the retained bytes
                        // (already parsed into `table.id` at read time).
                        (table.id, true)
                    } else {
                        while used.contains(&next_free) {
                            next_free += 1;
                        }
                        let id = next_free;
                        used.insert(id);
                        next_free += 1;
                        (id, false)
                    }
                })
                .collect()
        })
        .collect()
}

/// Whether the table at `idx` has retained source bytes AND is not dirty, i.e.
/// it should be re-emitted verbatim rather than serialized from the model.
fn is_clean_retained(ws: &Worksheet, idx: usize) -> bool {
    matches!(ws.table_retention.get(idx), Some(Some(raw)) if !raw.dirty)
}

fn write_content_types(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    sheet_count: usize,
    has_shared_strings: bool,
    commented_sheets: &[usize],
    total_tables: usize,
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
    if !commented_sheets.is_empty() {
        write_empty(
            &mut xml,
            "Default",
            &[
                ("Extension", "vml"),
                (
                    "ContentType",
                    "application/vnd.openxmlformats-officedocument.vmlDrawing",
                ),
            ],
        )?;
    }
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
    for sheet_idx in commented_sheets {
        let part_name = format!("/xl/comments{sheet_idx}.xml");
        write_empty(
            &mut xml,
            "Override",
            &[
                ("PartName", part_name.as_str()),
                (
                    "ContentType",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
                ),
            ],
        )?;
    }
    for number in 1..=total_tables {
        let part_name = format!("/xl/tables/table{number}.xml");
        write_empty(
            &mut xml,
            "Override",
            &[
                ("PartName", part_name.as_str()),
                (
                    "ContentType",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
                ),
            ],
        )?;
    }

    write_end(&mut xml, "Types")?;
    write_zip_entry(zip, options, "[Content_Types].xml", xml.into_inner())
}

fn write_content_types_from_source(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    source: &Path,
    commented_sheets: &[usize],
    total_tables: usize,
    has_shared_strings: bool,
) -> Result<(), OxlsxError> {
    let source_file = File::open(source)?;
    let mut archive = ZipArchive::new(source_file)?;
    let mut xml = String::new();
    archive
        .by_name("[Content_Types].xml")?
        .read_to_string(&mut xml)?;

    // FIX 1: Strip ALL existing `/xl/tables/tableN.xml` overrides before appending
    // the writer-owned ones. The writer owns all table parts (see `is_writer_owned_entry`),
    // so we renumber them to `1..=total_tables` and discard stale indices.
    // Support BOTH whitespace formats: `"/>` (oxlsx-written) and `" />` (openpyxl-written).
    for n in 1..=1000 {
        // Format WITHOUT space before "/>":
        let pattern_no_space = format!(
            r#"<Override PartName="/xl/tables/table{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>"#
        );
        xml = xml.replace(&pattern_no_space, "");

        // Format WITH space before "/>":
        let pattern_with_space = format!(
            r#"<Override PartName="/xl/tables/table{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml" />"#
        );
        xml = xml.replace(&pattern_with_space, "");
    }

    if !xml.contains("Extension=\"vml\"") && !commented_sheets.is_empty() {
        xml = xml.replace(
            "</Types>",
            "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/></Types>",
        );
    }

    for sheet_idx in commented_sheets {
        let part_name = format!("/xl/comments{sheet_idx}.xml");
        if xml.contains(part_name.as_str()) {
            continue;
        }

        let override_xml = format!(
            "<Override PartName=\"{part_name}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>"
        );
        xml = xml.replace("</Types>", format!("{override_xml}</Types>").as_str());
    }

    // FIX 8: Accumulate ALL table overrides BEFORE doing a single replace (O(N) not O(N^2)).
    let mut table_overrides = String::new();
    for number in 1..=total_tables {
        let part_name = format!("/xl/tables/table{number}.xml");
        let override_xml = format!(
            "<Override PartName=\"{part_name}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>"
        );
        table_overrides.push_str(&override_xml);
    }
    if !table_overrides.is_empty() {
        xml = xml.replace("</Types>", format!("{table_overrides}</Types>").as_str());
    }

    // CRITICAL FIX: Inject sharedStrings Override when writer creates xl/sharedStrings.xml
    // but source manifest lacks it (e.g., RMW of openpyxl-authored file with inline strings).
    if has_shared_strings && !xml.contains("/xl/sharedStrings.xml") {
        let shared_strings_override = r#"<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#;
        xml = xml.replace(
            "</Types>",
            format!("{shared_strings_override}</Types>").as_str(),
        );
    }

    write_zip_entry(zip, options, "[Content_Types].xml", xml.into_bytes())
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
    defined_names::write_defined_names_block(
        &mut xml,
        wb.worksheets(),
        wb.defined_names(),
        &wb.extra_defined_names,
    )?;
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

/// Relationship references that influence the sheet XML body: the legacy VML
/// drawing rId (for comments) and the table part rIds (for `<tableParts>`).
struct SheetRelRefs<'a> {
    legacy_drawing_rid: Option<&'a str>,
    table_rids: &'a [String],
    hyperlinks: &'a [HyperlinkRef<'a>],
}

/// A single `<hyperlink>` to emit, in document order. External links carry a
/// resolved `rel_id` (`r:id`); internal `location`-only links carry `None`.
struct HyperlinkRef<'a> {
    reference: &'a str,
    hyperlink: &'a Hyperlink,
    rel_id: Option<&'a str>,
}

fn write_sheet_xml(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    ws: &Worksheet,
    sheet_idx: usize,
    style_ids: &HashMap<(usize, String), u32>,
    shared_strings: &IndexMap<String, usize>,
    rel_refs: &SheetRelRefs<'_>,
) -> Result<(), OxlsxError> {
    let SheetRelRefs {
        legacy_drawing_rid,
        table_rids,
        hyperlinks,
    } = *rel_refs;
    let mut xml = xml_writer_with_decl()?;
    let mut worksheet_attrs = vec![(
        "xmlns",
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    )];
    let needs_rel_namespace = legacy_drawing_rid.is_some()
        || !table_rids.is_empty()
        || hyperlinks.iter().any(|link| link.rel_id.is_some());
    if needs_rel_namespace {
        worksheet_attrs.push((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
    }
    write_start(&mut xml, "worksheet", worksheet_attrs.as_slice())?;

    if let Some(coord) = ws.freeze_panes() {
        write_freeze_panes_view(&mut xml, coord)?;
    }

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

    if !ws.merged_cells.is_empty() {
        let count = ws.merged_cells.len().to_string();
        write_start(&mut xml, "mergeCells", &[("count", count.as_str())])?;
        for range_ref in &ws.merged_cells {
            write_empty(&mut xml, "mergeCell", &[("ref", range_ref.as_str())])?;
        }
        write_end(&mut xml, "mergeCells")?;
    }

    // `<hyperlinks>` MUST follow `<mergeCells>` and precede `<pageMargins>` /
    // `<legacyDrawing>` per CT_Worksheet ordering, or Excel flags the file for
    // repair.
    if !hyperlinks.is_empty() {
        write_start(&mut xml, "hyperlinks", &[])?;
        for link in hyperlinks {
            let mut attrs: Vec<(&str, &str)> = vec![("ref", link.reference)];
            if let Some(rel_id) = link.rel_id {
                attrs.push(("r:id", rel_id));
            }
            if let Some(location) = link.hyperlink.location.as_deref() {
                attrs.push(("location", location));
            }
            if let Some(tooltip) = link.hyperlink.tooltip.as_deref() {
                attrs.push(("tooltip", tooltip));
            }
            if let Some(display) = link.hyperlink.display.as_deref() {
                attrs.push(("display", display));
            }
            write_empty(&mut xml, "hyperlink", &attrs)?;
        }
        write_end(&mut xml, "hyperlinks")?;
    }

    if let Some(rel_id) = legacy_drawing_rid {
        write_empty(&mut xml, "legacyDrawing", &[("r:id", rel_id)])?;
    }

    // `<tableParts>` MUST be the last child of `<worksheet>` (after legacyDrawing
    // per the OOXML schema), or Excel flags the file for repair.
    if !table_rids.is_empty() {
        let count = table_rids.len().to_string();
        write_start(&mut xml, "tableParts", &[("count", count.as_str())])?;
        for rid in table_rids {
            write_empty(&mut xml, "tablePart", &[("r:id", rid.as_str())])?;
        }
        write_end(&mut xml, "tableParts")?;
    }

    write_end(&mut xml, "worksheet")?;
    let part = format!("xl/worksheets/sheet{}.xml", sheet_idx + 1);
    write_zip_entry(zip, options, part.as_str(), xml.into_inner())
}

fn has_comments(ws: &Worksheet) -> bool {
    ws.cells().values().any(|cell| cell.comment().is_some())
}

fn write_freeze_panes_view(xml: &mut Writer<Vec<u8>>, coord: &str) -> Result<(), OxlsxError> {
    let Some((row, col)) = parse_a1_coord(coord) else {
        return Ok(());
    };

    let x_split = col.saturating_sub(1);
    let y_split = row.saturating_sub(1);
    if x_split == 0 && y_split == 0 {
        return Ok(());
    }

    write_start(xml, "sheetViews", &[])?;
    write_start(xml, "sheetView", &[("workbookViewId", "0")])?;

    let active_pane = if x_split > 0 && y_split > 0 {
        "bottomRight"
    } else if y_split > 0 {
        "bottomLeft"
    } else {
        "topRight"
    };

    let mut pane_attrs = vec![
        ("topLeftCell", coord.to_string()),
        ("activePane", active_pane.to_string()),
        ("state", "frozen".to_string()),
    ];
    if x_split > 0 {
        pane_attrs.push(("xSplit", x_split.to_string()));
    }
    if y_split > 0 {
        pane_attrs.push(("ySplit", y_split.to_string()));
    }

    let pane_ref = pane_attrs
        .iter()
        .map(|(k, v)| (*k, v.as_str()))
        .collect::<Vec<_>>();
    write_empty(xml, "pane", pane_ref.as_slice())?;

    if x_split > 0 && y_split > 0 {
        write_empty(xml, "selection", &[("pane", "topRight")])?;
        write_empty(xml, "selection", &[("pane", "bottomLeft")])?;
        write_empty(
            xml,
            "selection",
            &[
                ("pane", "bottomRight"),
                ("activeCell", coord),
                ("sqref", coord),
            ],
        )?;
    } else {
        write_empty(
            xml,
            "selection",
            &[
                ("pane", active_pane),
                ("activeCell", coord),
                ("sqref", coord),
            ],
        )?;
    }

    write_end(xml, "sheetView")?;
    write_end(xml, "sheetViews")
}

fn parse_a1_coord(reference: &str) -> Option<(u32, u32)> {
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

    let row = row_text.parse::<u32>().ok()?;
    if row == 0 || col == 0 {
        return None;
    }

    Some((row, col))
}

fn sorted_comments(ws: &Worksheet) -> Vec<(&str, &Comment)> {
    let mut comments = ws
        .cells()
        .iter()
        .filter_map(|(reference, cell)| cell.comment().map(|comment| (reference.as_str(), comment)))
        .collect::<Vec<_>>();
    comments.sort_by_key(|(reference, _)| cell_sort_key(reference));
    comments
}

fn write_comments_xml(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    ws: &Worksheet,
    sheet_idx: usize,
) -> Result<(), OxlsxError> {
    let comments = sorted_comments(ws);
    if comments.is_empty() {
        return Ok(());
    }

    let mut authors = Vec::<String>::new();
    let mut author_ids = HashMap::<String, usize>::new();
    for (_, comment) in &comments {
        if !author_ids.contains_key(comment.author.as_str()) {
            let id = authors.len();
            authors.push(comment.author.clone());
            author_ids.insert(comment.author.clone(), id);
        }
    }

    let mut xml = xml_writer_with_decl()?;
    write_start(
        &mut xml,
        "comments",
        &[(
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        )],
    )?;

    write_start(&mut xml, "authors", &[])?;
    for author in &authors {
        write_start(&mut xml, "author", &[])?;
        xml.write_event(Event::Text(BytesText::new(author)))?;
        write_end(&mut xml, "author")?;
    }
    write_end(&mut xml, "authors")?;

    write_start(&mut xml, "commentList", &[])?;
    for (reference, comment) in comments {
        let author_id = author_ids
            .get(comment.author.as_str())
            .copied()
            .unwrap_or(0)
            .to_string();
        write_start(
            &mut xml,
            "comment",
            &[("ref", reference), ("authorId", author_id.as_str())],
        )?;
        write_start(&mut xml, "text", &[])?;
        write_start(&mut xml, "t", &[])?;
        xml.write_event(Event::Text(BytesText::new(comment.text.as_str())))?;
        write_end(&mut xml, "t")?;
        write_end(&mut xml, "text")?;
        write_end(&mut xml, "comment")?;
    }
    write_end(&mut xml, "commentList")?;
    write_end(&mut xml, "comments")?;

    let part = format!("xl/comments{sheet_idx}.xml");
    write_zip_entry(zip, options, part.as_str(), xml.into_inner())
}

fn write_vml_drawing(
    zip: &mut ZipWriter<File>,
    options: SimpleFileOptions,
    ws: &Worksheet,
    sheet_idx: usize,
) -> Result<(), OxlsxError> {
    let comments = sorted_comments(ws);
    if comments.is_empty() {
        return Ok(());
    }

    let mut shapes = String::new();
    for (idx, (reference, _)) in comments.iter().enumerate() {
        let (row, col) = cell_sort_key(reference);
        shapes.push_str(
            format!(
                r##"<v:shape id="_x0000_s{}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>1, 15, 0, 2, 3, 15, 4, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{}</x:Row><x:Column>{}</x:Column></x:ClientData></v:shape>"##,
                1024 + idx,
                row.saturating_sub(1),
                col.saturating_sub(1)
            )
            .as_str(),
        );
    }

    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
    <v:stroke joinstyle="miter"/>
    <v:path gradientshapeok="t" o:connecttype="rect"/>
  </v:shapetype>
  {shapes}
</xml>"#
    );

    let part = format!("xl/drawings/vmlDrawing{sheet_idx}.vml");
    write_zip_entry(zip, options, part.as_str(), xml.into_bytes())
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

    #[test]
    fn assign_table_ids_avoids_verbatim_collision() {
        use crate::Table;
        use crate::table::RawTable;

        let mut ws = Worksheet::new("Sheet");
        // One verbatim (clean retained) table whose internal id is 5.
        ws.add_table(Table::new("Verbatim", "A1:B2")).unwrap();
        ws.tables[0].id = 5;
        ws.table_retention[0] = Some(RawTable {
            bytes: b"<table id=\"5\"/>".to_vec(),
            dirty: false,
        });
        // One model table (no retained bytes) that must avoid id 5.
        ws.add_table(Table::new("Fresh", "D1:E2")).unwrap();

        let ids = super::assign_table_ids(&[ws]);
        let sheet = &ids[0];
        assert_eq!(sheet.len(), 2);

        // Verbatim keeps its original id and is flagged verbatim.
        assert_eq!(sheet[0], (5, true));

        // Model table is allocated, distinct, and never collides with 5.
        let (model_id, is_verbatim) = sheet[1];
        assert!(!is_verbatim);
        assert_ne!(model_id, 5);
        assert_ne!(model_id, sheet[0].0, "ids must be unique workbook-wide");
    }

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
