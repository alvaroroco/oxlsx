use chrono::NaiveDate;
use oxlsx::{
    Color, Comment, Fill, Font, OxlsxError, StyleId, Table, TableStyleInfo, Value, Workbook,
};
use std::env;
use std::fs;
use std::io::Read;
use std::io::Write;
use zip::ZipArchive;
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

fn temp_path(name: &str) -> std::path::PathBuf {
    env::temp_dir().join(name)
}

fn read_zip_entry(path: &std::path::Path, entry: &str) -> String {
    let file = fs::File::open(path).unwrap();
    let mut zip = ZipArchive::new(file).unwrap();
    let mut xml = String::new();
    zip.by_name(entry)
        .unwrap()
        .read_to_string(&mut xml)
        .unwrap();
    xml
}

fn attr_in(fragment: &str, attr: &str) -> Option<String> {
    let needle = format!("{attr}=\"");
    let start = fragment.find(&needle)? + needle.len();
    let rest = &fragment[start..];
    let end = rest.find('"')?;
    Some(rest[..end].to_string())
}

/// Every `Id` attribute across all `<Relationship>` elements, in document order.
fn relationship_ids(rels_xml: &str) -> Vec<String> {
    rels_xml
        .split("<Relationship")
        .skip(1)
        .filter_map(|fragment| attr_in(fragment, "Id"))
        .collect()
}

/// The `Id` of the first `<Relationship>` whose `Type` contains `type_suffix`.
fn rel_id_for_type(rels_xml: &str, type_suffix: &str) -> String {
    rels_xml
        .split("<Relationship")
        .skip(1)
        .find(|fragment| {
            attr_in(fragment, "Type")
                .map(|t| t.contains(type_suffix))
                .unwrap_or(false)
        })
        .and_then(|fragment| attr_in(fragment, "Id"))
        .unwrap_or_else(|| panic!("no relationship of type {type_suffix} in {rels_xml}"))
}

fn inject_unrelated_sheet_rel(path: &std::path::Path) {
    let source_file = fs::File::open(path).unwrap();
    let mut source_zip = ZipArchive::new(source_file).unwrap();
    let temp_path = path.with_extension("tmp.xlsx");
    let temp_file = fs::File::create(&temp_path).unwrap();
    let mut writer = ZipWriter::new(temp_file);
    let options = SimpleFileOptions::default();

    for idx in 0..source_zip.len() {
        let mut entry = source_zip.by_index(idx).unwrap();
        let name = entry.name().to_string();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();

        if name == "xl/worksheets/_rels/sheet1.xml.rels" {
            let xml = String::from_utf8(bytes).unwrap();
            let injected = xml.replace(
                "</Relationships>",
                "<Relationship Id=\"rId99\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings\" Target=\"../printerSettings/printerSettings1.bin\"/></Relationships>",
            );
            bytes = injected.into_bytes();
        }

        writer.start_file(name, options).unwrap();
        writer.write_all(&bytes).unwrap();
    }

    writer.finish().unwrap();
    fs::rename(temp_path, path).unwrap();
}

fn inject_sheet_freeze_pane(path: &std::path::Path, coord: &str) {
    let source_file = fs::File::open(path).unwrap();
    let mut source_zip = ZipArchive::new(source_file).unwrap();
    let temp_path = path.with_extension("tmp.xlsx");
    let temp_file = fs::File::create(&temp_path).unwrap();
    let mut writer = ZipWriter::new(temp_file);
    let options = SimpleFileOptions::default();

    for idx in 0..source_zip.len() {
        let mut entry = source_zip.by_index(idx).unwrap();
        let name = entry.name().to_string();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();

        if name == "xl/worksheets/sheet1.xml" {
            let xml = String::from_utf8(bytes).unwrap();
            let frozen_xml = xml.replace(
                "<sheetData>",
                format!(
                    "<sheetViews><sheetView workbookViewId=\"0\"><pane xSplit=\"1\" ySplit=\"1\" topLeftCell=\"{coord}\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"topRight\"/><selection pane=\"bottomLeft\"/><selection pane=\"bottomRight\" activeCell=\"{coord}\" sqref=\"{coord}\"/></sheetView></sheetViews><sheetData>"
                )
                .as_str(),
            );
            bytes = frozen_xml.into_bytes();
        }

        writer.start_file(name, options).unwrap();
        writer.write_all(&bytes).unwrap();
    }

    writer.finish().unwrap();
    fs::rename(temp_path, path).unwrap();
}

fn inject_workbook_defined_names(path: &std::path::Path, defined_names_xml: &str) {
    let source_file = fs::File::open(path).unwrap();
    let mut source_zip = ZipArchive::new(source_file).unwrap();
    let temp_path = path.with_extension("tmp.xlsx");
    let temp_file = fs::File::create(&temp_path).unwrap();
    let mut writer = ZipWriter::new(temp_file);
    let options = SimpleFileOptions::default();

    for idx in 0..source_zip.len() {
        let mut entry = source_zip.by_index(idx).unwrap();
        let name = entry.name().to_string();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();

        if name == "xl/workbook.xml" {
            let xml = String::from_utf8(bytes).unwrap();
            let injected = xml.replace(
                "</workbook>",
                format!("{defined_names_xml}</workbook>").as_str(),
            );
            bytes = injected.into_bytes();
        }

        writer.start_file(name, options).unwrap();
        writer.write_all(&bytes).unwrap();
    }

    writer.finish().unwrap();
    fs::rename(temp_path, path).unwrap();
}

#[test]
fn phase3_multi_sheet_workbook() {
    let wb = Workbook::open("samples/multi_sheet.xlsx").unwrap();

    assert_eq!(wb.worksheets().len(), 2);
    assert_eq!(wb.sheetnames(), vec!["Empleados", "Productos"]);
    assert_eq!(wb.sheet(0).unwrap().title(), "Empleados");
    assert_eq!(wb.sheet(1).unwrap().title(), "Productos");
    assert!(wb.sheet_by_name("Empleados").is_some());
    assert!(wb.sheet_by_name("NoExiste").is_none());
    assert_eq!(wb.active().unwrap().title(), "Empleados");

    let ws0 = wb.sheet(0).unwrap();
    assert_eq!(
        ws0.cell("A1").unwrap().value(),
        &Value::String("Juan".to_string())
    );
    assert_eq!(
        ws0.cell("B1").unwrap().value(),
        &Value::Date(NaiveDate::from_ymd_opt(2025, 1, 21).unwrap())
    );
    assert!(ws0.cell("B1").unwrap().text().is_none());

    let ws1 = wb.sheet(1).unwrap();
    assert_eq!(
        ws1.cell("A1").unwrap().value(),
        &Value::String("Laptop".to_string())
    );
}

#[test]
fn phase3_regression_test_xlsx() {
    let wb = Workbook::open("samples/test.xlsx").unwrap();
    assert_eq!(wb.sheet(0).unwrap().title(), "Hoja1");
    assert_eq!(wb.sheetnames(), vec!["Hoja1"]);

    let ws = wb.sheet(0).unwrap();
    let a1 = ws.cell("A1").unwrap();
    assert_eq!(a1.style_id(), StyleId(1));
    assert_eq!(a1.value(), &Value::String("Hola".to_string()));
    assert!(a1.font().unwrap().bold);
    assert_eq!(a1.fill().unwrap().pattern_type, "solid");
    assert_eq!(a1.fill().unwrap().fg_color, Some(Color::Indexed(2)));

    let b1 = ws.cell("B1").unwrap();
    assert_eq!(b1.style_id(), StyleId(2));
    assert_eq!(b1.value(), &Value::String("Hola".to_string()));
    assert_eq!(b1.font().unwrap().color, Some(Color::Indexed(4)));
    assert_eq!(b1.fill().unwrap().fg_color, Some(Color::Indexed(5)));

    let c1 = ws.cell("C1").unwrap();
    assert_eq!(c1.style_id(), StyleId(0));
    assert_eq!(c1.value(), &Value::Number(42.0));

    assert!(ws.cell("D1").is_none());
    assert!(ws.cell("ZZ999").is_none());
}

#[test]
fn phase4_round_trip_string() {
    let path = temp_path("oxlsx_rt_string.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_cell("A1", "Hello");
    wb.save(&path).unwrap();
    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(0).unwrap().cell("A1").unwrap().value(),
        &Value::String("Hello".to_string())
    );
    std::fs::remove_file(&path).ok();
}

#[test]
fn phase4_round_trip_number() {
    let path = temp_path("oxlsx_rt_number.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_cell("B1", 42.0_f64);
    wb.save(&path).unwrap();
    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(0).unwrap().cell("B1").unwrap().value(),
        &Value::Number(42.0)
    );
    std::fs::remove_file(&path).ok();
}

#[test]
fn phase4_round_trip_bool() {
    let path = temp_path("oxlsx_rt_bool.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_cell("C1", true);
    wb.save(&path).unwrap();
    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(0).unwrap().cell("C1").unwrap().value(),
        &Value::Bool(true)
    );
    std::fs::remove_file(&path).ok();
}

#[test]
fn phase4_round_trip_bold_font() {
    let path = temp_path("oxlsx_rt_font.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.active_mut().unwrap();
    ws.set_cell("A1", "Bold");
    ws.set_cell_font(
        "A1",
        Font {
            bold: true,
            ..Default::default()
        },
    );
    wb.save(&path).unwrap();
    let wb2 = Workbook::open(&path).unwrap();
    assert!(
        wb2.sheet(0)
            .unwrap()
            .cell("A1")
            .unwrap()
            .font()
            .unwrap()
            .bold
    );
    std::fs::remove_file(&path).ok();
}

#[test]
fn phase4_round_trip_fill() {
    let path = temp_path("oxlsx_rt_fill.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.active_mut().unwrap();
    ws.set_cell("A1", "Filled");
    ws.set_cell_fill(
        "A1",
        Fill {
            pattern_type: "solid".to_string(),
            fg_color: Some(Color::Argb("FFFF00".to_string())),
            bg_color: None,
        },
    );
    wb.save(&path).unwrap();
    let wb2 = Workbook::open(&path).unwrap();
    let fill = wb2.sheet(0).unwrap().cell("A1").unwrap().fill().unwrap();
    assert_eq!(fill.pattern_type, "solid");
    assert_eq!(fill.fg_color, Some(Color::Argb("FFFF00".to_string())));
    std::fs::remove_file(&path).ok();
}

#[test]
fn phase4_round_trip_multi_sheet() {
    let path = temp_path("oxlsx_rt_multi.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_cell("A1", "Sheet1");
    wb.create_sheet("Data").set_cell("A1", "Sheet2");
    wb.save(&path).unwrap();
    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(wb2.sheetnames(), vec!["Sheet", "Data"]);
    assert_eq!(
        wb2.sheet(0).unwrap().cell("A1").unwrap().value(),
        &Value::String("Sheet1".to_string())
    );
    assert_eq!(
        wb2.sheet(1).unwrap().cell("A1").unwrap().value(),
        &Value::String("Sheet2".to_string())
    );
    std::fs::remove_file(&path).ok();
}

#[test]
fn readonly_open_returns_ok() {
    let wb = Workbook::open_readonly("samples/test.xlsx");
    assert!(wb.is_ok());
}

#[test]
fn readonly_sheetnames() {
    let wb = Workbook::open_readonly("samples/multi_sheet.xlsx").unwrap();
    assert_eq!(wb.sheetnames(), vec!["Empleados", "Productos"]);
}

#[test]
fn readonly_active_title() {
    let wb = Workbook::open_readonly("samples/multi_sheet.xlsx").unwrap();
    assert_eq!(wb.active().unwrap().title(), "Empleados");
}

#[test]
fn readonly_sheet_by_index() {
    let wb = Workbook::open_readonly("samples/multi_sheet.xlsx").unwrap();
    assert_eq!(wb.sheet(1).unwrap().title(), "Productos");
}

#[test]
fn readonly_rows_a1_value() {
    let wb = Workbook::open_readonly("samples/test.xlsx").unwrap();
    let ws = wb.sheet(0).unwrap();
    let first_row = ws.rows().next().unwrap();
    assert_eq!(first_row[0].value(), &Value::String("Hola".to_string()));
}

#[test]
fn readonly_rows_c1_value() {
    let wb = Workbook::open_readonly("samples/test.xlsx").unwrap();
    let ws = wb.sheet(0).unwrap();
    let first_row = ws.rows().next().unwrap();
    assert_eq!(first_row[2].value(), &Value::Number(42.0));
}

#[test]
fn readonly_rows_multi_sheet() {
    let wb = Workbook::open_readonly("samples/multi_sheet.xlsx").unwrap();
    let ws = wb.sheet(1).unwrap();
    let first_row = ws.rows().next().unwrap();
    assert_eq!(first_row[0].value(), &Value::String("Laptop".to_string()));
}

#[test]
fn readonly_independent_iterators() {
    let wb = Workbook::open_readonly("samples/test.xlsx").unwrap();
    let ws = wb.sheet(0).unwrap();

    let mut rows_a = ws.rows();
    let mut rows_b = ws.rows();

    let row_a = rows_a.next().unwrap();
    let row_b = rows_b.next().unwrap();

    assert_eq!(row_a[0].value(), row_b[0].value());
    assert_eq!(row_a[2].value(), row_b[2].value());
}

#[test]
fn readonly_streaming_style_resolution() {
    let wb = Workbook::open_readonly("samples/test.xlsx").unwrap();
    let ws = wb.sheet(0).unwrap();
    let first_row = ws.rows().next().unwrap();

    assert!(first_row[0].font().unwrap().bold);
    assert_eq!(
        first_row[0].fill().unwrap().fg_color,
        Some(Color::Indexed(2))
    );
    assert_eq!(first_row[1].font().unwrap().color, Some(Color::Indexed(4)));
    assert_eq!(
        first_row[1].fill().unwrap().fg_color,
        Some(Color::Indexed(5))
    );
}

#[test]
fn rmw_modify_string_cell() {
    let path = temp_path("oxlsx_rmw_modify_string.xlsx");
    let mut wb = Workbook::open("samples/test.xlsx").unwrap();
    wb.active_mut().unwrap().set_cell("A1", "Modified");
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(0).unwrap().cell("A1").unwrap().value(),
        &Value::String("Modified".to_string())
    );

    fs::remove_file(path).ok();
}

#[test]
fn rmw_modify_number_cell() {
    let path = temp_path("oxlsx_rmw_modify_number.xlsx");
    let mut wb = Workbook::open("samples/test.xlsx").unwrap();
    wb.active_mut().unwrap().set_cell("C1", 99.0_f64);
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(0).unwrap().cell("C1").unwrap().value(),
        &Value::Number(99.0)
    );

    fs::remove_file(path).ok();
}

#[test]
fn rmw_add_new_cell() {
    let path = temp_path("oxlsx_rmw_add_new_cell.xlsx");
    let mut wb = Workbook::open("samples/test.xlsx").unwrap();
    wb.active_mut().unwrap().set_cell("D1", "New");
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(0).unwrap().cell("D1").unwrap().value(),
        &Value::String("New".to_string())
    );

    fs::remove_file(path).ok();
}

#[test]
fn rmw_preserves_other_cells() {
    let path = temp_path("oxlsx_rmw_preserve_other_cells.xlsx");
    let mut wb = Workbook::open("samples/test.xlsx").unwrap();
    wb.active_mut().unwrap().set_cell("A1", "Changed");
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(0).unwrap().cell("B1").unwrap().value(),
        &Value::String("Hola".to_string())
    );

    fs::remove_file(path).ok();
}

#[test]
fn rmw_preserves_unmodified_styles() {
    let path = temp_path("oxlsx_rmw_preserve_styles.xlsx");
    let mut wb = Workbook::open("samples/test.xlsx").unwrap();
    wb.active_mut().unwrap().set_cell("C1", 99.0_f64);
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert!(
        wb2.sheet(0)
            .unwrap()
            .cell("A1")
            .unwrap()
            .font()
            .unwrap()
            .bold
    );

    fs::remove_file(path).ok();
}

#[test]
fn rmw_multi_sheet_modify() {
    let path = temp_path("oxlsx_rmw_multi_sheet.xlsx");
    let mut wb = Workbook::open("samples/multi_sheet.xlsx").unwrap();
    wb.sheet_mut(0).unwrap().set_cell("A1", "Pedro");
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(1).unwrap().cell("A1").unwrap().value(),
        &Value::String("Laptop".to_string())
    );

    fs::remove_file(path).ok();
}

#[test]
fn rmw_save_to_same_path() {
    let path = temp_path("oxlsx_rmw_same_path.xlsx");
    fs::copy("samples/test.xlsx", &path).unwrap();

    let mut wb = Workbook::open(&path).unwrap();
    wb.active_mut().unwrap().set_cell("A1", "SamePath");
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(0).unwrap().cell("A1").unwrap().value(),
        &Value::String("SamePath".to_string())
    );

    fs::remove_file(path).ok();
}

#[test]
fn write_formula_cell() {
    let path = temp_path("oxlsx_write_formula.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut()
        .unwrap()
        .set_cell("A1", Value::Formula("SUM(B1:B3)".to_string()));
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(0).unwrap().cell("A1").unwrap().value(),
        &Value::Formula("SUM(B1:B3)".to_string())
    );

    fs::remove_file(path).ok();
}

#[test]
fn write_date_with_num_fmt() {
    let path = temp_path("oxlsx_write_date_num_fmt.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_cell(
        "A1",
        Value::Date(NaiveDate::from_ymd_opt(2025, 1, 21).unwrap()),
    );
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(0).unwrap().cell("A1").unwrap().value(),
        &Value::Date(NaiveDate::from_ymd_opt(2025, 1, 21).unwrap())
    );

    fs::remove_file(path).ok();
}

#[test]
fn write_column_width() {
    let path = temp_path("oxlsx_write_column_width.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().column_width(1, 20.0);
    wb.save(&path).unwrap();

    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(sheet_xml.contains("customWidth=\"1\""));

    fs::remove_file(path).ok();
}

#[test]
fn write_row_height() {
    let path = temp_path("oxlsx_write_row_height.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_cell("A1", "x");
    wb.active_mut().unwrap().row_height(1, 30.0);
    wb.save(&path).unwrap();

    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(sheet_xml.contains("customHeight=\"1\""));

    fs::remove_file(path).ok();
}

#[test]
fn write_merge_cells() {
    let path = temp_path("oxlsx_write_merge_cells.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().merge_cells("A1:C1");
    wb.save(&path).unwrap();

    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(sheet_xml.contains("<mergeCell ref=\"A1:C1\"/>"));

    fs::remove_file(path).ok();
}

#[test]
fn write_unmerge_cells() {
    let path = temp_path("oxlsx_write_unmerge_cells.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.active_mut().unwrap();
    ws.merge_cells("A1:C1");
    ws.unmerge_cells("A1:C1");
    wb.save(&path).unwrap();

    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(!sheet_xml.contains("mergeCell"));

    fs::remove_file(path).ok();
}

#[test]
fn comment_round_trip() {
    let path = temp_path("oxlsx_comment_round_trip.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.active_mut().unwrap();
    ws.set_cell("B2", "Cell with comment");
    ws.set_cell_comment(
        "B2",
        Comment {
            text: "Hello from comment".to_string(),
            author: "Alice".to_string(),
        },
    );

    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    let comment = wb2.active().unwrap().cell("B2").unwrap().comment().unwrap();
    assert_eq!(comment.text, "Hello from comment");
    assert_eq!(comment.author, "Alice");

    fs::remove_file(path).ok();
}

#[test]
fn comment_write_new_cell() {
    let path = temp_path("oxlsx_comment_write_new_cell.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_cell_comment(
        "C3",
        Comment {
            text: "Newly added".to_string(),
            author: "Bob".to_string(),
        },
    );

    wb.save(&path).unwrap();

    let sheet_rels = read_zip_entry(&path, "xl/worksheets/_rels/sheet1.xml.rels");
    assert!(sheet_rels.contains("/relationships/comments"));
    assert!(sheet_rels.contains("/relationships/vmlDrawing"));

    // Bug #274: on a from-scratch sheet the comment and vmlDrawing rels were
    // both allocated `rId1`, producing duplicate `@Id` (an xsd:ID) — an OPC
    // violation Excel may flag for repair. Every relationship Id must be unique.
    let rel_ids = relationship_ids(&sheet_rels);
    let mut unique = rel_ids.clone();
    unique.sort();
    unique.dedup();
    assert_eq!(
        unique.len(),
        rel_ids.len(),
        "duplicate relationship Id found in {sheet_rels}"
    );

    // `<legacyDrawing r:id=...>` must resolve to the vmlDrawing rel, NOT the
    // comments rel.
    let vml_id = rel_id_for_type(&sheet_rels, "/relationships/vmlDrawing");
    let comments_id = rel_id_for_type(&sheet_rels, "/relationships/comments");
    assert_ne!(vml_id, comments_id);
    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(
        sheet_xml.contains(&format!("<legacyDrawing r:id=\"{vml_id}\"")),
        "legacyDrawing must point to vml rel {vml_id} in {sheet_xml}"
    );

    let wb2 = Workbook::open(&path).unwrap();
    let comment = wb2.active().unwrap().cell("C3").unwrap().comment().unwrap();
    assert_eq!(comment.text, "Newly added");
    assert_eq!(comment.author, "Bob");

    fs::remove_file(path).ok();
}

#[test]
fn sheet_comment_relationships_preserved() {
    let seed_path = temp_path("oxlsx_comment_preserve_seed.xlsx");
    let out_path = temp_path("oxlsx_comment_preserve_out.xlsx");

    let mut seed = Workbook::new();
    seed.active_mut().unwrap().set_cell_comment(
        "A1",
        Comment {
            text: "Seed".to_string(),
            author: "Author".to_string(),
        },
    );
    seed.save(&seed_path).unwrap();
    inject_unrelated_sheet_rel(&seed_path);

    let mut wb = Workbook::open(&seed_path).unwrap();
    wb.active_mut().unwrap().set_cell_comment(
        "A1",
        Comment {
            text: "Updated".to_string(),
            author: "Author".to_string(),
        },
    );
    wb.save(&out_path).unwrap();

    let rels = read_zip_entry(&out_path, "xl/worksheets/_rels/sheet1.xml.rels");
    assert!(rels.contains("relationships/printerSettings"));
    assert!(rels.contains("relationships/comments"));
    assert!(rels.contains("relationships/vmlDrawing"));

    fs::remove_file(seed_path).ok();
    fs::remove_file(out_path).ok();
}

#[test]
fn phase9_empty_worksheet_bounds_and_rows_are_defined() {
    let wb = Workbook::new();
    let ws = wb.active().unwrap();

    assert_eq!(ws.min_row(), 1);
    assert_eq!(ws.max_row(), 1);
    assert_eq!(ws.min_column(), 1);
    assert_eq!(ws.max_column(), 1);

    let rows = ws.rows();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].len(), 1);
    assert_eq!(rows[0][0].value(), &Value::Empty);
}

#[test]
fn phase9_rows_columns_and_values_match_rectangular_iteration() {
    let wb = Workbook::open("samples/test.xlsx").unwrap();
    let ws = wb.active().unwrap();

    let rows = ws.rows();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].len(), 3);
    assert_eq!(rows[0][0].value(), &Value::String("Hola".to_string()));
    assert_eq!(rows[0][1].value(), &Value::String("Hola".to_string()));
    assert_eq!(rows[0][2].value(), &Value::Number(42.0));

    let columns = ws.columns();
    assert_eq!(columns.len(), 3);
    assert_eq!(columns[0].len(), 1);
    assert_eq!(columns[0][0].value(), &Value::String("Hola".to_string()));
    assert_eq!(columns[2][0].value(), &Value::Number(42.0));

    let values = ws.values();
    assert_eq!(
        values,
        vec![vec![
            Value::String("Hola".to_string()),
            Value::String("Hola".to_string()),
            Value::Number(42.0)
        ]]
    );
}

#[test]
fn write_shared_strings() {
    let path = temp_path("oxlsx_write_shared_strings.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.active_mut().unwrap();
    ws.set_cell("A1", "hello");
    ws.set_cell("A2", "hello");
    ws.set_cell("A3", "hello");
    wb.save(&path).unwrap();

    let shared_strings_xml = read_zip_entry(&path, "xl/sharedStrings.xml");
    let count = shared_strings_xml
        .match_indices("<si><t>hello</t></si>")
        .count();
    assert_eq!(count, 1);

    fs::remove_file(path).ok();
}

#[test]
fn write_shared_strings_index() {
    let path = temp_path("oxlsx_write_shared_strings_index.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.active_mut().unwrap();
    ws.set_cell("A1", "hello");
    ws.set_cell("B1", "world");
    wb.save(&path).unwrap();

    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(sheet_xml.contains("r=\"A1\""));
    assert!(sheet_xml.contains("r=\"B1\""));
    assert!(sheet_xml.contains("t=\"s\""));
    assert!(sheet_xml.contains("<v>0</v>"));
    assert!(sheet_xml.contains("<v>1</v>"));

    fs::remove_file(path).ok();
}

#[test]
fn freeze_panes_round_trip_existing_workbook() {
    let seeded_path = temp_path("oxlsx_freeze_seed.xlsx");
    let out_path = temp_path("oxlsx_freeze_roundtrip.xlsx");

    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    inject_sheet_freeze_pane(&seeded_path, "B2");

    let wb = Workbook::open(&seeded_path).unwrap();
    assert_eq!(wb.active().unwrap().freeze_panes(), Some("B2"));

    wb.save(&out_path).unwrap();
    let wb_reloaded = Workbook::open(&out_path).unwrap();
    assert_eq!(wb_reloaded.active().unwrap().freeze_panes(), Some("B2"));

    fs::remove_file(seeded_path).ok();
    fs::remove_file(out_path).ok();
}

#[test]
fn freeze_panes_write_new_workbook() {
    let path = temp_path("oxlsx_freeze_write_new.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_freeze_panes(Some("B2"));
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(wb2.active().unwrap().freeze_panes(), Some("B2"));

    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(sheet_xml.contains("<sheetViews>"));
    assert!(sheet_xml.contains("state=\"frozen\""));
    assert!(sheet_xml.contains("topLeftCell=\"B2\""));

    fs::remove_file(path).ok();
}

#[test]
fn freeze_panes_clear_removes_pane() {
    let seeded_path = temp_path("oxlsx_freeze_clear_seed.xlsx");
    let out_path = temp_path("oxlsx_freeze_clear_out.xlsx");

    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    inject_sheet_freeze_pane(&seeded_path, "B2");

    let mut wb = Workbook::open(&seeded_path).unwrap();
    assert_eq!(wb.active().unwrap().freeze_panes(), Some("B2"));
    wb.active_mut().unwrap().set_freeze_panes(None);
    wb.save(&out_path).unwrap();

    let wb2 = Workbook::open(&out_path).unwrap();
    assert_eq!(wb2.active().unwrap().freeze_panes(), None);

    let sheet_xml = read_zip_entry(&out_path, "xl/worksheets/sheet1.xml");
    assert!(!sheet_xml.contains("<pane"));

    fs::remove_file(seeded_path).ok();
    fs::remove_file(out_path).ok();
}

#[test]
fn print_area_round_trip_existing_workbook() {
    let seeded_path = temp_path("oxlsx_print_area_seed.xlsx");
    let out_path = temp_path("oxlsx_print_area_roundtrip.xlsx");

    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    inject_workbook_defined_names(
        &seeded_path,
        "<definedNames><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">'Sheet'!$A$1:$C$10</definedName><definedName name=\"KeepMe\">Sheet!$D$1</definedName></definedNames>",
    );

    let wb = Workbook::open(&seeded_path).unwrap();
    assert_eq!(
        wb.active().unwrap().print_area(),
        Some("'Sheet'!$A$1:$C$10".to_string())
    );

    wb.save(&out_path).unwrap();
    let wb2 = Workbook::open(&out_path).unwrap();
    assert_eq!(
        wb2.active().unwrap().print_area(),
        Some("'Sheet'!$A$1:$C$10".to_string())
    );

    let workbook_xml = read_zip_entry(&out_path, "xl/workbook.xml");
    assert!(workbook_xml.contains("_xlnm.Print_Area"));
    assert!(
        workbook_xml.contains("&apos;Sheet&apos;!$A$1:$C$10")
            || workbook_xml.contains("'Sheet'!$A$1:$C$10")
    );
    assert!(workbook_xml.contains("KeepMe"));

    fs::remove_file(seeded_path).ok();
    fs::remove_file(out_path).ok();
}

#[test]
fn print_area_write_new_workbook() {
    let path = temp_path("oxlsx_print_area_write_new.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_print_area(Some("$B$2:$D$20"));
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.active().unwrap().print_area(),
        Some("'Sheet'!$B$2:$D$20".to_string())
    );

    let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
    assert!(workbook_xml.contains("_xlnm.Print_Area"));
    // XML escapes ' to &apos;
    assert!(
        workbook_xml.contains("&apos;Sheet&apos;!$B$2:$D$20")
            || workbook_xml.contains("'Sheet'!$B$2:$D$20")
    );

    fs::remove_file(path).ok();
}

#[test]
fn print_area_clear_removes_defined_name() {
    let seeded_path = temp_path("oxlsx_print_area_clear_seed.xlsx");
    let out_path = temp_path("oxlsx_print_area_clear_out.xlsx");

    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    inject_workbook_defined_names(
        &seeded_path,
        "<definedNames><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">'Sheet'!$A$1:$A$5</definedName><definedName name=\"KeepMe\">Sheet!$D$1</definedName></definedNames>",
    );

    let mut wb = Workbook::open(&seeded_path).unwrap();
    assert_eq!(
        wb.active().unwrap().print_area(),
        Some("'Sheet'!$A$1:$A$5".to_string())
    );
    wb.active_mut().unwrap().set_print_area(None);
    wb.save(&out_path).unwrap();

    let wb2 = Workbook::open(&out_path).unwrap();
    assert_eq!(wb2.active().unwrap().print_area(), None);

    let workbook_xml = read_zip_entry(&out_path, "xl/workbook.xml");
    assert!(!workbook_xml.contains("_xlnm.Print_Area"));
    assert!(workbook_xml.contains("KeepMe"));

    fs::remove_file(seeded_path).ok();
    fs::remove_file(out_path).ok();
}

#[test]
fn print_title_rows_and_cols_round_trip_existing_workbook() {
    let seeded_path = temp_path("oxlsx_print_titles_seed.xlsx");
    let out_path = temp_path("oxlsx_print_titles_roundtrip.xlsx");

    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    inject_workbook_defined_names(
        &seeded_path,
        "<definedNames><definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">'Sheet'!$A:$C,'Sheet'!$1:$3</definedName><definedName name=\"KeepMe\">Sheet!$D$1</definedName></definedNames>",
    );

    let wb = Workbook::open(&seeded_path).unwrap();
    let ws = wb.active().unwrap();
    assert_eq!(ws.print_title_cols(), Some("$A:$C"));
    assert_eq!(ws.print_title_rows(), Some("$1:$3"));

    wb.save(&out_path).unwrap();
    let wb2 = Workbook::open(&out_path).unwrap();
    let ws2 = wb2.active().unwrap();
    assert_eq!(ws2.print_title_cols(), Some("$A:$C"));
    assert_eq!(ws2.print_title_rows(), Some("$1:$3"));

    let workbook_xml = read_zip_entry(&out_path, "xl/workbook.xml");
    assert!(workbook_xml.contains("_xlnm.Print_Titles"));
    assert!(workbook_xml.contains("KeepMe"));

    fs::remove_file(seeded_path).ok();
    fs::remove_file(out_path).ok();
}

#[test]
fn print_title_rows_and_cols_write_new_workbook() {
    let path = temp_path("oxlsx_print_titles_write_new.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.active_mut().unwrap();
    ws.set_print_title_rows(Some("$1:$2"));
    ws.set_print_title_cols(Some("$A:$B"));
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    let ws2 = wb2.active().unwrap();
    assert_eq!(ws2.print_title_rows(), Some("$1:$2"));
    assert_eq!(ws2.print_title_cols(), Some("$A:$B"));

    let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
    assert!(workbook_xml.contains("_xlnm.Print_Titles"));
    assert!(
        workbook_xml.contains("&apos;Sheet&apos;!$1:$2") || workbook_xml.contains("'Sheet'!$1:$2")
    );
    assert!(
        workbook_xml.contains("&apos;Sheet&apos;!$A:$B") || workbook_xml.contains("'Sheet'!$A:$B")
    );

    fs::remove_file(path).ok();
}

#[test]
fn print_title_rows_and_cols_clear_removes_defined_name() {
    let seeded_path = temp_path("oxlsx_print_titles_clear_seed.xlsx");
    let out_path = temp_path("oxlsx_print_titles_clear_out.xlsx");

    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    inject_workbook_defined_names(
        &seeded_path,
        "<definedNames><definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">'Sheet'!$A:$C,'Sheet'!$1:$3</definedName><definedName name=\"KeepMe\">Sheet!$D$1</definedName></definedNames>",
    );

    let mut wb = Workbook::open(&seeded_path).unwrap();
    let ws = wb.active_mut().unwrap();
    ws.set_print_title_rows(None);
    assert_eq!(ws.print_title_cols(), Some("$A:$C"));
    wb.save(&out_path).unwrap();

    let wb2 = Workbook::open(&out_path).unwrap();
    let ws2 = wb2.active().unwrap();
    assert_eq!(ws2.print_title_rows(), None);
    assert_eq!(ws2.print_title_cols(), Some("$A:$C"));

    let workbook_xml = read_zip_entry(&out_path, "xl/workbook.xml");
    assert!(workbook_xml.contains("_xlnm.Print_Titles"));
    assert!(!workbook_xml.contains("$1:$3"));
    assert!(workbook_xml.contains("$A:$C"));
    assert!(workbook_xml.contains("KeepMe"));

    fs::remove_file(seeded_path).ok();
    fs::remove_file(out_path).ok();
}

#[test]
fn print_settings_accept_bare_ranges_without_dollar_sign() {
    let path = temp_path("oxlsx_print_bare_ranges.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.active_mut().unwrap();

    // Test bare ranges without $ anchors (openpyxl documented format)
    ws.set_print_title_rows(Some("1:3"));
    ws.set_print_title_cols(Some("A:C"));
    ws.set_print_area(Some("A1:D4"));

    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    let ws2 = wb2.active().unwrap();

    // Getters should return absolute form (qualified for print_area, bare for titles)
    assert_eq!(ws2.print_title_rows(), Some("$1:$3"));
    assert_eq!(ws2.print_title_cols(), Some("$A:$C"));
    assert_eq!(ws2.print_area(), Some("'Sheet'!$A$1:$D$4".to_string()));

    // Verify serialization includes sheet qualifier and absolute anchors
    let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
    assert!(
        workbook_xml.contains("&apos;Sheet&apos;!$1:$3") || workbook_xml.contains("'Sheet'!$1:$3")
    );
    assert!(
        workbook_xml.contains("&apos;Sheet&apos;!$A:$C") || workbook_xml.contains("'Sheet'!$A:$C")
    );
    assert!(
        workbook_xml.contains("&apos;Sheet&apos;!$A$1:$D$4")
            || workbook_xml.contains("'Sheet'!$A$1:$D$4")
    );

    fs::remove_file(path).ok();
}

#[test]
fn print_settings_strip_sheet_prefix_on_input() {
    let path = temp_path("oxlsx_print_strip_prefix.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.active_mut().unwrap();

    // Test with sheet prefix (should be stripped and stored bare)
    ws.set_print_title_rows(Some("'Sheet'!$1:$2"));
    ws.set_print_title_cols(Some("'OtherSheet'!$A:$B"));
    ws.set_print_area(Some("Sheet!$A$1:$D$4"));

    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    let ws2 = wb2.active().unwrap();

    // Getters: print_area returns qualified, titles return bare
    assert_eq!(ws2.print_title_rows(), Some("$1:$2"));
    assert_eq!(ws2.print_title_cols(), Some("$A:$B"));
    assert_eq!(ws2.print_area(), Some("'Sheet'!$A$1:$D$4".to_string()));

    // Serialization adds the correct sheet prefix (current sheet, not input prefix)
    let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
    assert!(
        workbook_xml.contains("&apos;Sheet&apos;!$1:$2") || workbook_xml.contains("'Sheet'!$1:$2")
    );
    assert!(
        workbook_xml.contains("&apos;Sheet&apos;!$A:$B") || workbook_xml.contains("'Sheet'!$A:$B")
    );
    assert!(
        workbook_xml.contains("&apos;Sheet&apos;!$A$1:$D$4")
            || workbook_xml.contains("'Sheet'!$A$1:$D$4")
    );

    fs::remove_file(path).ok();
}

// STEP 1 TDD: Tests that mirror openpyxl behavior EXACTLY (based on probe results)

#[test]
fn openpyxl_mirror_print_area_getter_returns_qualified_absolute() {
    let path = temp_path("test_print_area_qualified.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_print_area(Some("A1:D4"));
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    // MUST return qualified form like openpyxl: 'Sheet'!$A$1:$D$4
    assert_eq!(
        wb2.active().unwrap().print_area(),
        Some("'Sheet'!$A$1:$D$4".to_string())
    );

    fs::remove_file(path).ok();
}

#[test]
fn openpyxl_mirror_print_area_multi_range_qualified() {
    let path = temp_path("test_print_area_multi.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_print_area(Some("A1:B2,D1:E2"));
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    // MUST return qualified form for EACH segment
    assert_eq!(
        wb2.active().unwrap().print_area(),
        Some("'Sheet'!$A$1:$B$2,'Sheet'!$D$1:$E$2".to_string())
    );

    fs::remove_file(path).ok();
}

#[test]
fn openpyxl_mirror_print_area_xml_exact_match() {
    let path = temp_path("test_print_area_xml.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_print_area(Some("A1:D4"));
    wb.save(&path).unwrap();

    let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
    // EXACT assertion — no loose contains
    assert!(workbook_xml.contains(
        r#"<definedName name="_xlnm.Print_Area" localSheetId="0">&apos;Sheet&apos;!$A$1:$D$4</definedName>"#
    ));

    fs::remove_file(path).ok();
}

#[test]
fn openpyxl_mirror_print_area_multi_range_xml_exact() {
    let path = temp_path("test_print_area_multi_xml.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_print_area(Some("A1:B2,D1:E2"));
    wb.save(&path).unwrap();

    let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
    // EXACT assertion for multi-range
    assert!(workbook_xml.contains(
        r#"<definedName name="_xlnm.Print_Area" localSheetId="0">&apos;Sheet&apos;!$A$1:$B$2,&apos;Sheet&apos;!$D$1:$E$2</definedName>"#
    ));

    fs::remove_file(path).ok();
}

#[test]
fn openpyxl_mirror_print_title_rows_xml_exact() {
    let path = temp_path("test_print_title_rows_xml.xlsx");
    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_print_title_rows(Some("1:3"));
    wb.save(&path).unwrap();

    let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
    // EXACT assertion
    assert!(workbook_xml.contains(
        r#"<definedName name="_xlnm.Print_Titles" localSheetId="0">&apos;Sheet&apos;!$1:$3</definedName>"#
    ));

    fs::remove_file(path).ok();
}

#[test]
fn openpyxl_mirror_print_title_both_xml_exact_rows_first() {
    let path = temp_path("test_print_title_both_xml.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.active_mut().unwrap();
    ws.set_print_title_rows(Some("1:2"));
    ws.set_print_title_cols(Some("A:B"));
    wb.save(&path).unwrap();

    let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
    // EXACT assertion — rows FIRST per openpyxl
    assert!(workbook_xml.contains(
        r#"<definedName name="_xlnm.Print_Titles" localSheetId="0">&apos;Sheet&apos;!$1:$2,&apos;Sheet&apos;!$A:$B</definedName>"#
    ));

    fs::remove_file(path).ok();
}

// --- Phase 13 slice 3a: global-scope named ranges (conservative/lossless) ---

#[test]
fn named_ranges_promoted_from_existing_workbook() {
    let seeded_path = temp_path("oxlsx_named_promote_seed.xlsx");
    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    inject_workbook_defined_names(
        &seeded_path,
        "<definedNames><definedName name=\"Sales\">'Sheet'!$A$1:$A$5</definedName></definedNames>",
    );

    let wb = Workbook::open(&seeded_path).unwrap();
    let dn = wb
        .defined_names()
        .get("Sales")
        .expect("Sales should be promoted to the named-range collection");
    assert_eq!(dn.name, "Sales");
    assert_eq!(dn.value, "'Sheet'!$A$1:$A$5");

    fs::remove_file(seeded_path).ok();
}

#[test]
fn named_ranges_add_and_save_round_trip() {
    use oxlsx::DefinedName;
    let path = temp_path("oxlsx_named_add_roundtrip.xlsx");

    let mut wb = Workbook::new();
    wb.defined_names_mut()
        .add(DefinedName::new("Region", "'Sheet'!$B$1:$B$9"))
        .unwrap();
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.defined_names().get("Region").unwrap().value,
        "'Sheet'!$B$1:$B$9"
    );

    fs::remove_file(path).ok();
}

#[test]
fn named_ranges_remove_drops_from_workbook_xml() {
    let seeded_path = temp_path("oxlsx_named_remove_seed.xlsx");
    let out_path = temp_path("oxlsx_named_remove_out.xlsx");

    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    // "Region" is a promotable global; "Keep" is sheet-local => stays raw passthrough.
    inject_workbook_defined_names(
        &seeded_path,
        "<definedNames><definedName name=\"Region\">'Sheet'!$B$1:$B$9</definedName><definedName name=\"Keep\" localSheetId=\"0\">Sheet!$D$1</definedName></definedNames>",
    );

    let mut wb = Workbook::open(&seeded_path).unwrap();
    assert!(wb.defined_names().contains("Region"));
    assert!(wb.defined_names_mut().remove("Region").is_some());
    wb.save(&out_path).unwrap();

    let workbook_xml = read_zip_entry(&out_path, "xl/workbook.xml");
    assert!(!workbook_xml.contains("name=\"Region\""));
    assert!(workbook_xml.contains("Keep"));

    let wb2 = Workbook::open(&out_path).unwrap();
    assert!(!wb2.defined_names().contains("Region"));

    fs::remove_file(seeded_path).ok();
    fs::remove_file(out_path).ok();
}

#[test]
fn named_ranges_exclude_reserved_xlnm() {
    use oxlsx::DefinedName;
    let seeded_path = temp_path("oxlsx_named_reserved_seed.xlsx");

    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    inject_workbook_defined_names(
        &seeded_path,
        "<definedNames><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">'Sheet'!$A$1:$C$10</definedName></definedNames>",
    );

    let mut wb = Workbook::open(&seeded_path).unwrap();
    assert!(!wb.defined_names().contains("_xlnm.Print_Area"));
    assert!(wb.defined_names().names().next().is_none());

    let result = wb
        .defined_names_mut()
        .add(DefinedName::new("_xlnm.Print_Area", "'Sheet'!$A$1:$B$2"));
    assert!(result.is_err());

    fs::remove_file(seeded_path).ok();
}

#[test]
fn named_ranges_conservative_passthrough_preserves_attr_names() {
    let seeded_path = temp_path("oxlsx_named_passthrough_seed.xlsx");
    let out_path = temp_path("oxlsx_named_passthrough_out.xlsx");

    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    // `function="1"` is an unmodeled attribute => must NOT be promoted, must round-trip verbatim.
    inject_workbook_defined_names(
        &seeded_path,
        "<definedNames><definedName name=\"Macro\" function=\"1\">SUM(1,2)</definedName></definedNames>",
    );

    let wb = Workbook::open(&seeded_path).unwrap();
    assert!(wb.defined_names().get("Macro").is_none());
    wb.save(&out_path).unwrap();

    let workbook_xml = read_zip_entry(&out_path, "xl/workbook.xml");
    assert!(workbook_xml.contains("function=\"1\""));
    assert!(workbook_xml.contains("Macro"));

    let wb2 = Workbook::open(&out_path).unwrap();
    assert!(wb2.defined_names().get("Macro").is_none());

    fs::remove_file(seeded_path).ok();
    fs::remove_file(out_path).ok();
}

#[test]
fn named_ranges_coexist_with_print_settings() {
    use oxlsx::DefinedName;
    let path = temp_path("oxlsx_named_coexist.xlsx");

    let mut wb = Workbook::new();
    wb.active_mut().unwrap().set_print_area(Some("$A$1:$C$10"));
    wb.defined_names_mut()
        .add(DefinedName::new("Region", "'Sheet'!$B$1:$B$9"))
        .unwrap();
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.active().unwrap().print_area(),
        Some("'Sheet'!$A$1:$C$10".to_string())
    );
    assert_eq!(
        wb2.defined_names().get("Region").unwrap().value,
        "'Sheet'!$B$1:$B$9"
    );

    let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
    assert!(workbook_xml.contains("_xlnm.Print_Area"));
    assert!(workbook_xml.contains("name=\"Region\""));

    fs::remove_file(path).ok();
}

// --- Phase 13 slice 3b: sheet-local named ranges (ws.defined_names) ---

#[test]
fn sheet_local_names_promoted_from_existing_workbook() {
    let seeded_path = temp_path("oxlsx_sheet_local_promote_seed.xlsx");
    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    inject_workbook_defined_names(
        &seeded_path,
        "<definedNames><definedName name=\"LocalRegion\" localSheetId=\"0\">Sheet!$A$1</definedName></definedNames>",
    );

    let wb = Workbook::open(&seeded_path).unwrap();
    // Promoted onto the owning worksheet...
    let dn = wb
        .sheet(0)
        .unwrap()
        .defined_names()
        .get("LocalRegion")
        .expect("LocalRegion should be promoted onto its worksheet");
    assert_eq!(dn.value, "Sheet!$A$1");
    // ...and NOT visible in the workbook-global collection.
    assert!(!wb.defined_names().contains("LocalRegion"));

    fs::remove_file(seeded_path).ok();
}

#[test]
fn sheet_local_names_add_and_save_round_trip() {
    use oxlsx::DefinedName;
    let path = temp_path("oxlsx_sheet_local_add_roundtrip.xlsx");

    let mut wb = Workbook::new();
    wb.active_mut()
        .unwrap()
        .defined_names_mut()
        .add(DefinedName::new("Region", "Sheet!$B$1"))
        .unwrap();
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.sheet(0)
            .unwrap()
            .defined_names()
            .get("Region")
            .unwrap()
            .value,
        "Sheet!$B$1"
    );
    // Stays scoped to the sheet, never leaks to the global collection.
    assert!(!wb2.defined_names().contains("Region"));

    fs::remove_file(path).ok();
}

#[test]
fn sheet_local_names_localsheetid_matches_sheet_index() {
    use oxlsx::DefinedName;
    let path = temp_path("oxlsx_sheet_local_index.xlsx");

    let mut wb = Workbook::new();
    // Add the name to the SECOND sheet (index 1).
    wb.create_sheet("Data")
        .defined_names_mut()
        .add(DefinedName::new("Region", "Data!$A$1"))
        .unwrap();
    wb.save(&path).unwrap();

    let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
    // Full element assertion: localSheetId computed from the live worksheet position.
    assert!(
        workbook_xml
            .contains(r#"<definedName name="Region" localSheetId="1">Data!$A$1</definedName>"#)
    );

    fs::remove_file(path).ok();
}

#[test]
fn sheet_local_names_exclude_reserved_xlnm() {
    use oxlsx::DefinedName;
    let seeded_path = temp_path("oxlsx_sheet_local_reserved_seed.xlsx");

    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    inject_workbook_defined_names(
        &seeded_path,
        "<definedNames><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">'Sheet'!$A$1:$C$10</definedName></definedNames>",
    );

    let mut wb = Workbook::open(&seeded_path).unwrap();
    // Reserved print name is owned by print settings, never surfaced as a sheet-local name.
    assert!(
        !wb.sheet(0)
            .unwrap()
            .defined_names()
            .contains("_xlnm.Print_Area")
    );
    assert!(
        wb.sheet(0)
            .unwrap()
            .defined_names()
            .names()
            .next()
            .is_none()
    );

    // add() of a reserved name on a worksheet collection is rejected.
    let result = wb
        .sheet_mut(0)
        .unwrap()
        .defined_names_mut()
        .add(DefinedName::new("_xlnm.Print_Area", "'Sheet'!$A$1:$B$2"));
    assert!(result.is_err());

    fs::remove_file(seeded_path).ok();
}

#[test]
fn sheet_local_and_global_same_name_coexist() {
    use oxlsx::DefinedName;
    let path = temp_path("oxlsx_sheet_local_global_same_name.xlsx");

    let mut wb = Workbook::new();
    wb.defined_names_mut()
        .add(DefinedName::new("Data", "Sheet!$A$1"))
        .unwrap();
    wb.active_mut()
        .unwrap()
        .defined_names_mut()
        .add(DefinedName::new("Data", "Sheet!$B$2"))
        .unwrap();
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    // Independent scopes: same name, different values.
    assert_eq!(wb2.defined_names().get("Data").unwrap().value, "Sheet!$A$1");
    assert_eq!(
        wb2.sheet(0)
            .unwrap()
            .defined_names()
            .get("Data")
            .unwrap()
            .value,
        "Sheet!$B$2"
    );

    fs::remove_file(path).ok();
}

#[test]
fn sheet_local_conservative_passthrough_preserves_attr_names() {
    let seeded_path = temp_path("oxlsx_sheet_local_passthrough_seed.xlsx");
    let out_path = temp_path("oxlsx_sheet_local_passthrough_out.xlsx");

    let wb_seed = Workbook::new();
    wb_seed.save(&seeded_path).unwrap();
    // `function="1"` is an unmodeled attribute => NOT promoted, must round-trip verbatim.
    inject_workbook_defined_names(
        &seeded_path,
        "<definedNames><definedName name=\"LocalMacro\" localSheetId=\"0\" function=\"1\">SUM(1,2)</definedName></definedNames>",
    );

    let wb = Workbook::open(&seeded_path).unwrap();
    assert!(
        wb.sheet(0)
            .unwrap()
            .defined_names()
            .get("LocalMacro")
            .is_none()
    );
    assert!(!wb.defined_names().contains("LocalMacro"));
    wb.save(&out_path).unwrap();

    let workbook_xml = read_zip_entry(&out_path, "xl/workbook.xml");
    assert!(workbook_xml.contains("function=\"1\""));
    assert!(workbook_xml.contains("LocalMacro"));

    let wb2 = Workbook::open(&out_path).unwrap();
    assert!(
        wb2.sheet(0)
            .unwrap()
            .defined_names()
            .get("LocalMacro")
            .is_none()
    );

    fs::remove_file(seeded_path).ok();
    fs::remove_file(out_path).ok();
}

#[test]
fn sheet_local_names_remove_drops_from_workbook_xml() {
    use oxlsx::DefinedName;
    let path = temp_path("oxlsx_sheet_local_remove.xlsx");

    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.defined_names_mut()
            .add(DefinedName::new("Region", "Sheet!$B$1"))
            .unwrap();
        ws.defined_names_mut()
            .add(DefinedName::new("Keep", "Sheet!$D$1"))
            .unwrap();
        assert!(ws.defined_names_mut().remove("Region").is_some());
    }
    wb.save(&path).unwrap();

    let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
    assert!(!workbook_xml.contains("name=\"Region\""));
    // Tightened (was a loose substring in 3a): assert the FULL element survives.
    assert!(
        workbook_xml
            .contains(r#"<definedName name="Keep" localSheetId="0">Sheet!$D$1</definedName>"#)
    );

    let wb2 = Workbook::open(&path).unwrap();
    assert!(!wb2.sheet(0).unwrap().defined_names().contains("Region"));
    assert!(wb2.sheet(0).unwrap().defined_names().contains("Keep"));

    fs::remove_file(path).ok();
}

#[test]
fn tables_read_from_existing_workbook() {
    let wb = Workbook::open("samples/tables.xlsx").unwrap();
    let ws = wb.sheet(0).unwrap();

    let table = ws
        .tables()
        .iter()
        .find(|t| t.display_name == "Table1")
        .expect("Table1 should be present in ws.tables");

    assert_eq!(table.display_name, "Table1");
    assert_eq!(table.name, "Table1");
    assert_eq!(table.ref_range, "A1:C3");
    assert_eq!(table.header_row_count, 1);

    let column_names: Vec<&str> = table.columns.iter().map(|c| c.name.as_str()).collect();
    assert_eq!(column_names, vec!["h1", "h2", "h3"]);
    assert_eq!(
        table.columns.iter().map(|c| c.id).collect::<Vec<_>>(),
        vec![1, 2, 3]
    );

    let style = table.style.as_ref().expect("table style should be present");
    assert_eq!(style.name, "TableStyleMedium9");
    assert!(style.show_row_stripes);
    assert!(!style.show_first_column);
}

#[test]
fn tables_add_and_save_round_trip() {
    let path = temp_path("oxlsx_tables_round_trip.xlsx");

    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "h1");
        ws.set_cell("B1", "h2");
        ws.set_cell("C1", "h3");
        ws.add_table(Table::new("Table1", "A1:C3")).unwrap();
    }
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    let table = wb2
        .sheet(0)
        .unwrap()
        .tables()
        .iter()
        .find(|t| t.display_name == "Table1")
        .expect("Table1 round-trips");
    assert_eq!(table.ref_range, "A1:C3");
    assert_eq!(
        table
            .columns
            .iter()
            .map(|c| c.name.as_str())
            .collect::<Vec<_>>(),
        vec!["h1", "h2", "h3"]
    );

    fs::remove_file(path).ok();
}

#[test]
fn tables_add_emits_table_part_and_sheet_rel() {
    let path = temp_path("oxlsx_tables_emit.xlsx");

    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "h1");
        ws.set_cell("B1", "h2");
        ws.add_table(Table::new("Table1", "A1:B2")).unwrap();
    }
    wb.save(&path).unwrap();

    // Table part exists with the right shape.
    let table_xml = read_zip_entry(&path, "xl/tables/table1.xml");
    assert!(table_xml.contains(r#"displayName="Table1""#));
    assert!(table_xml.contains(r#"<autoFilter ref="A1:B2""#));

    // Content-type override present.
    let content_types = read_zip_entry(&path, "[Content_Types].xml");
    assert!(content_types.contains("/xl/tables/table1.xml"));
    assert!(
        content_types
            .contains("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
    );

    // <tableParts> in the sheet XML.
    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(sheet_xml.contains("<tableParts"));
    assert!(sheet_xml.contains("<tablePart"));

    // Sheet rel entry of the table relationship type.
    let rels = read_zip_entry(&path, "xl/worksheets/_rels/sheet1.xml.rels");
    assert!(rels.contains("/relationships/table"));
    assert!(rels.contains("tables/table1.xml"));

    fs::remove_file(path).ok();
}

#[test]
fn tables_columns_derived_from_header_row() {
    let path = temp_path("oxlsx_tables_derive.xlsx");

    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "Name");
        // B1 deliberately left empty to exercise the fallback.
        ws.set_cell("C1", "Age");
        ws.add_table(Table::new("People", "A1:C2")).unwrap();
    }
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    let table = &wb2.sheet(0).unwrap().tables()[0];
    assert_eq!(
        table
            .columns
            .iter()
            .map(|c| c.name.as_str())
            .collect::<Vec<_>>(),
        vec!["Name", "Column2", "Age"]
    );
    assert_eq!(
        table.columns.iter().map(|c| c.id).collect::<Vec<_>>(),
        vec![1, 2, 3]
    );

    fs::remove_file(path).ok();
}

#[test]
fn tables_style_info_round_trip() {
    let path = temp_path("oxlsx_tables_style.xlsx");

    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "h1");
        ws.set_cell("B1", "h2");
        let mut table = Table::new("Styled", "A1:B2");
        table.style = Some(TableStyleInfo {
            name: "TableStyleMedium9".to_string(),
            show_first_column: false,
            show_last_column: true,
            show_row_stripes: true,
            show_column_stripes: false,
        });
        ws.add_table(table).unwrap();
    }
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    let style = wb2.sheet(0).unwrap().tables()[0]
        .style
        .clone()
        .expect("style round-trips");
    assert_eq!(style.name, "TableStyleMedium9");
    assert!(style.show_row_stripes);
    assert!(style.show_last_column);
    assert!(!style.show_first_column);
    assert!(!style.show_column_stripes);

    fs::remove_file(path).ok();
}

#[test]
fn tables_rmw_preserves_existing_tables() {
    // Start from a table-bearing file and inject an UNRELATED sheet rel so we can
    // prove RMW preserves both the table AND unrelated relationships.
    let path = temp_path("oxlsx_tables_rmw.xlsx");
    fs::copy("samples/tables.xlsx", &path).unwrap();
    inject_unrelated_sheet_rel(&path);

    // RMW: open (reads the existing table), modify an unrelated cell, save.
    let mut wb = Workbook::open(&path).unwrap();
    assert_eq!(
        wb.sheet(0).unwrap().tables().len(),
        1,
        "existing table read on open"
    );
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("E5", "touched");
    }
    wb.save(&path).unwrap();

    // Regression for the latent orphan bug: <tableParts> MUST be re-emitted.
    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(
        sheet_xml.contains("<tableParts"),
        "tableParts must be re-emitted on RMW (orphan-fix)"
    );

    // The table part + rel survive, and the unrelated rel is preserved.
    let table_xml = read_zip_entry(&path, "xl/tables/table1.xml");
    assert!(table_xml.contains(r#"displayName="Table1""#));
    let rels = read_zip_entry(&path, "xl/worksheets/_rels/sheet1.xml.rels");
    assert!(rels.contains("/relationships/table"));
    assert!(
        rels.contains("printerSettings"),
        "unrelated sheet rel must be preserved across RMW"
    );

    // Reopen: the table model survives intact.
    let wb2 = Workbook::open(&path).unwrap();
    let table = &wb2.sheet(0).unwrap().tables()[0];
    assert_eq!(table.display_name, "Table1");
    assert_eq!(table.ref_range, "A1:C3");

    fs::remove_file(path).ok();
}

#[test]
fn tables_multiple_tables_unique_ids_and_rids() {
    let path = temp_path("oxlsx_tables_multi.xlsx");

    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "a");
        ws.set_cell("B1", "b");
        ws.set_cell("D1", "d");
        ws.set_cell("E1", "e");
        // A comment on this sheet forces comment + vml rels to coexist with table rels.
        ws.set_cell_comment(
            "G1",
            Comment {
                text: "note".to_string(),
                author: "me".to_string(),
            },
        );
        ws.add_table(Table::new("TblA", "A1:B3")).unwrap();
        ws.add_table(Table::new("TblB", "D1:E3")).unwrap();
    }
    {
        let ws2 = wb.create_sheet("Second");
        ws2.set_cell("A1", "x");
        ws2.add_table(Table::new("TblC", "A1:A2")).unwrap();
    }
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    let mut ids: Vec<u32> = wb2
        .sheet(0)
        .unwrap()
        .tables()
        .iter()
        .map(|t| t.id)
        .collect();
    ids.extend(wb2.sheet(1).unwrap().tables().iter().map(|t| t.id));
    assert_eq!(ids.len(), 3);
    assert!(ids.iter().all(|id| *id != 0), "ids must be assigned");
    let mut unique = ids.clone();
    unique.sort_unstable();
    unique.dedup();
    assert_eq!(unique.len(), ids.len(), "table ids unique workbook-wide");

    // Comment survives alongside the tables.
    assert!(
        wb2.sheet(0)
            .unwrap()
            .cell("G1")
            .and_then(|c| c.comment())
            .is_some(),
        "comment preserved alongside tables"
    );

    // Sheet 1's rels contain comment + vml + two table rels. Unit 2 guarantees
    // the TABLE rIds are distinct from each other AND disjoint from every other
    // (comment/vml) rId — i.e. tables never collide with pre-existing rels.
    let rels = read_zip_entry(&path, "xl/worksheets/_rels/sheet1.xml.rels");
    let mut table_rids = Vec::new();
    let mut other_rids = Vec::new();
    for chunk in rels.split("<Relationship ").skip(1) {
        let Some(id) = chunk
            .split("Id=\"")
            .nth(1)
            .and_then(|c| c.split('"').next())
        else {
            continue;
        };
        if chunk.contains("/relationships/table") {
            table_rids.push(id.to_string());
        } else {
            other_rids.push(id.to_string());
        }
    }
    assert_eq!(table_rids.len(), 2, "two table rels on sheet 1");
    assert_ne!(table_rids[0], table_rids[1], "table rIds distinct");
    for table_rid in &table_rids {
        assert!(
            !other_rids.contains(table_rid),
            "table rId must not collide with comment/vml rels"
        );
    }

    // FIX 4 remainder: ALL relationship IDs in this sheet's .rels must be unique.
    // (Sheet 1 has tables; sheet 2 has no comments, so sheet 2 is a clean sample for
    // uniqueness without entangling pre-existing comment/vml collision #274.)
    let sheet2_rels = read_zip_entry(&path, "xl/worksheets/_rels/sheet2.xml.rels");
    let mut all_rids_sheet2 = Vec::new();
    for chunk in sheet2_rels.split("<Relationship ").skip(1) {
        if let Some(id) = chunk
            .split("Id=\"")
            .nth(1)
            .and_then(|c| c.split('"').next())
        {
            all_rids_sheet2.push(id.to_string());
        }
    }
    let mut unique_rids = all_rids_sheet2.clone();
    unique_rids.sort_unstable();
    unique_rids.dedup();
    assert_eq!(
        unique_rids.len(),
        all_rids_sheet2.len(),
        "all relationship IDs in sheet2.xml.rels must be unique (FIX 4)"
    );

    fs::remove_file(path).ok();
}

#[test]
fn tables_save_rejects_duplicate_display_name() {
    let path = temp_path("oxlsx_tables_dup.xlsx");

    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "h1");
        ws.add_table(Table::new("Dup", "A1:A2")).unwrap();
    }
    {
        // Same displayName on a DIFFERENT sheet: add_table's same-sheet check does
        // not catch it, but the workbook-wide save guard must.
        let ws2 = wb.create_sheet("Second");
        ws2.set_cell("A1", "h1");
        ws2.add_table(Table::new("Dup", "A1:A2")).unwrap();
    }

    let result = wb.save(&path);
    assert!(
        matches!(result, Err(oxlsx::OxlsxError::DuplicateTableName(_))),
        "workbook-wide duplicate displayName must be rejected at save"
    );
    assert!(
        !path.exists(),
        "no file (not even a 0-byte one) must be created when the duplicate guard trips"
    );

    fs::remove_file(path).ok();
}

// --- Judgment Day Phase 13 fixes ---

/// FIX 2: Invalid ref → add_table Err (Rust) / Python ValueError
#[test]
fn tables_invalid_ref_rejected_at_add_time() {
    let mut wb = Workbook::new();
    let ws = wb.active_mut().unwrap();
    ws.set_cell("A1", "h1");

    // FIX 2 WEAK TEST: Test ACTUAL width validation (reversed columns).
    // B1:A2 means start_col > end_col → zero/negative width.
    let result = ws.add_table(Table::new("BadTable", "B1:A2"));
    assert!(
        result.is_err(),
        "reversed/zero-width column ref (B1:A2) must be rejected at add_table"
    );

    // Also test unparseable ref (row 0):
    let result2 = ws.add_table(Table::new("BadTable2", "A1:A0"));
    assert!(
        result2.is_err(),
        "unparseable ref A1:A0 (row 0) must be rejected"
    );
}

/// FIX 5: Duplicate column headers → deduped column names
#[test]
fn tables_duplicate_header_columns_deduped() {
    let path = temp_path("oxlsx_tables_dedup_headers.xlsx");

    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "Total");
        ws.set_cell("B1", "Total");
        ws.set_cell("C1", "Total");
        // D1 left empty to also test "Column{id}" fallback coexistence.
        ws.add_table(Table::new("DupHeaders", "A1:D2")).unwrap();
    }
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    let table = &wb2.sheet(0).unwrap().tables()[0];
    let column_names: Vec<&str> = table.columns.iter().map(|c| c.name.as_str()).collect();
    // Mirror openpyxl: first instance keeps its name, subsequent duplicates get "Total1", "Total2", etc.
    // Empty cell gets "Column4" (because id=4).
    assert_eq!(
        column_names,
        vec!["Total", "Total1", "Total2", "Column4"],
        "column names must be unique case-insensitively"
    );

    fs::remove_file(path).ok();
}

/// FIX 5 CRITICAL 2: Re-collision case where suffix itself is a duplicate
#[test]
fn tables_dedup_re_collision() {
    let path = temp_path("oxlsx_tables_dedup_re_collision.xlsx");

    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "Total");
        ws.set_cell("B1", "Total1"); // Already exists
        ws.set_cell("C1", "Total"); // Would collide with "Total1" if not re-checked
        ws.add_table(Table::new("ReCollision", "A1:C2")).unwrap();
    }
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    let table = &wb2.sheet(0).unwrap().tables()[0];
    let column_names: Vec<&str> = table.columns.iter().map(|c| c.name.as_str()).collect();
    // Expected: ["Total", "Total1", "Total2"]
    // NOT: ["Total", "Total1", "Total1"] (duplicate!)
    assert_eq!(
        column_names,
        vec!["Total", "Total1", "Total2"],
        "dedup must re-check suffixed names to prevent re-collision"
    );

    fs::remove_file(path).ok();
}

/// FIX 6: Case-insensitive displayName uniqueness (Rust save guard)
#[test]
fn tables_case_insensitive_duplicate_rejected_rust() {
    let path = temp_path("oxlsx_tables_case_insensitive_rust.xlsx");

    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "h1");
        ws.add_table(Table::new("MyTable", "A1:A2")).unwrap();
    }
    {
        let ws2 = wb.create_sheet("Second");
        ws2.set_cell("A1", "h1");
        ws2.add_table(Table::new("MYTABLE", "A1:A2")).unwrap(); // Case differs
    }

    let result = wb.save(&path);
    assert!(
        matches!(result, Err(OxlsxError::DuplicateTableName(_))),
        "case-insensitive duplicate displayName must be rejected"
    );

    fs::remove_file(path).ok();
}

/// FIX 1: Stale content-type Overrides on RMW
#[test]
fn tables_rmw_strips_stale_content_type_overrides() {
    // This test requires a fixture whose table part is NOT table1.xml.
    // We'll create it in the venv with openpyxl, ensuring a higher index.
    // For simplicity, we'll craft one manually or use openpyxl via bash if needed.
    // For now, create a seed with table part table5.xml, then RMW to table1.xml.

    let seed_path = temp_path("oxlsx_tables_rmw_stale_seed.xlsx");
    let out_path = temp_path("oxlsx_tables_rmw_stale_out.xlsx");

    // Step 1: Create a workbook with a table, save it.
    let mut wb_seed = Workbook::new();
    {
        let ws = wb_seed.active_mut().unwrap();
        ws.set_cell("A1", "h1");
        ws.add_table(Table::new("Original", "A1:A2")).unwrap();
    }
    wb_seed.save(&seed_path).unwrap();

    // Step 2: Manually inject a FAKE content-type override for table5.xml to simulate
    // a file that originally had a higher table index. In real-world RMW this would
    // happen if the source file had multiple tables and some were removed.
    // For strict TDD, we'll use the inject pattern from other tests.
    inject_fake_table_override(&seed_path, 5);

    // Step 3: RMW — open the seed (which has table1.xml + a stale table5.xml override), save.
    let mut wb = Workbook::open(&seed_path).unwrap();
    wb.active_mut().unwrap().set_cell("B1", "touched");
    wb.save(&out_path).unwrap();

    // Step 4: Assert: [Content_Types].xml MUST have ONLY table1.xml override, NO table5.xml.
    let content_types = read_zip_entry(&out_path, "[Content_Types].xml");
    assert!(
        content_types.contains("/xl/tables/table1.xml"),
        "writer-owned table1.xml must be present"
    );
    assert!(
        !content_types.contains("/xl/tables/table5.xml"),
        "stale table5.xml override must be stripped"
    );

    // FIX 1: Assert EXACTLY one occurrence of table1.xml (no duplicates).
    let table1_count = content_types.matches("/xl/tables/table1.xml").count();
    assert_eq!(
        table1_count, 1,
        "table1.xml override must appear EXACTLY once (was: {table1_count})"
    );

    fs::remove_file(seed_path).ok();
    fs::remove_file(out_path).ok();
}

/// FIX 1: Test RMW of REAL openpyxl-authored samples/tables.xlsx
#[test]
fn tables_rmw_real_sample_no_duplicate_overrides() {
    let out_path = temp_path("oxlsx_tables_real_sample_rmw.xlsx");

    // Open the REAL openpyxl-authored sample (which has space before "/>").
    let mut wb = Workbook::open("samples/tables.xlsx").unwrap();
    wb.active_mut().unwrap().set_cell("B1", "touched");
    wb.save(&out_path).unwrap();

    // Assert: [Content_Types].xml has EXACTLY one table1.xml override (no duplicate).
    let content_types = read_zip_entry(&out_path, "[Content_Types].xml");
    let table1_count = content_types.matches("/xl/tables/table1.xml").count();
    assert_eq!(
        table1_count, 1,
        "RMW of samples/tables.xlsx must produce exactly one table1.xml override (was: {table1_count})"
    );

    // CRITICAL: RMW that creates sharedStrings MUST emit the content-type override,
    // otherwise openpyxl cannot reopen the file (verified by test_tables_rmw_openpyxl_reopen in test_python.py).
    assert!(
        content_types.contains("/xl/sharedStrings.xml"),
        "RMW that creates shared strings must declare the sharedStrings content-type override"
    );

    fs::remove_file(out_path).ok();
}

// Helper for FIX 1 test: inject a fake content-type override for tableN.xml
fn inject_fake_table_override(path: &std::path::Path, table_number: usize) {
    let source_file = fs::File::open(path).unwrap();
    let mut source_zip = ZipArchive::new(source_file).unwrap();
    let temp_path = path.with_extension("tmp.xlsx");
    let temp_file = fs::File::create(&temp_path).unwrap();
    let mut writer = ZipWriter::new(temp_file);
    let options = SimpleFileOptions::default();

    for idx in 0..source_zip.len() {
        let mut entry = source_zip.by_index(idx).unwrap();
        let name = entry.name().to_string();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();

        if name == "[Content_Types].xml" {
            let xml = String::from_utf8(bytes).unwrap();
            // Use REAL openpyxl format: space before "/>" (not self-closing without space).
            let fake_override = format!(
                r#"<Override PartName="/xl/tables/table{table_number}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml" />"#
            );
            let injected = xml.replace("</Types>", &format!("{fake_override}</Types>"));
            bytes = injected.into_bytes();
        }

        writer.start_file(name, options).unwrap();
        writer.write_all(&bytes).unwrap();
    }

    writer.finish().unwrap();
    fs::rename(temp_path, path).unwrap();
}

// --- Lossless RMW of unmodeled table features (Unit 1) ---------------------

fn zip_has_entry(path: &std::path::Path, entry: &str) -> bool {
    let file = fs::File::open(path).unwrap();
    let mut zip = ZipArchive::new(file).unwrap();
    zip.by_name(entry).is_ok()
}

#[test]
fn table_unmodified_rmw_preserves_raw_table_xml() {
    let path = temp_path("oxlsx_tables_totals_rmw.xlsx");
    fs::copy("samples/tables_totals.xlsx", &path).unwrap();

    // Verbatim bytes of the SalesTable part BEFORE the round-trip.
    let original_table1 = read_zip_entry(&path, "xl/tables/table1.xml");

    let mut wb = Workbook::open(&path).unwrap();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("H10", "unrelated"); // a cell outside every table
    }
    wb.save(&path).unwrap();

    let saved_table1 = read_zip_entry(&path, "xl/tables/table1.xml");
    assert_eq!(
        saved_table1, original_table1,
        "clean table must be re-emitted byte-for-byte"
    );
    assert!(saved_table1.contains(r#"totalsRowCount="1""#));
    assert!(saved_table1.contains(r#"totalsRowFunction="sum""#));
    assert!(saved_table1.contains("calculatedColumnFormula"));
    assert!(saved_table1.contains("<sortState"));

    // Wiring intact.
    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(sheet_xml.contains("<tableParts"));
    let content_types = read_zip_entry(&path, "[Content_Types].xml");
    assert!(content_types.contains("/xl/tables/table1.xml"));

    fs::remove_file(path).ok();
}

#[test]
fn table_mutated_rmw_serializes_from_model() {
    let path = temp_path("oxlsx_tables_totals_mut.xlsx");
    fs::copy("samples/tables_totals.xlsx", &path).unwrap();

    let mut wb = Workbook::open(&path).unwrap();
    {
        let ws = wb.active_mut().unwrap();
        let table = ws.table_mut("SalesTable").expect("SalesTable present");
        table.style = Some(TableStyleInfo {
            name: "TableStyleMedium2".to_string(),
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: false,
            show_column_stripes: false,
        });
    }
    wb.save(&path).unwrap();

    let saved_table1 = read_zip_entry(&path, "xl/tables/table1.xml");
    // Documented feature loss on the model path.
    assert!(!saved_table1.contains("totalsRowCount"));
    assert!(!saved_table1.contains("sortState"));
    assert!(!saved_table1.contains("calculatedColumnFormula"));
    // Modeled fields round-trip.
    assert!(saved_table1.contains(r#"displayName="SalesTable""#));
    assert!(saved_table1.contains(r#"ref="A1:C5""#));
    assert!(saved_table1.contains(r#"name="Region""#));
    assert!(saved_table1.contains("TableStyleMedium2"));

    // Inventory (untouched) stays verbatim as table2.
    let saved_table2 = read_zip_entry(&path, "xl/tables/table2.xml");
    assert!(saved_table2.contains(r#"displayName="Inventory""#));

    let wb2 = Workbook::open(&path).unwrap();
    let table = &wb2.active().unwrap().tables()[0];
    assert_eq!(table.display_name, "SalesTable");
    assert_eq!(table.ref_range, "A1:C5");

    fs::remove_file(path).ok();
}

#[test]
fn table_added_then_saved_uses_model_path() {
    let path = temp_path("oxlsx_tables_added_model.xlsx");
    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "h1");
        ws.set_cell("B1", "h2");
        ws.set_cell("A2", "v1");
        ws.set_cell("B2", "v2");
        ws.add_table(Table::new("NewTbl", "A1:B2")).unwrap();
    }
    wb.save(&path).unwrap();

    let table1 = read_zip_entry(&path, "xl/tables/table1.xml");
    assert!(table1.contains(r#"displayName="NewTbl""#));
    assert!(table1.contains(r#"<autoFilter ref="A1:B2""#));

    let wb2 = Workbook::open(&path).unwrap();
    let table = &wb2.active().unwrap().tables()[0];
    assert_eq!(table.display_name, "NewTbl");
    assert_eq!(table.id, 1, "model table gets an allocated id");

    fs::remove_file(path).ok();
}

#[test]
fn table_removed_drops_part_rel_and_tablepart() {
    let path = temp_path("oxlsx_tables_removed.xlsx");
    fs::copy("samples/tables_totals.xlsx", &path).unwrap();

    let mut wb = Workbook::open(&path).unwrap();
    {
        let ws = wb.active_mut().unwrap();
        assert_eq!(ws.tables().len(), 2, "fixture has two retained tables");
        let removed = ws.remove_table("Inventory").expect("Inventory present");
        assert_eq!(removed.display_name, "Inventory");
        assert_eq!(ws.tables().len(), 1);
    }
    wb.save(&path).unwrap();

    // The removed table's part is gone; only one table part remains.
    assert!(zip_has_entry(&path, "xl/tables/table1.xml"));
    assert!(!zip_has_entry(&path, "xl/tables/table2.xml"));

    // The surviving table is SalesTable, still verbatim (totals preserved).
    let table1 = read_zip_entry(&path, "xl/tables/table1.xml");
    assert!(table1.contains(r#"displayName="SalesTable""#));
    assert!(table1.contains("totalsRowCount"));

    // <tableParts> reduced to one; content-type override for table2 gone.
    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(sheet_xml.contains(r#"<tableParts count="1""#));
    let content_types = read_zip_entry(&path, "[Content_Types].xml");
    assert!(content_types.contains("/xl/tables/table1.xml"));
    assert!(!content_types.contains("/xl/tables/table2.xml"));

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(wb2.active().unwrap().tables().len(), 1);

    fs::remove_file(path).ok();
}

#[test]
fn mixed_verbatim_and_model_tables_unique_ids() {
    let path = temp_path("oxlsx_tables_mixed.xlsx");
    fs::copy("samples/tables_totals.xlsx", &path).unwrap();

    let mut wb = Workbook::open(&path).unwrap();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A10", "n1");
        ws.set_cell("B10", "n2");
        ws.set_cell("A11", "x");
        ws.set_cell("B11", "y");
        ws.add_table(Table::new("NewModelTbl", "A10:B11")).unwrap();
    }
    wb.save(&path).unwrap();

    // Internal <table id> from every emitted part.
    let mut ids = Vec::new();
    for n in 1..=3 {
        let xml = read_zip_entry(&path, &format!("xl/tables/table{n}.xml"));
        ids.push(attr_in(&xml, "id").unwrap());
    }
    let mut uniq = ids.clone();
    uniq.sort();
    uniq.dedup();
    assert_eq!(
        uniq.len(),
        3,
        "table ids must be unique workbook-wide: {ids:?}"
    );
    // Verbatim tables kept their original ids; the model table avoided them.
    assert!(ids.contains(&"1".to_string()));
    assert!(ids.contains(&"2".to_string()));

    let wb2 = Workbook::open(&path).unwrap();
    let names: Vec<&str> = wb2
        .active()
        .unwrap()
        .tables()
        .iter()
        .map(|t| t.display_name.as_str())
        .collect();
    assert_eq!(names.len(), 3);
    assert!(names.contains(&"SalesTable"));
    assert!(names.contains(&"Inventory"));
    assert!(names.contains(&"NewModelTbl"));

    fs::remove_file(path).ok();
}

#[test]
fn tables_readonly_access_keeps_table_verbatim() {
    let path = temp_path("oxlsx_tables_readonly.xlsx");
    fs::copy("samples/tables_totals.xlsx", &path).unwrap();
    let original_table1 = read_zip_entry(&path, "xl/tables/table1.xml");

    let mut wb = Workbook::open(&path).unwrap();
    {
        // Read-only access (by-name lookup through tables()) must NOT mark dirty.
        let ws = wb.active().unwrap();
        let found = ws.tables().iter().find(|t| t.display_name == "SalesTable");
        assert!(found.is_some());
    }
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("H1", "unrelated");
    }
    wb.save(&path).unwrap();

    let saved_table1 = read_zip_entry(&path, "xl/tables/table1.xml");
    assert_eq!(
        saved_table1, original_table1,
        "read-only access must leave the table clean and verbatim"
    );

    fs::remove_file(path).ok();
}

// --- Hyperlink Slice 1: byte-identity guard (Scenario 5) -------------------

/// Adding the `target_mode` field to `SheetRelEntry` MUST NOT change the
/// serialization of existing comment/table/printerSettings rels. The
/// discriminator: RMW `samples/tables_totals.xlsx` (whose sheet rels are two
/// `table` relationships with no `TargetMode`) and assert that the re-emitted
/// sheet rels carry NO `TargetMode` attribute, still reference both table
/// parts, and that the table part itself stays byte-for-byte identical (the
/// existing no-repair / lossless-RMW byte-equality discriminator).
#[test]
fn existing_rels_remain_byte_identical_without_hyperlinks() {
    let path = temp_path("oxlsx_hyperlink_byte_identity.xlsx");
    fs::copy("samples/tables_totals.xlsx", &path).unwrap();

    let original_table1 = read_zip_entry(&path, "xl/tables/table1.xml");

    let mut wb = Workbook::open(&path).unwrap();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("H10", "unrelated"); // outside every table / hyperlink
    }
    wb.save(&path).unwrap();

    // The sheet rels must NOT gain a TargetMode attribute on the legacy table
    // rels — the whole point of defaulting `target_mode` to `None`.
    let sheet_rels = read_zip_entry(&path, "xl/worksheets/_rels/sheet1.xml.rels");
    assert!(
        !sheet_rels.contains("TargetMode"),
        "no TargetMode may appear on non-hyperlink rels: {sheet_rels}"
    );

    // Both table relationships still resolve to their parts.
    let table_targets: Vec<String> = sheet_rels
        .split("<Relationship")
        .skip(1)
        .filter(|frag| {
            attr_in(frag, "Type")
                .map(|t| t.ends_with("/table"))
                .unwrap_or(false)
        })
        .filter_map(|frag| attr_in(frag, "Target"))
        .collect();
    assert_eq!(
        table_targets.len(),
        2,
        "both table rels preserved: {sheet_rels}"
    );
    assert!(table_targets.iter().any(|t| t.contains("table1.xml")));
    assert!(table_targets.iter().any(|t| t.contains("table2.xml")));

    // The table part is still re-emitted byte-for-byte (lossless-RMW guard).
    let saved_table1 = read_zip_entry(&path, "xl/tables/table1.xml");
    assert_eq!(
        saved_table1, original_table1,
        "clean table part must stay byte-identical after target_mode was added"
    );

    fs::remove_file(path).ok();
}

// --- Hyperlink Slice 2: reader/writer round-trip + RMW (Scenarios 1-4, 6) ---

/// Scenario 1: open the openpyxl-authored fixture and read the external
/// hyperlink attached to A1, resolving `r:id` → external rel target.
#[test]
fn hyperlink_read_from_existing_workbook() {
    let wb = Workbook::open("samples/hyperlinks.xlsx").unwrap();
    let ws = wb.active().unwrap();
    let hyperlink = ws.cell("A1").unwrap().hyperlink().unwrap();
    assert_eq!(hyperlink.target, "https://example.com");
}

/// Scenario 2: author an external hyperlink on a new workbook, save, reopen.
/// The target round-trips, `<hyperlinks>` is emitted AFTER `<mergeCells>`, and
/// the sheet rel carries `TargetMode="External"`.
#[test]
fn hyperlink_write_new_workbook() {
    let path = temp_path("oxlsx_hyperlink_write_new.xlsx");
    {
        let mut wb = Workbook::new();
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "link");
        ws.merge_cells("C1:D1");
        ws.set_cell_hyperlink("A1", "http://rust-lang.org");
        wb.save(&path).unwrap();
    }

    // Target round-trips through a reopen.
    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(
        wb2.active()
            .unwrap()
            .cell("A1")
            .unwrap()
            .hyperlink()
            .unwrap()
            .target,
        "http://rust-lang.org"
    );

    // `<hyperlinks>` must come strictly AFTER `<mergeCells>` (CT_Worksheet order;
    // misplacement triggers Excel repair).
    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    let merge_pos = sheet_xml.find("<mergeCells").expect("mergeCells present");
    let hyperlinks_pos = sheet_xml.find("<hyperlinks").expect("hyperlinks present");
    assert!(
        merge_pos < hyperlinks_pos,
        "<hyperlinks> must be emitted after <mergeCells>: {sheet_xml}"
    );

    // External rel present with TargetMode="External".
    let rels = read_zip_entry(&path, "xl/worksheets/_rels/sheet1.xml.rels");
    assert!(
        rels.contains("TargetMode=\"External\""),
        "external hyperlink rel must carry TargetMode=External: {rels}"
    );
    assert!(rels.contains("http://rust-lang.org"), "rel target: {rels}");

    fs::remove_file(path).ok();
}

/// Scenario 3: RMW preserves an existing external hyperlink and its
/// `tooltip`/`display`, plus the unrelated external rel, when an unrelated cell
/// is edited.
#[test]
fn hyperlink_rmw_preserves_existing() {
    let path = temp_path("oxlsx_hyperlink_rmw_preserve.xlsx");
    fs::copy("samples/hyperlinks.xlsx", &path).unwrap();

    {
        let mut wb = Workbook::open(&path).unwrap();
        wb.active_mut().unwrap().set_cell("B2", "unrelated edit");
        wb.save(&path).unwrap();
    }

    let wb2 = Workbook::open(&path).unwrap();
    let ws = wb2.active().unwrap();
    let hyperlink = ws.cell("A1").unwrap().hyperlink().unwrap();
    assert_eq!(hyperlink.target, "https://example.com");
    assert_eq!(hyperlink.tooltip.as_deref(), Some("Example"));
    assert_eq!(hyperlink.display.as_deref(), Some("Example Site"));

    let rels = read_zip_entry(&path, "xl/worksheets/_rels/sheet1.xml.rels");
    assert!(
        rels.contains("TargetMode=\"External\"") && rels.contains("https://example.com"),
        "unrelated external hyperlink rel preserved on RMW: {rels}"
    );

    fs::remove_file(path).ok();
}

/// Scenario 4: a sheet holding a hyperlink AND a comment AND a table must
/// round-trip with unique, non-colliding rIds.
#[test]
fn hyperlink_coexists_with_comments_and_tables() {
    let path = temp_path("oxlsx_hyperlink_coexist.xlsx");
    {
        let mut wb = Workbook::new();
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "header");
        ws.set_cell("B1", "h2");
        ws.set_cell("A2", "r1");
        ws.set_cell("B2", "r2");
        ws.set_cell("A3", "r3");
        ws.set_cell("B3", "r4");
        ws.set_cell_comment(
            "A1",
            Comment {
                text: "note".to_string(),
                author: "Al".to_string(),
            },
        );
        ws.add_table(Table::new("CoTable", "A1:B3")).unwrap();
        ws.set_cell("A4", "link");
        ws.set_cell_hyperlink("A4", "https://example.org");
        wb.save(&path).unwrap();
    }

    let wb2 = Workbook::open(&path).unwrap();
    let ws = wb2.active().unwrap();
    assert_eq!(
        ws.cell("A4").unwrap().hyperlink().unwrap().target,
        "https://example.org"
    );
    assert!(
        ws.cell("A1").unwrap().comment().is_some(),
        "comment round-trips"
    );
    assert_eq!(ws.tables().len(), 1, "table round-trips");

    // Every sheet rId must be unique (hyperlink + comment + vml + table).
    let rels = read_zip_entry(&path, "xl/worksheets/_rels/sheet1.xml.rels");
    let ids = relationship_ids(&rels);
    let mut unique = ids.clone();
    unique.sort();
    unique.dedup();
    assert_eq!(
        unique.len(),
        ids.len(),
        "duplicate relationship Id found: {rels}"
    );
    assert!(
        rels.contains("TargetMode=\"External\""),
        "hyperlink rel present: {rels}"
    );

    fs::remove_file(path).ok();
}

/// Scenario 6: an internal `location`-only hyperlink (no rId) survives RMW
/// verbatim — no rId is fabricated and no external target is invented.
#[test]
fn internal_location_hyperlink_preserved_on_rmw() {
    let path = temp_path("oxlsx_hyperlink_internal_rmw.xlsx");
    fs::copy("samples/hyperlinks.xlsx", &path).unwrap();

    {
        let mut wb = Workbook::open(&path).unwrap();
        wb.active_mut().unwrap().set_cell("E5", "unrelated");
        wb.save(&path).unwrap();
    }

    let wb2 = Workbook::open(&path).unwrap();
    let ws = wb2.active().unwrap();
    let hyperlink = ws.cell("B1").unwrap().hyperlink().unwrap();
    assert_eq!(hyperlink.location.as_deref(), Some("Sheet!A1"));
    assert!(
        hyperlink.target.is_empty(),
        "internal hyperlink must not fabricate an external target/rId"
    );

    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(
        sheet_xml.contains("location=\"Sheet!A1\""),
        "internal location hyperlink preserved: {sheet_xml}"
    );

    fs::remove_file(path).ok();
}

/// Sheet-level `<autoFilter>` survives a write/reopen round trip, and the
/// element is emitted after `</sheetData>` and before `<mergeCells>` when both
/// are present (CT_Worksheet child ordering).
#[test]
fn auto_filter_round_trip_write() {
    let path = temp_path("oxlsx_auto_filter_write.xlsx");

    let mut wb = Workbook::new();
    {
        let ws = wb.active_mut().unwrap();
        ws.set_cell("A1", "h");
        ws.merge_cells("A5:C5");
        ws.set_auto_filter(Some("A1:C5"));
    }
    wb.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(wb2.active().unwrap().auto_filter(), Some("A1:C5"));

    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(
        sheet_xml.contains("<autoFilter ref=\"A1:C5\""),
        "autoFilter element present: {sheet_xml}"
    );

    let filter_pos = sheet_xml.find("<autoFilter").expect("autoFilter present");
    let merge_pos = sheet_xml.find("<mergeCells").expect("mergeCells present");
    let data_end = sheet_xml.find("</sheetData>").expect("sheetData present");
    assert!(
        data_end < filter_pos && filter_pos < merge_pos,
        "autoFilter must sit after </sheetData> and before <mergeCells>: {sheet_xml}"
    );

    fs::remove_file(path).ok();
}

/// A sheet-level autofilter authored by openpyxl is read into `auto_filter`.
#[test]
fn auto_filter_read_from_existing() {
    let wb = Workbook::open("samples/autofilter.xlsx").unwrap();
    assert_eq!(wb.active().unwrap().auto_filter(), Some("A1:C3"));
}

/// Editing an unrelated cell preserves the sheet-level autofilter across RMW.
#[test]
fn auto_filter_rmw_preserves() {
    let path = temp_path("oxlsx_auto_filter_rmw.xlsx");
    fs::copy("samples/autofilter.xlsx", &path).unwrap();

    {
        let mut wb = Workbook::open(&path).unwrap();
        assert_eq!(wb.active().unwrap().auto_filter(), Some("A1:C3"));
        wb.active_mut().unwrap().set_cell("E10", "unrelated");
        wb.save(&path).unwrap();
    }

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(wb2.active().unwrap().auto_filter(), Some("A1:C3"));

    let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
    assert!(
        sheet_xml.contains("<autoFilter ref=\"A1:C3\""),
        "autoFilter preserved on RMW: {sheet_xml}"
    );

    fs::remove_file(path).ok();
}
