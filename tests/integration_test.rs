use chrono::NaiveDate;
use oxlsx::{Color, Fill, Font, StyleId, Value, Workbook};
use std::env;

fn temp_path(name: &str) -> std::path::PathBuf {
    env::temp_dir().join(name)
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
