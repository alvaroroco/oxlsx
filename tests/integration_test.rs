use oxlsx::{Color, StyleId, Value, Workbook};
use chrono::NaiveDate;

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
