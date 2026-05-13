use oxlsx::{Color, Workbook};

#[test]
fn phase1_styles_and_cells_from_sample_workbook() {
    // A1: style_id=1, bold font, solid fill, indexed color 2
    let wb = Workbook::open("samples/test.xlsx").unwrap();
    let ws = wb.sheet(0).unwrap();
    let a1 = ws.cell("A1").unwrap();
    assert_eq!(a1.style_id(), 1);
    assert!(a1.font().unwrap().bold);
    assert_eq!(a1.fill().unwrap().pattern_type, "solid");
    assert_eq!(a1.fill().unwrap().fg_color, Some(Color::Indexed(2)));

    // B1: style_id=2, blue font (indexed 4), yellow fill (indexed 5)
    let b1 = ws.cell("B1").unwrap();
    assert_eq!(b1.style_id(), 2);
    assert_eq!(b1.font().unwrap().color, Some(Color::Indexed(4)));
    assert_eq!(b1.fill().unwrap().fg_color, Some(Color::Indexed(5)));

    // C1: style_id=0 (default), number value 42
    let c1 = ws.cell("C1").unwrap();
    assert_eq!(c1.style_id(), 0);

    // D1: does not exist
    assert!(ws.cell("D1").is_none());
    assert!(ws.cell("ZZ999").is_none());
}
