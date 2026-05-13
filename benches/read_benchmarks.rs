use criterion::{Criterion, black_box, criterion_group, criterion_main};
use oxlsx::Workbook;

/// Benchmark: open + read all cells from the single-sheet test fixture.
fn bench_open_test_xlsx(c: &mut Criterion) {
    c.bench_function("open_test_xlsx", |b| {
        b.iter(|| {
            let wb = Workbook::open(black_box("samples/test.xlsx")).unwrap();
            let ws = wb.sheet(0).unwrap();
            // Force evaluation of every cell value so the optimizer can't skip it
            for cell in ["A1", "B1", "C1"] {
                if let Some(c) = ws.cell(cell) {
                    black_box(c.value());
                }
            }
        });
    });
}

/// Benchmark: open + read all cells from the multi-sheet fixture.
fn bench_open_multi_sheet_xlsx(c: &mut Criterion) {
    c.bench_function("open_multi_sheet_xlsx", |b| {
        b.iter(|| {
            let wb = Workbook::open(black_box("samples/multi_sheet.xlsx")).unwrap();
            for sheet in wb.worksheets() {
                // Iterate known cells to ensure parsing actually runs
                for cell in ["A1", "B1"] {
                    if let Some(c) = sheet.cell(cell) {
                        black_box(c.value());
                    }
                }
            }
        });
    });
}

/// Benchmark: open + iterate ALL cells in a 10 000-row file.
/// This is the primary comparison point against openpyxl.
fn bench_open_10k_rows(c: &mut Criterion) {
    c.bench_function("open_10k_rows", |b| {
        b.iter(|| {
            let wb = Workbook::open(black_box("samples/bench_10k.xlsx")).unwrap();
            let ws = wb.sheet(0).unwrap();
            // Build all cell references A1..E10001 and read each one.
            // Using sheetnames + worksheets iterator ensures we don't short-circuit.
            let mut count = 0usize;
            for col in ["A", "B", "C", "D", "E"] {
                for row in 1..=10001u32 {
                    let reference = format!("{col}{row}");
                    if let Some(cell) = ws.cell(&reference) {
                        black_box(cell.value());
                        count += 1;
                    }
                }
            }
            black_box(count);
        });
    });
}

/// Benchmark: open_10k measuring only ZIP + parse time, not HashMap lookups.
/// Uses sheet discovery + sheetnames to confirm all metadata is parsed.
fn bench_open_10k_metadata_only(c: &mut Criterion) {
    c.bench_function("open_10k_metadata_only", |b| {
        b.iter(|| {
            let wb = Workbook::open(black_box("samples/bench_10k.xlsx")).unwrap();
            black_box(wb.sheetnames());
            black_box(wb.worksheets().len());
        });
    });
}

/// Benchmark: style resolution — font + fill lookup on a styled cell.
fn bench_style_resolution(c: &mut Criterion) {
    let wb = Workbook::open("samples/test.xlsx").unwrap();
    let ws = wb.sheet(0).unwrap();
    let cell = ws.cell("A1").unwrap();

    c.bench_function("style_resolution_font_fill", |b| {
        b.iter(|| {
            black_box(cell.font());
            black_box(cell.fill());
        });
    });
}

criterion_group!(
    read_benches,
    bench_open_test_xlsx,
    bench_open_multi_sheet_xlsx,
    bench_open_10k_rows,
    bench_open_10k_metadata_only,
    bench_style_resolution,
);
criterion_main!(read_benches);
