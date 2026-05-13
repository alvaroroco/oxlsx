use criterion::{Criterion, black_box, criterion_group, criterion_main};
use oxlsx::{Value, Workbook};
use std::env;

fn temp_path(name: &str) -> std::path::PathBuf {
    env::temp_dir().join(name)
}

/// Benchmark: create a new workbook with a single cell and save it.
fn bench_write_single_cell(c: &mut Criterion) {
    let path = temp_path("bench_single_cell.xlsx");
    c.bench_function("write_single_cell", |b| {
        b.iter(|| {
            let mut wb = Workbook::new();
            wb.active_mut().unwrap().set_cell("A1", black_box("Hello"));
            wb.save(black_box(&path)).unwrap();
        });
    });
    std::fs::remove_file(&path).ok();
}

/// Benchmark: write 1 000 rows × 5 columns (5k cells) to a new workbook.
fn bench_write_1k_rows(c: &mut Criterion) {
    let path = temp_path("bench_write_1k.xlsx");
    c.bench_function("write_1k_rows", |b| {
        b.iter(|| {
            let mut wb = Workbook::new();
            let ws = wb.active_mut().unwrap();
            let cols = ["A", "B", "C", "D", "E"];
            for row in 1u32..=1000 {
                for col in &cols {
                    let reference = format!("{col}{row}");
                    ws.set_cell(&reference, Value::Number(row as f64));
                }
            }
            wb.save(black_box(&path)).unwrap();
        });
    });
    std::fs::remove_file(&path).ok();
}

/// Benchmark: write 10 000 rows × 5 columns (50k cells) — main comparison vs openpyxl.
fn bench_write_10k_rows(c: &mut Criterion) {
    let path = temp_path("bench_write_10k.xlsx");
    c.bench_function("write_10k_rows", |b| {
        b.iter(|| {
            let mut wb = Workbook::new();
            let ws = wb.active_mut().unwrap();
            let cols = ["A", "B", "C", "D", "E"];
            for row in 1u32..=10000 {
                for col in &cols {
                    let reference = format!("{col}{row}");
                    ws.set_cell(&reference, Value::Number(row as f64));
                }
            }
            wb.save(black_box(&path)).unwrap();
        });
    });
    std::fs::remove_file(&path).ok();
}

/// Benchmark: write a multi-sheet workbook (3 sheets × 1 000 rows each).
fn bench_write_multi_sheet(c: &mut Criterion) {
    let path = temp_path("bench_write_multi.xlsx");
    c.bench_function("write_multi_sheet_3x1k", |b| {
        b.iter(|| {
            let mut wb = Workbook::new();
            for sheet_n in 0..3 {
                let ws = if sheet_n == 0 {
                    wb.active_mut().unwrap()
                } else {
                    wb.create_sheet(&format!("Sheet{}", sheet_n + 1))
                };
                for row in 1u32..=1000 {
                    ws.set_cell(&format!("A{row}"), Value::String(format!("row {row}")));
                    ws.set_cell(&format!("B{row}"), Value::Number(row as f64));
                }
            }
            wb.save(black_box(&path)).unwrap();
        });
    });
    std::fs::remove_file(&path).ok();
}

criterion_group!(
    write_benches,
    bench_write_single_cell,
    bench_write_1k_rows,
    bench_write_10k_rows,
    bench_write_multi_sheet,
);
criterion_main!(write_benches);
