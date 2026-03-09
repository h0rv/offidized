//! Benchmarks for offidized-xlview parsing and conversion performance.
//!
//! Benchmarks the full pipeline: xlsx bytes -> offidized_xlsx::Workbook -> xlview Workbook.
//!
//! Run with: cargo bench -p offidized-xlview
//!
//! Results are saved to `target/criterion/` with HTML reports.
#![allow(
    clippy::expect_used,
    clippy::expect_fun_call,
    clippy::cast_possible_truncation
)]

use criterion::{black_box, criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};
use std::fs;

/// Benchmark parsing + conversion of the minimal test file.
fn bench_minimal(c: &mut Criterion) {
    let path = "crates/offidized-xlview/test/minimal.xlsx";
    if !std::path::Path::new(path).exists() {
        eprintln!("Skipping minimal benchmark - {} not found", path);
        return;
    }

    let data = fs::read(path).expect("Failed to read minimal.xlsx");

    c.bench_function("parse_minimal", |b| {
        b.iter(|| {
            let wb =
                offidized_xlsx::Workbook::from_bytes(black_box(&data)).expect("Failed to parse");
            let _viewer_wb = offidized_xlview::adapter::convert_workbook(&wb);
        })
    });
}

/// Benchmark parsing + conversion of the styled test file.
fn bench_styled(c: &mut Criterion) {
    let path = "crates/offidized-xlview/test/styled.xlsx";
    if !std::path::Path::new(path).exists() {
        eprintln!("Skipping styled benchmark - {} not found", path);
        return;
    }

    let data = fs::read(path).expect("Failed to read styled.xlsx");

    c.bench_function("parse_styled", |b| {
        b.iter(|| {
            let wb =
                offidized_xlsx::Workbook::from_bytes(black_box(&data)).expect("Failed to parse");
            let _viewer_wb = offidized_xlview::adapter::convert_workbook(&wb);
        })
    });
}

/// Benchmark parsing + conversion of the kitchen sink file (comprehensive features).
fn bench_kitchen_sink(c: &mut Criterion) {
    let path = "crates/offidized-xlview/test/kitchen_sink.xlsx";
    if !std::path::Path::new(path).exists() {
        eprintln!("Skipping kitchen_sink benchmark - {} not found", path);
        return;
    }

    let data = fs::read(path).expect("Failed to read kitchen_sink.xlsx");

    c.bench_function("parse_kitchen_sink", |b| {
        b.iter(|| {
            let wb =
                offidized_xlsx::Workbook::from_bytes(black_box(&data)).expect("Failed to parse");
            let _viewer_wb = offidized_xlview::adapter::convert_workbook(&wb);
        })
    });
}

/// Benchmark parsing + conversion of kitchen_sink_v2 (more features).
fn bench_kitchen_sink_v2(c: &mut Criterion) {
    let path = "crates/offidized-xlview/test/kitchen_sink_v2.xlsx";
    if !std::path::Path::new(path).exists() {
        eprintln!("Skipping kitchen_sink_v2 benchmark - {} not found", path);
        return;
    }

    let data = fs::read(path).expect("Failed to read kitchen_sink_v2.xlsx");

    c.bench_function("parse_kitchen_sink_v2", |b| {
        b.iter(|| {
            let wb =
                offidized_xlsx::Workbook::from_bytes(black_box(&data)).expect("Failed to parse");
            let _viewer_wb = offidized_xlview::adapter::convert_workbook(&wb);
        })
    });
}

/// Benchmark parsing + conversion of the large file (5000 rows x 20 cols) with throughput.
fn bench_large_file(c: &mut Criterion) {
    let path = "crates/offidized-xlview/test/large_5000x20.xlsx";
    if !std::path::Path::new(path).exists() {
        eprintln!("Skipping large file benchmark - {} not found", path);
        return;
    }

    let data = fs::read(path).expect("Failed to read large_5000x20.xlsx");
    let size = data.len();

    let mut group = c.benchmark_group("large_file");
    group.throughput(Throughput::Bytes(size as u64));

    group.bench_function("parse_5000x20", |b| {
        b.iter(|| {
            let wb =
                offidized_xlsx::Workbook::from_bytes(black_box(&data)).expect("Failed to parse");
            let _viewer_wb = offidized_xlview::adapter::convert_workbook(&wb);
        })
    });

    group.finish();
}

/// Benchmark parsing + conversion of the conditional formatting samples with throughput.
fn bench_cf_samples(c: &mut Criterion) {
    let path = "crates/offidized-xlview/test/ms_cf_samples.xlsx";
    if !std::path::Path::new(path).exists() {
        eprintln!("Skipping CF samples benchmark - {} not found", path);
        return;
    }

    let data = fs::read(path).expect("Failed to read ms_cf_samples.xlsx");
    let size = data.len();

    let mut group = c.benchmark_group("cf_samples");
    group.throughput(Throughput::Bytes(size as u64));

    group.bench_function("parse_ms_cf_samples", |b| {
        b.iter(|| {
            let wb =
                offidized_xlsx::Workbook::from_bytes(black_box(&data)).expect("Failed to parse");
            let _viewer_wb = offidized_xlview::adapter::convert_workbook(&wb);
        })
    });

    group.finish();
}

/// Compare parsing + conversion performance across file sizes with throughput.
fn bench_file_sizes(c: &mut Criterion) {
    let files = [
        ("minimal", "crates/offidized-xlview/test/minimal.xlsx"),
        ("styled", "crates/offidized-xlview/test/styled.xlsx"),
        (
            "kitchen_sink",
            "crates/offidized-xlview/test/kitchen_sink.xlsx",
        ),
        (
            "kitchen_sink_v2",
            "crates/offidized-xlview/test/kitchen_sink_v2.xlsx",
        ),
    ];

    let mut group = c.benchmark_group("file_size_comparison");

    for (name, path) in files {
        if !std::path::Path::new(path).exists() {
            continue;
        }

        let data = fs::read(path).expect(&format!("Failed to read {}", path));
        let size = data.len();

        group.throughput(Throughput::Bytes(size as u64));
        group.bench_with_input(BenchmarkId::new("parse", name), &data, |b, data| {
            b.iter(|| {
                let wb =
                    offidized_xlsx::Workbook::from_bytes(black_box(data)).expect("Failed to parse");
                let _viewer_wb = offidized_xlview::adapter::convert_workbook(&wb);
            })
        });
    }

    group.finish();
}

criterion_group!(
    benches,
    bench_minimal,
    bench_styled,
    bench_kitchen_sink,
    bench_kitchen_sink_v2,
    bench_large_file,
    bench_cf_samples,
    bench_file_sizes,
);

criterion_main!(benches);
