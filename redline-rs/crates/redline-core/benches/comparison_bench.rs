//! Benchmarks for SML (Excel) and PML (PowerPoint) comparison.
//!
//! Run with: cargo bench --bench comparison_bench

use criterion::{criterion_group, criterion_main, Criterion, Throughput};
use redline_core::sml::{SmlComparer, SmlComparerSettings, SmlDocument};
use redline_core::pml::{PmlComparer, PmlComparerSettings, PmlDocument};
use std::fs;
use std::path::Path;

fn sml_benchmark(c: &mut Criterion) {
    let mut group = c.benchmark_group("sml_comparer");

    let base_path = Path::new("/Users/evan/development/openxml-worktree/rust-port-phase0/TestFiles");

    // Test cases for Excel comparison
    let test_cases = [
        ("Identical", "SH001-Table.xlsx", "SH001-Table.xlsx"),
        ("SmallDiff", "SH001-Table.xlsx", "SH007-One-Cell-Table.xlsx"),
        ("MultiSheet", "SH001-Table.xlsx", "SH002-TwoTablesTwoSheets.xlsx"),
    ];

    for (name, path1, path2) in test_cases {
        let file1 = base_path.join(path1);
        let file2 = base_path.join(path2);

        let bytes1 = match fs::read(&file1) {
            Ok(b) => b,
            Err(e) => {
                eprintln!("Skipping {}: {}", name, e);
                continue;
            }
        };
        let bytes2 = match fs::read(&file2) {
            Ok(b) => b,
            Err(e) => {
                eprintln!("Skipping {}: {}", name, e);
                continue;
            }
        };

        let doc1 = match SmlDocument::from_bytes(&bytes1) {
            Ok(d) => d,
            Err(e) => {
                eprintln!("Skipping {}: parse error: {}", name, e);
                continue;
            }
        };
        let doc2 = match SmlDocument::from_bytes(&bytes2) {
            Ok(d) => d,
            Err(e) => {
                eprintln!("Skipping {}: parse error: {}", name, e);
                continue;
            }
        };

        let settings = SmlComparerSettings::default();

        group.throughput(Throughput::Bytes((bytes1.len() + bytes2.len()) as u64));
        group.sample_size(10);
        group.measurement_time(std::time::Duration::from_secs(3));

        group.bench_function(name, |b| {
            b.iter(|| SmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap())
        });
    }

    group.finish();
}

fn pml_benchmark(c: &mut Criterion) {
    let mut group = c.benchmark_group("pml_comparer");

    let base_path = Path::new("/Users/evan/development/openxml-worktree/rust-port-phase0/TestFiles");

    // Test cases for PowerPoint comparison
    let test_cases = [
        ("Identical", "PmlComparer-Base.pptx", "PmlComparer-Identical.pptx"),
        ("ShapeAdded", "PmlComparer-Base.pptx", "PmlComparer-ShapeAdded.pptx"),
        ("SlideAdded", "PmlComparer-Base.pptx", "PmlComparer-SlideAdded.pptx"),
        ("TextChanged", "PmlComparer-Base.pptx", "PmlComparer-TextChanged.pptx"),
    ];

    for (name, path1, path2) in test_cases {
        let file1 = base_path.join(path1);
        let file2 = base_path.join(path2);

        let bytes1 = match fs::read(&file1) {
            Ok(b) => b,
            Err(e) => {
                eprintln!("Skipping {}: {}", name, e);
                continue;
            }
        };
        let bytes2 = match fs::read(&file2) {
            Ok(b) => b,
            Err(e) => {
                eprintln!("Skipping {}: {}", name, e);
                continue;
            }
        };

        let doc1 = match PmlDocument::from_bytes(&bytes1) {
            Ok(d) => d,
            Err(e) => {
                eprintln!("Skipping {}: parse error: {}", name, e);
                continue;
            }
        };
        let doc2 = match PmlDocument::from_bytes(&bytes2) {
            Ok(d) => d,
            Err(e) => {
                eprintln!("Skipping {}: parse error: {}", name, e);
                continue;
            }
        };

        let settings = PmlComparerSettings::default();

        group.throughput(Throughput::Bytes((bytes1.len() + bytes2.len()) as u64));
        group.sample_size(10);
        group.measurement_time(std::time::Duration::from_secs(3));

        group.bench_function(name, |b| {
            b.iter(|| PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap())
        });
    }

    group.finish();
}

criterion_group!(benches, sml_benchmark, pml_benchmark);
criterion_main!(benches);
