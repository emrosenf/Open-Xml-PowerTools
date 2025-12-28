use criterion::{criterion_group, criterion_main, Criterion, Throughput};
use redline_core::{WmlComparer, WmlDocument, WmlComparerSettings};
use std::fs;
use std::path::Path;

fn wml_benchmark(c: &mut Criterion) {
    println!("CWD: {:?}", std::env::current_dir());
    let mut group = c.benchmark_group("wml_comparer");
    
    // Test cases covering different complexity levels
    // Paths are relative to crate root
    let test_cases = [
        ("Basic", "WC/WC002-Unmodified.docx", "WC/WC002-InsertInMiddle.docx"),
        ("Table", "WC/WC006-Table.docx", "WC/WC006-Table-Delete-Row.docx"),
        ("Large", "WC/WC004-Large.docx", "WC/WC004-Large-Mod.docx"),
    ];

    let base_path = Path::new("/Users/evan/development/openxml-worktree/rust-port-phase0/TestFiles");

    for (name, path1, path2) in test_cases {
        let file1 = base_path.join(path1);
        let file2 = base_path.join(path2);

        // Read files once
        let bytes1 = fs::read(&file1).expect(&format!("Failed to read {}", file1.display()));
        let bytes2 = fs::read(&file2).expect(&format!("Failed to read {}", file2.display()));

        let doc1 = WmlDocument::from_bytes(&bytes1).unwrap();
        let doc2 = WmlDocument::from_bytes(&bytes2).unwrap();
        
        let settings = WmlComparerSettings::default();

        group.throughput(Throughput::Bytes((bytes1.len() + bytes2.len()) as u64));
        group.sample_size(10);
        group.measurement_time(std::time::Duration::from_secs(5));
        
        group.bench_function(name, |b| {
            b.iter(|| {
                WmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap()
            })
        });
    }

    group.finish();
}

criterion_group!(benches, wml_benchmark);
criterion_main!(benches);
