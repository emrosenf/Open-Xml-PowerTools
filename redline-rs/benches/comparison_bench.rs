use criterion::{black_box, criterion_group, criterion_main, Criterion};

fn wml_comparison_benchmark(c: &mut Criterion) {
    c.bench_function("wml_parse_small", |b| {
        b.iter(|| {
            black_box(1 + 1)
        })
    });
}

fn sml_comparison_benchmark(c: &mut Criterion) {
    c.bench_function("sml_parse_small", |b| {
        b.iter(|| {
            black_box(2 + 2)
        })
    });
}

fn pml_comparison_benchmark(c: &mut Criterion) {
    c.bench_function("pml_parse_small", |b| {
        b.iter(|| {
            black_box(3 + 3)
        })
    });
}

criterion_group!(
    benches,
    wml_comparison_benchmark,
    sml_comparison_benchmark,
    pml_comparison_benchmark
);
criterion_main!(benches);
