//! Simple profiling for PML (PowerPoint) comparison stages.
//!
//! Run with: cargo bench --bench profile_pml

use std::fs;
use std::path::Path;
use std::time::Instant;

use redline_core::{PmlComparer, PmlDocument, PmlComparerSettings, reset_lcs_counters, get_lcs_counters};

fn main() {
    let base_path = Path::new("/Users/evan/development/openxml-worktree/rust-port-phase0/TestFiles");

    let file1 = base_path.join("Presentation.pptx");
    let file2 = base_path.join("Presentation.pptx");

    println!("=== PML Large Presentation Profiling ===\n");

    // Measure file reading
    let start = Instant::now();
    let bytes1 = fs::read(&file1).expect("Failed to read file1");
    let bytes2 = fs::read(&file2).expect("Failed to read file2");
    let read_time = start.elapsed();
    println!("File reading:    {:>8.2?} ({} bytes)", read_time, bytes1.len());

    // Measure document parsing
    let start = Instant::now();
    let doc1 = PmlDocument::from_bytes(&bytes1).unwrap();
    let parse1_time = start.elapsed();

    let start = Instant::now();
    let doc2 = PmlDocument::from_bytes(&bytes2).unwrap();
    let parse2_time = start.elapsed();
    println!("Document parse:  {:>8.2?} (doc1: {:?}, doc2: {:?})", parse1_time + parse2_time, parse1_time, parse2_time);

    // Measure comparison (multiple runs for averaging)
    let settings = PmlComparerSettings::default();

    // Warm-up run
    let _ = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();

    // Timed runs
    let mut times = Vec::new();
    reset_lcs_counters();
    for _ in 0..5 {
        let start = Instant::now();
        let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
        times.push(start.elapsed());

        // Print result stats on first run
        if times.len() == 1 {
            println!("\nResult: {} total changes (slides: +{}-{}, shapes: +{}-{})",
                result.total_changes, result.slides_inserted, result.slides_deleted,
                result.shapes_inserted, result.shapes_deleted);
        }
    }

    let avg_time = times.iter().sum::<std::time::Duration>() / times.len() as u32;
    let min_time = times.iter().min().unwrap();
    let max_time = times.iter().max().unwrap();

    println!("\nComparison timing (5 runs):");
    println!("  Average:       {:>8.2?}", avg_time);
    println!("  Min:           {:>8.2?}", min_time);
    println!("  Max:           {:>8.2?}", max_time);

    let total_bytes = bytes1.len() as f64;
    let throughput_kbps = (total_bytes / 1024.0) / avg_time.as_secs_f64();
    println!("  Throughput:    {:>8.2} KB/s", throughput_kbps);

    // Additional comparison cases
    println!("\n=== Additional Test Cases ===\n");

    // Test 1: Identical documents (best case)
    reset_lcs_counters();
    let start = Instant::now();
    let result = PmlComparer::compare(&doc1, &doc1, Some(&settings)).unwrap();
    let identical_time = start.elapsed();
    println!("Identical docs:  {:>8.2?} ({} total changes)",
        identical_time, result.total_changes);
    println!("  LCS counters:  {}", get_lcs_counters());

    // Test 2: Small file comparison
    let small1 = base_path.join("PmlComparer-Base.pptx");
    let small2 = base_path.join("PmlComparer-ShapeAdded.pptx");
    if small1.exists() && small2.exists() {
        let b1 = fs::read(&small1).unwrap();
        let b2 = fs::read(&small2).unwrap();
        let d1 = PmlDocument::from_bytes(&b1).unwrap();
        let d2 = PmlDocument::from_bytes(&b2).unwrap();

        let start = Instant::now();
        let result = PmlComparer::compare(&d1, &d2, Some(&settings)).unwrap();
        let small_time = start.elapsed();
        println!("Small ({} bytes): {:>8.2?} ({} changes)",
            b1.len(), small_time, result.total_changes);
    }
}
