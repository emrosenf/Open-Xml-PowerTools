//! Simple profiling for WML comparison stages.
//!
//! Run with: cargo bench --bench profile_wml

use std::fs;
use std::path::Path;
use std::time::Instant;

use redline_core::{WmlComparer, WmlDocument, WmlComparerSettings, reset_lcs_counters, get_lcs_counters};

fn main() {
    let base_path = Path::new("/Users/evan/development/openxml-worktree/rust-port-phase0/TestFiles");

    let file1 = base_path.join("WC/WC004-Large.docx");
    let file2 = base_path.join("WC/WC004-Large-Mod.docx");

    println!("=== WML Large Document Profiling ===\n");

    // Measure file reading
    let start = Instant::now();
    let bytes1 = fs::read(&file1).expect("Failed to read file1");
    let bytes2 = fs::read(&file2).expect("Failed to read file2");
    let read_time = start.elapsed();
    println!("File reading:    {:>8.2?} ({} + {} bytes)", read_time, bytes1.len(), bytes2.len());

    // Measure document parsing
    let start = Instant::now();
    let doc1 = WmlDocument::from_bytes(&bytes1).unwrap();
    let parse1_time = start.elapsed();

    let start = Instant::now();
    let doc2 = WmlDocument::from_bytes(&bytes2).unwrap();
    let parse2_time = start.elapsed();
    println!("Document parse:  {:>8.2?} (doc1: {:?}, doc2: {:?})", parse1_time + parse2_time, parse1_time, parse2_time);

    // Measure comparison (multiple runs for averaging)
    let settings = WmlComparerSettings::default();

    // Warm-up run
    let _ = WmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();

    // Timed runs
    let mut times = Vec::new();
    reset_lcs_counters();
    for _ in 0..5 {
        let start = Instant::now();
        let result = WmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
        times.push(start.elapsed());

        // Print result stats on first run
        if times.len() == 1 {
            println!("\nResult: {} insertions, {} deletions, {} format changes",
                result.insertions, result.deletions, result.format_changes);
            println!("LCS counters: {}", get_lcs_counters());
        }
    }

    let avg_time = times.iter().sum::<std::time::Duration>() / times.len() as u32;
    let min_time = times.iter().min().unwrap();
    let max_time = times.iter().max().unwrap();

    println!("\nComparison timing (5 runs):");
    println!("  Average:       {:>8.2?}", avg_time);
    println!("  Min:           {:>8.2?}", min_time);
    println!("  Max:           {:>8.2?}", max_time);

    let total_bytes = (bytes1.len() + bytes2.len()) as f64;
    let throughput_kbps = (total_bytes / 1024.0) / avg_time.as_secs_f64();
    println!("  Throughput:    {:>8.2} KB/s", throughput_kbps);

    // Additional comparison cases
    println!("\n=== Additional Test Cases ===\n");

    // Test 1: Identical documents (best case)
    // Parse fresh copies to see parsing overhead
    let start = Instant::now();
    let doc1a = WmlDocument::from_bytes(&bytes1).unwrap();
    let doc1b = WmlDocument::from_bytes(&bytes1).unwrap();
    let dual_parse_time = start.elapsed();
    println!("Dual parse same: {:>8.2?}", dual_parse_time);

    reset_lcs_counters();
    let start = Instant::now();
    let result = WmlComparer::compare(&doc1a, &doc1b, Some(&settings)).unwrap();
    let identical_time = start.elapsed();
    println!("Identical docs:  {:>8.2?} ({} insertions, {} deletions)",
        identical_time, result.insertions, result.deletions);
    println!("  LCS counters:  {}", get_lcs_counters());

    // Test 2: Basic modification
    let basic1 = base_path.join("WC/WC002-Unmodified.docx");
    let basic2 = base_path.join("WC/WC002-InsertInMiddle.docx");
    if basic1.exists() && basic2.exists() {
        let b1 = fs::read(&basic1).unwrap();
        let b2 = fs::read(&basic2).unwrap();
        let d1 = WmlDocument::from_bytes(&b1).unwrap();
        let d2 = WmlDocument::from_bytes(&b2).unwrap();

        let start = Instant::now();
        let result = WmlComparer::compare(&d1, &d2, Some(&settings)).unwrap();
        let basic_time = start.elapsed();
        println!("Basic ({}+{} bytes): {:>8.2?} ({} ins, {} del)",
            b1.len(), b2.len(), basic_time, result.insertions, result.deletions);
    }
}
