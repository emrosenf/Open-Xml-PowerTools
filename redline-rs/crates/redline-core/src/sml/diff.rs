// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

//! SmlDiffEngine - Computes differences between two WorkbookSignatures.
//!
//! This module implements the core comparison logic for Excel spreadsheets, including:
//! - Sheet matching (exact name match + rename detection via content similarity)
//! - Row alignment using LCS (Longest Common Subsequence) algorithm
//! - Cell-by-cell comparison (values, formulas, formatting)
//! - Phase 3 features: comments, data validations, merged cells, hyperlinks, named ranges
//!
//! 100% parity with C# SmlDiffEngine implementation.

use crate::sml::result::SmlComparisonResult;
use crate::sml::settings::SmlComparerSettings;
use crate::sml::signatures::{
    CellSignature, WorkbookSignature, WorksheetSignature,
};
use crate::sml::types::{SmlChange, SmlChangeType};
use std::collections::{HashMap, HashSet};

/// Main entry point for diff computation.
pub(crate) fn compute_diff(
    sig1: &WorkbookSignature,
    sig2: &WorkbookSignature,
    settings: &SmlComparerSettings,
) -> SmlComparisonResult {
    let mut result = SmlComparisonResult::new();

    // Build sheet matching (handles renames)
    let sheet_matches = match_sheets(sig1, sig2, settings);

    if settings.compare_sheet_structure {
        // Report sheet-level changes
        for m in &sheet_matches {
            match m.match_type {
                SheetMatchType::Added => {
                    result.add_change(SmlChange {
                        change_type: SmlChangeType::SheetAdded,
                        sheet_name: Some(m.new_name.clone()),
                        ..Default::default()
                    });
                }
                SheetMatchType::Deleted => {
                    result.add_change(SmlChange {
                        change_type: SmlChangeType::SheetDeleted,
                        sheet_name: m.old_name.clone(),
                        ..Default::default()
                    });
                }
                SheetMatchType::Renamed => {
                    result.add_change(SmlChange {
                        change_type: SmlChangeType::SheetRenamed,
                        sheet_name: Some(m.new_name.clone()),
                        old_sheet_name: m.old_name.clone(),
                        ..Default::default()
                    });
                }
                SheetMatchType::Matched => {}
            }
        }
    }

    // Compare matched sheets (including renamed ones)
    for m in sheet_matches.iter().filter(|m| {
        matches!(
            m.match_type,
            SheetMatchType::Matched | SheetMatchType::Renamed
        )
    }) {
        let old_name = m.old_name.as_ref().unwrap();
        let ws1 = &sig1.sheets[old_name];
        let ws2 = &sig2.sheets[&m.new_name];

        if settings.enable_row_alignment {
            compare_worksheets_with_alignment(ws1, ws2, &m.new_name, settings, &mut result);
        } else {
            compare_worksheets_cell_by_cell(ws1, ws2, &m.new_name, settings, &mut result);
        }

        // Phase 3: Compare comments
        if settings.compare_comments {
            compare_comments(ws1, ws2, &m.new_name, &mut result);
        }

        // Phase 3: Compare data validations
        if settings.compare_data_validation {
            compare_data_validations(ws1, ws2, &m.new_name, &mut result);
        }

        // Phase 3: Compare merged cells
        if settings.compare_merged_cells {
            compare_merged_cells(ws1, ws2, &m.new_name, &mut result);
        }

        // Phase 3: Compare hyperlinks
        if settings.compare_hyperlinks {
            compare_hyperlinks(ws1, ws2, &m.new_name, &mut result);
        }
    }

    // Phase 3: Compare named ranges at workbook level
    if settings.compare_named_ranges {
        compare_named_ranges(sig1, sig2, &mut result);
    }

    result
}

// ============================================================================
// Sheet Matching
// ============================================================================

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SheetMatchType {
    Matched,
    Added,
    Deleted,
    Renamed,
}

#[derive(Debug, Clone)]
struct SheetMatch {
    match_type: SheetMatchType,
    old_name: Option<String>,
    new_name: String,
    similarity: f64,
}

fn match_sheets(
    sig1: &WorkbookSignature,
    sig2: &WorkbookSignature,
    settings: &SmlComparerSettings,
) -> Vec<SheetMatch> {
    let mut matches = Vec::new();
    let sheets1: HashSet<_> = sig1.sheets.keys().cloned().collect();
    let sheets2: HashSet<_> = sig2.sheets.keys().cloned().collect();

    // Exact name matches
    let common_sheets: Vec<_> = sheets1.intersection(&sheets2).cloned().collect();
    for name in &common_sheets {
        matches.push(SheetMatch {
            match_type: SheetMatchType::Matched,
            old_name: Some(name.clone()),
            new_name: name.clone(),
            similarity: 1.0,
        });
    }

    let mut unmatched1: Vec<_> = sheets1.difference(&common_sheets.iter().cloned().collect()).cloned().collect();
    let mut unmatched2: Vec<_> = sheets2.difference(&common_sheets.iter().cloned().collect()).cloned().collect();

    // Try to detect renames based on content similarity
    if settings.enable_sheet_rename_detection && !unmatched1.is_empty() && !unmatched2.is_empty() {
        let renamed = detect_renamed_sheets(sig1, sig2, &unmatched1, &unmatched2, settings);
        
        // Remove matched sheets from unmatched lists
        for r in &renamed {
            if let Some(old) = &r.old_name {
                unmatched1.retain(|n| n != old);
            }
            unmatched2.retain(|n| n != &r.new_name);
        }
        
        matches.extend(renamed);
    }

    // Remaining unmatched sheets are added/deleted
    for deleted in unmatched1 {
        matches.push(SheetMatch {
            match_type: SheetMatchType::Deleted,
            old_name: Some(deleted),
            new_name: String::new(),
            similarity: 0.0,
        });
    }

    for added in unmatched2 {
        matches.push(SheetMatch {
            match_type: SheetMatchType::Added,
            old_name: None,
            new_name: added,
            similarity: 0.0,
        });
    }

    matches
}

fn detect_renamed_sheets(
    sig1: &WorkbookSignature,
    sig2: &WorkbookSignature,
    unmatched1: &[String],
    unmatched2: &[String],
    settings: &SmlComparerSettings,
) -> Vec<SheetMatch> {
    let mut renames = Vec::new();
    let mut used1 = HashSet::new();
    let mut used2 = HashSet::new();

    // Compute content hashes for unmatched sheets
    let hashes1: HashMap<_, _> = unmatched1
        .iter()
        .map(|n| (n.clone(), sig1.sheets[n].compute_content_hash()))
        .collect();
    let hashes2: HashMap<_, _> = unmatched2
        .iter()
        .map(|n| (n.clone(), sig2.sheets[n].compute_content_hash()))
        .collect();

    // First pass: exact content match (definite rename)
    for name1 in unmatched1 {
        let hash1 = &hashes1[name1];
        if let Some(exact_match) = unmatched2
            .iter()
            .find(|n2| !used2.contains(*n2) && &hashes2[*n2] == hash1)
        {
            renames.push(SheetMatch {
                match_type: SheetMatchType::Renamed,
                old_name: Some(name1.clone()),
                new_name: exact_match.clone(),
                similarity: 1.0,
            });
            used1.insert(name1.clone());
            used2.insert(exact_match.clone());
        }
    }

    // Second pass: similarity-based matching
    // Collect keys to avoid simultaneous borrow issues
    let unmatched1_filtered: Vec<_> = unmatched1.iter()
        .filter(|n| !used1.contains(*n))
        .cloned()
        .collect();
    
    for name1 in unmatched1_filtered {
        let ws1 = &sig1.sheets[&name1];
        let mut best_similarity = 0.0;
        let mut best_match: Option<String> = None;

        for name2 in unmatched2.iter().filter(|n| !used2.contains(*n)) {
            let ws2 = &sig2.sheets[name2];
            let similarity = compute_sheet_similarity(ws1, ws2);

            if similarity > best_similarity && similarity >= settings.sheet_rename_similarity_threshold {
                best_similarity = similarity;
                best_match = Some(name2.clone());
            }
        }

        if let Some(best) = best_match {
            renames.push(SheetMatch {
                match_type: SheetMatchType::Renamed,
                old_name: Some(name1.clone()),
                new_name: best.clone(),
                similarity: best_similarity,
            });
            used1.insert(name1.clone());
            used2.insert(best);
        }
    }

    renames
}

fn compute_sheet_similarity(ws1: &WorksheetSignature, ws2: &WorksheetSignature) -> f64 {
    // Jaccard similarity on cell addresses with matching values
    if ws1.cells.is_empty() && ws2.cells.is_empty() {
        return 1.0;
    }
    if ws1.cells.is_empty() || ws2.cells.is_empty() {
        return 0.0;
    }

    let all_addresses: HashSet<_> = ws1.cells.keys().chain(ws2.cells.keys()).collect();
    let mut matching_count = 0;

    for addr in &all_addresses {
        if let (Some(c1), Some(c2)) = (ws1.cells.get(*addr), ws2.cells.get(*addr)) {
            if c1.resolved_value == c2.resolved_value {
                matching_count += 1;
            }
        }
    }

    matching_count as f64 / all_addresses.len() as f64
}

// ============================================================================
// Row Alignment (Phase 2)
// ============================================================================

fn compare_worksheets_with_alignment(
    ws1: &WorksheetSignature,
    ws2: &WorksheetSignature,
    sheet_name: &str,
    settings: &SmlComparerSettings,
    result: &mut SmlComparisonResult,
) {
    // Get row alignment using LCS
    let rows1: Vec<_> = ws1.populated_rows.iter().copied().collect();
    let rows2: Vec<_> = ws2.populated_rows.iter().copied().collect();

    let row_alignment = compute_row_alignment(ws1, ws2, &rows1, &rows2);

    // Report inserted/deleted rows
    for (old_row, new_row) in &row_alignment {
        match (old_row, new_row) {
            (None, Some(new_r)) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::RowInserted,
                    sheet_name: Some(sheet_name.to_string()),
                    row_index: Some(*new_r),
                    ..Default::default()
                });
            }
            (Some(old_r), None) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::RowDeleted,
                    sheet_name: Some(sheet_name.to_string()),
                    row_index: Some(*old_r),
                    ..Default::default()
                });
            }
            (Some(old_r), Some(new_r)) => {
                // Aligned rows - compare cells within the row
                compare_aligned_rows(ws1, ws2, *old_r, *new_r, sheet_name, settings, result);
            }
            (None, None) => {} // Should not happen
        }
    }
}

fn compute_row_alignment(
    ws1: &WorksheetSignature,
    ws2: &WorksheetSignature,
    rows1: &[i32],
    rows2: &[i32],
) -> Vec<(Option<i32>, Option<i32>)> {
    // Get row signatures
    let sigs1: Vec<_> = rows1
        .iter()
        .map(|r| ws1.row_signatures.get(r).cloned().unwrap_or_default())
        .collect();
    let sigs2: Vec<_> = rows2
        .iter()
        .map(|r| ws2.row_signatures.get(r).cloned().unwrap_or_default())
        .collect();

    // Compute LCS
    let lcs = compute_lcs(&sigs1, &sigs2);

    // Build alignment from LCS
    let mut alignment = Vec::new();
    let mut i = 0;
    let mut j = 0;
    let mut k = 0;

    while i < rows1.len() || j < rows2.len() {
        if k < lcs.len() && i < rows1.len() && sigs1[i] == lcs[k] {
            // Find matching position in rows2
            while j < rows2.len() && sigs2[j] != lcs[k] {
                // Row inserted in newer
                alignment.push((None, Some(rows2[j])));
                j += 1;
            }

            if j < rows2.len() {
                // Matched row
                alignment.push((Some(rows1[i]), Some(rows2[j])));
                i += 1;
                j += 1;
                k += 1;
            }
        } else if i < rows1.len() {
            // Row deleted from older
            alignment.push((Some(rows1[i]), None));
            i += 1;
        } else if j < rows2.len() {
            // Row inserted in newer
            alignment.push((None, Some(rows2[j])));
            j += 1;
        }
    }

    alignment
}

fn compute_lcs(seq1: &[String], seq2: &[String]) -> Vec<String> {
    let m = seq1.len();
    let n = seq2.len();

    // DP table
    let mut dp = vec![vec![0; n + 1]; m + 1];

    for i in 1..=m {
        for j in 1..=n {
            if seq1[i - 1] == seq2[j - 1] {
                dp[i][j] = dp[i - 1][j - 1] + 1;
            } else {
                dp[i][j] = dp[i - 1][j].max(dp[i][j - 1]);
            }
        }
    }

    // Backtrack to find LCS
    let mut lcs = Vec::new();
    let mut ii = m;
    let mut jj = n;

    while ii > 0 && jj > 0 {
        if seq1[ii - 1] == seq2[jj - 1] {
            lcs.push(seq1[ii - 1].clone());
            ii -= 1;
            jj -= 1;
        } else if dp[ii - 1][jj] > dp[ii][jj - 1] {
            ii -= 1;
        } else {
            jj -= 1;
        }
    }

    lcs.reverse();
    lcs
}

fn compare_aligned_rows(
    ws1: &WorksheetSignature,
    ws2: &WorksheetSignature,
    row1: i32,
    row2: i32,
    sheet_name: &str,
    settings: &SmlComparerSettings,
    result: &mut SmlComparisonResult,
) {
    let cells1: HashMap<_, _> = ws1
        .get_cells_in_row(row1)
        .into_iter()
        .map(|c| (c.column, c))
        .collect();
    let cells2: HashMap<_, _> = ws2
        .get_cells_in_row(row2)
        .into_iter()
        .map(|c| (c.column, c))
        .collect();

    let all_columns: HashSet<_> = cells1.keys().chain(cells2.keys()).copied().collect();

    for col in all_columns {
        let has1 = cells1.get(&col);
        let has2 = cells2.get(&col);

        // Use the new address from the new row
        let new_addr = has2
            .map(|c| c.address.clone())
            .unwrap_or_else(|| get_cell_address(col, row2));
        let old_addr = has1
            .map(|c| c.address.clone())
            .unwrap_or_else(|| get_cell_address(col, row1));

        match (has1, has2) {
            (None, Some(cell2)) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::CellAdded,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(new_addr),
                    new_value: cell2.resolved_value.clone(),
                    new_formula: cell2.formula.clone(),
                    new_format: Some(cell2.format.clone()),
                    ..Default::default()
                });
            }
            (Some(cell1), None) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::CellDeleted,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(old_addr),
                    old_value: cell1.resolved_value.clone(),
                    old_formula: cell1.formula.clone(),
                    old_format: Some(cell1.format.clone()),
                    ..Default::default()
                });
            }
            (Some(cell1), Some(cell2)) => {
                // Compare cells - use new address for reporting
                compare_cells(cell1, cell2, sheet_name, settings, result, Some(&new_addr));
            }
            (None, None) => {} // Should not happen
        }
    }
}

fn get_cell_address(col: i32, row: i32) -> String {
    let mut col_letter = String::new();
    let mut c = col;
    while c > 0 {
        c -= 1;
        col_letter.insert(0, (b'A' + (c % 26) as u8) as char);
        c /= 26;
    }
    format!("{}{}", col_letter, row)
}

// ============================================================================
// Cell-by-Cell Comparison (Phase 1 fallback)
// ============================================================================

fn compare_worksheets_cell_by_cell(
    ws1: &WorksheetSignature,
    ws2: &WorksheetSignature,
    sheet_name: &str,
    settings: &SmlComparerSettings,
    result: &mut SmlComparisonResult,
) {
    // Get union of all cell addresses
    let all_addresses: HashSet<_> = ws1.cells.keys().chain(ws2.cells.keys()).collect();

    for addr in all_addresses {
        let has1 = ws1.cells.get(addr);
        let has2 = ws2.cells.get(addr);

        match (has1, has2) {
            (None, Some(cell2)) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::CellAdded,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(addr.clone()),
                    new_value: cell2.resolved_value.clone(),
                    new_formula: cell2.formula.clone(),
                    new_format: Some(cell2.format.clone()),
                    ..Default::default()
                });
            }
            (Some(cell1), None) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::CellDeleted,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(addr.clone()),
                    old_value: cell1.resolved_value.clone(),
                    old_formula: cell1.formula.clone(),
                    old_format: Some(cell1.format.clone()),
                    ..Default::default()
                });
            }
            (Some(cell1), Some(cell2)) => {
                compare_cells(cell1, cell2, sheet_name, settings, result, None);
            }
            (None, None) => {} // Should not happen
        }
    }
}

fn compare_cells(
    cell1: &CellSignature,
    cell2: &CellSignature,
    sheet_name: &str,
    settings: &SmlComparerSettings,
    result: &mut SmlComparisonResult,
    address_override: Option<&str>,
) {
    // Use the override address if provided, otherwise use the original cell address
    let report_address = address_override.unwrap_or(&cell1.address);

    // Quick check via content hash (value + formula)
    if cell1.content_hash == cell2.content_hash
        && (!settings.compare_formatting || cell1.format == cell2.format)
    {
        return; // No changes
    }

    // Check value change
    if settings.compare_values {
        let val1 = cell1.resolved_value.as_deref().unwrap_or("");
        let val2 = cell2.resolved_value.as_deref().unwrap_or("");

        let values_equal = if settings.case_insensitive_values {
            val1.eq_ignore_ascii_case(val2)
        } else if settings.numeric_tolerance > 0.0 {
            // Try numeric comparison with tolerance
            match (val1.parse::<f64>(), val2.parse::<f64>()) {
                (Ok(d1), Ok(d2)) => (d1 - d2).abs() <= settings.numeric_tolerance,
                _ => val1 == val2,
            }
        } else {
            val1 == val2
        };

        if !values_equal {
            result.add_change(SmlChange {
                change_type: SmlChangeType::ValueChanged,
                sheet_name: Some(sheet_name.to_string()),
                cell_address: Some(report_address.to_string()),
                old_value: cell1.resolved_value.clone(),
                new_value: cell2.resolved_value.clone(),
                old_formula: cell1.formula.clone(),
                new_formula: cell2.formula.clone(),
                ..Default::default()
            });
            return; // Don't report formula change if value changed
        }
    }

    // Check formula change
    if settings.compare_formulas {
        let formula1 = cell1.formula.as_deref().unwrap_or("");
        let formula2 = cell2.formula.as_deref().unwrap_or("");

        if formula1 != formula2 {
            result.add_change(SmlChange {
                change_type: SmlChangeType::FormulaChanged,
                sheet_name: Some(sheet_name.to_string()),
                cell_address: Some(report_address.to_string()),
                old_formula: cell1.formula.clone(),
                new_formula: cell2.formula.clone(),
                old_value: cell1.resolved_value.clone(),
                new_value: cell2.resolved_value.clone(),
                ..Default::default()
            });
            return; // Don't report format change if formula changed
        }
    }

    // Check format change
    if settings.compare_formatting && cell1.format != cell2.format {
        result.add_change(SmlChange {
            change_type: SmlChangeType::FormatChanged,
            sheet_name: Some(sheet_name.to_string()),
            cell_address: Some(report_address.to_string()),
            old_format: Some(cell1.format.clone()),
            new_format: Some(cell2.format.clone()),
            old_value: cell1.resolved_value.clone(),
            new_value: cell2.resolved_value.clone(),
            ..Default::default()
        });
    }
}

// ============================================================================
// Phase 3: Comparison Methods
// ============================================================================

fn compare_named_ranges(
    sig1: &WorkbookSignature,
    sig2: &WorkbookSignature,
    result: &mut SmlComparisonResult,
) {
    let all_names: HashSet<_> = sig1
        .defined_names
        .keys()
        .chain(sig2.defined_names.keys())
        .collect();

    for name in all_names {
        let has1 = sig1.defined_names.get(name);
        let has2 = sig2.defined_names.get(name);

        match (has1, has2) {
            (None, Some(value2)) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::NamedRangeAdded,
                    named_range_name: Some(name.clone()),
                    new_named_range_value: Some(value2.clone()),
                    ..Default::default()
                });
            }
            (Some(value1), None) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::NamedRangeDeleted,
                    named_range_name: Some(name.clone()),
                    old_named_range_value: Some(value1.clone()),
                    ..Default::default()
                });
            }
            (Some(value1), Some(value2)) if value1 != value2 => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::NamedRangeChanged,
                    named_range_name: Some(name.clone()),
                    old_named_range_value: Some(value1.clone()),
                    new_named_range_value: Some(value2.clone()),
                    ..Default::default()
                });
            }
            _ => {}
        }
    }
}

fn compare_comments(
    ws1: &WorksheetSignature,
    ws2: &WorksheetSignature,
    sheet_name: &str,
    result: &mut SmlComparisonResult,
) {
    let all_addresses: HashSet<_> = ws1.comments.keys().chain(ws2.comments.keys()).collect();

    for addr in all_addresses {
        let has1 = ws1.comments.get(addr);
        let has2 = ws2.comments.get(addr);

        match (has1, has2) {
            (None, Some(comment2)) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::CommentAdded,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(addr.clone()),
                    new_comment: Some(comment2.text.clone()),
                    comment_author: Some(comment2.author.clone()),
                    ..Default::default()
                });
            }
            (Some(comment1), None) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::CommentDeleted,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(addr.clone()),
                    old_comment: Some(comment1.text.clone()),
                    comment_author: Some(comment1.author.clone()),
                    ..Default::default()
                });
            }
            (Some(comment1), Some(comment2))
                if comment1.text != comment2.text || comment1.author != comment2.author =>
            {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::CommentChanged,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(addr.clone()),
                    old_comment: Some(comment1.text.clone()),
                    new_comment: Some(comment2.text.clone()),
                    comment_author: Some(comment2.author.clone()),
                    ..Default::default()
                });
            }
            _ => {}
        }
    }
}

fn compare_data_validations(
    ws1: &WorksheetSignature,
    ws2: &WorksheetSignature,
    sheet_name: &str,
    result: &mut SmlComparisonResult,
) {
    let all_keys: HashSet<_> = ws1
        .data_validations
        .keys()
        .chain(ws2.data_validations.keys())
        .collect();

    for key in all_keys {
        let has1 = ws1.data_validations.get(key);
        let has2 = ws2.data_validations.get(key);

        match (has1, has2) {
            (None, Some(dv2)) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::DataValidationAdded,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(key.clone()),
                    data_validation_type: Some(dv2.validation_type.clone()),
                    new_data_validation: Some(dv2.to_string()),
                    ..Default::default()
                });
            }
            (Some(dv1), None) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::DataValidationDeleted,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(key.clone()),
                    data_validation_type: Some(dv1.validation_type.clone()),
                    old_data_validation: Some(dv1.to_string()),
                    ..Default::default()
                });
            }
            (Some(dv1), Some(dv2)) if dv1.compute_hash() != dv2.compute_hash() => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::DataValidationChanged,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(key.clone()),
                    data_validation_type: Some(dv2.validation_type.clone()),
                    old_data_validation: Some(dv1.to_string()),
                    new_data_validation: Some(dv2.to_string()),
                    ..Default::default()
                });
            }
            _ => {}
        }
    }
}

fn compare_merged_cells(
    ws1: &WorksheetSignature,
    ws2: &WorksheetSignature,
    sheet_name: &str,
    result: &mut SmlComparisonResult,
) {
    let all_ranges: HashSet<_> = ws1
        .merged_cell_ranges
        .iter()
        .chain(ws2.merged_cell_ranges.iter())
        .collect();

    for range in all_ranges {
        let has1 = ws1.merged_cell_ranges.contains(range);
        let has2 = ws2.merged_cell_ranges.contains(range);

        match (has1, has2) {
            (false, true) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::MergedCellAdded,
                    sheet_name: Some(sheet_name.to_string()),
                    merged_cell_range: Some(range.clone()),
                    ..Default::default()
                });
            }
            (true, false) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::MergedCellDeleted,
                    sheet_name: Some(sheet_name.to_string()),
                    merged_cell_range: Some(range.clone()),
                    ..Default::default()
                });
            }
            _ => {}
        }
    }
}

fn compare_hyperlinks(
    ws1: &WorksheetSignature,
    ws2: &WorksheetSignature,
    sheet_name: &str,
    result: &mut SmlComparisonResult,
) {
    let all_addresses: HashSet<_> = ws1.hyperlinks.keys().chain(ws2.hyperlinks.keys()).collect();

    for addr in all_addresses {
        let has1 = ws1.hyperlinks.get(addr);
        let has2 = ws2.hyperlinks.get(addr);

        match (has1, has2) {
            (None, Some(hl2)) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::HyperlinkAdded,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(addr.clone()),
                    new_hyperlink: Some(hl2.target.clone()),
                    ..Default::default()
                });
            }
            (Some(hl1), None) => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::HyperlinkDeleted,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(addr.clone()),
                    old_hyperlink: Some(hl1.target.clone()),
                    ..Default::default()
                });
            }
            (Some(hl1), Some(hl2)) if hl1.compute_hash() != hl2.compute_hash() => {
                result.add_change(SmlChange {
                    change_type: SmlChangeType::HyperlinkChanged,
                    sheet_name: Some(sheet_name.to_string()),
                    cell_address: Some(addr.clone()),
                    old_hyperlink: Some(hl1.target.clone()),
                    new_hyperlink: Some(hl2.target.clone()),
                    ..Default::default()
                });
            }
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sml::signatures::{CellFormatSignature, CellSignature, WorksheetSignature};

    #[test]
    fn compute_lcs_works() {
        let seq1 = vec!["A".to_string(), "B".to_string(), "C".to_string()];
        let seq2 = vec!["A".to_string(), "X".to_string(), "B".to_string(), "C".to_string()];
        let lcs = compute_lcs(&seq1, &seq2);
        assert_eq!(lcs, vec!["A", "B", "C"]);
    }

    #[test]
    fn get_cell_address_works() {
        assert_eq!(get_cell_address(1, 1), "A1");
        assert_eq!(get_cell_address(26, 1), "Z1");
        assert_eq!(get_cell_address(27, 1), "AA1");
        assert_eq!(get_cell_address(702, 5), "ZZ5");
    }

    #[test]
    fn compute_sheet_similarity_identical_sheets() {
        let mut ws1 = WorksheetSignature::new("Sheet1".to_string(), "rId1".to_string());
        ws1.cells.insert(
            "A1".to_string(),
            CellSignature {
                address: "A1".to_string(),
                row: 1,
                column: 1,
                resolved_value: Some("test".to_string()),
                formula: None,
                content_hash: String::new(),
                format: CellFormatSignature::default(),
            },
        );

        let ws2 = ws1.clone();
        let similarity = compute_sheet_similarity(&ws1, &ws2);
        assert!((similarity - 1.0).abs() < f64::EPSILON);
    }

    #[test]
    fn compute_sheet_similarity_different_sheets() {
        let mut ws1 = WorksheetSignature::new("Sheet1".to_string(), "rId1".to_string());
        ws1.cells.insert(
            "A1".to_string(),
            CellSignature {
                address: "A1".to_string(),
                row: 1,
                column: 1,
                resolved_value: Some("test1".to_string()),
                formula: None,
                content_hash: String::new(),
                format: CellFormatSignature::default(),
            },
        );

        let mut ws2 = WorksheetSignature::new("Sheet2".to_string(), "rId2".to_string());
        ws2.cells.insert(
            "A1".to_string(),
            CellSignature {
                address: "A1".to_string(),
                row: 1,
                column: 1,
                resolved_value: Some("test2".to_string()),
                formula: None,
                content_hash: String::new(),
                format: CellFormatSignature::default(),
            },
        );

        let similarity = compute_sheet_similarity(&ws1, &ws2);
        assert!((similarity - 0.0).abs() < f64::EPSILON);
    }

    #[test]
    fn match_sheets_exact_match() {
        let mut sig1 = WorkbookSignature::new();
        sig1.sheets.insert(
            "Sheet1".to_string(),
            WorksheetSignature::new("Sheet1".to_string(), "rId1".to_string()),
        );

        let mut sig2 = WorkbookSignature::new();
        sig2.sheets.insert(
            "Sheet1".to_string(),
            WorksheetSignature::new("Sheet1".to_string(), "rId1".to_string()),
        );

        let settings = SmlComparerSettings::default();
        let matches = match_sheets(&sig1, &sig2, &settings);

        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].match_type, SheetMatchType::Matched);
        assert_eq!(matches[0].new_name, "Sheet1");
    }

    #[test]
    fn compute_diff_empty_workbooks() {
        let sig1 = WorkbookSignature::new();
        let sig2 = WorkbookSignature::new();
        let settings = SmlComparerSettings::default();

        let result = compute_diff(&sig1, &sig2, &settings);
        assert_eq!(result.total_changes(), 0);
    }
}
