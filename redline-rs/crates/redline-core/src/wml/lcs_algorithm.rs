//! Full LCS Algorithm for WmlComparer
//!
//! This is a faithful port of the C# LCS algorithm from WmlComparer.cs.
//!
//! The algorithm works in three phases:
//! 1. ProcessCorrelatedHashes - Match groups by pre-computed CorrelatedSHA1Hash
//! 2. FindCommonAtBeginningAndEnd - Find common prefix/suffix by SHA1Hash
//! 3. DoLcsAlgorithm - Full LCS with edge case handling
//!
//! Key insight: This processes `Unknown` CorrelatedSequences iteratively until
//! all are resolved to Equal, Deleted, or Inserted.

use super::comparison_unit::{
    generate_unid, ComparisonCorrelationStatus, ComparisonUnit, ComparisonUnitAtom,
    ComparisonUnitGroup, ComparisonUnitGroupContents, ComparisonUnitGroupType, ComparisonUnitWord,
    ContentElement,
};
use super::settings::WmlComparerSettings;
#[cfg(feature = "trace")]
use super::settings::{LcsTraceOperation, LcsTraceOutput};
#[cfg(feature = "trace")]
use std::collections::HashMap;

use std::sync::atomic::{AtomicUsize, Ordering};

// Global profiling counters
static LCS_CALLS: AtomicUsize = AtomicUsize::new(0);
static LCS_ITERATIONS: AtomicUsize = AtomicUsize::new(0);
static IDENTICAL_HITS: AtomicUsize = AtomicUsize::new(0);
static UNRELATED_HITS: AtomicUsize = AtomicUsize::new(0);
static CORR_HASH_HITS: AtomicUsize = AtomicUsize::new(0);
static BEGIN_END_HITS: AtomicUsize = AtomicUsize::new(0);
static LCS_ALGO_HITS: AtomicUsize = AtomicUsize::new(0);

/// Reset all profiling counters
pub fn reset_lcs_counters() {
    LCS_CALLS.store(0, Ordering::Relaxed);
    LCS_ITERATIONS.store(0, Ordering::Relaxed);
    IDENTICAL_HITS.store(0, Ordering::Relaxed);
    UNRELATED_HITS.store(0, Ordering::Relaxed);
    CORR_HASH_HITS.store(0, Ordering::Relaxed);
    BEGIN_END_HITS.store(0, Ordering::Relaxed);
    LCS_ALGO_HITS.store(0, Ordering::Relaxed);
}

/// Get profiling counters as a formatted string
pub fn get_lcs_counters() -> String {
    format!(
        "LCS calls: {}, iterations: {}, identical: {}, unrelated: {}, corr_hash: {}, begin_end: {}, lcs_algo: {}",
        LCS_CALLS.load(Ordering::Relaxed),
        LCS_ITERATIONS.load(Ordering::Relaxed),
        IDENTICAL_HITS.load(Ordering::Relaxed),
        UNRELATED_HITS.load(Ordering::Relaxed),
        CORR_HASH_HITS.load(Ordering::Relaxed),
        BEGIN_END_HITS.load(Ordering::Relaxed),
        LCS_ALGO_HITS.load(Ordering::Relaxed),
    )
}

/// Correlation status for sequences
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CorrelationStatus {
    #[default]
    Unknown,
    Equal,
    Deleted,
    Inserted,
}

/// A correlated sequence of comparison units
///
/// Matches C# CorrelatedSequence (WmlComparer.cs:5656)
#[derive(Debug, Clone)]
pub struct CorrelatedSequence {
    /// Correlation status
    pub status: CorrelationStatus,
    /// Comparison units from document 1 (original)
    pub units1: Option<Vec<ComparisonUnit>>,
    /// Comparison units from document 2 (modified)
    pub units2: Option<Vec<ComparisonUnit>>,
}

impl CorrelatedSequence {
    /// Create a new Unknown sequence
    pub fn unknown(units1: Vec<ComparisonUnit>, units2: Vec<ComparisonUnit>) -> Self {
        Self {
            status: CorrelationStatus::Unknown,
            units1: Some(units1),
            units2: Some(units2),
        }
    }

    /// Create a new Equal sequence
    pub fn equal(units1: Vec<ComparisonUnit>, units2: Vec<ComparisonUnit>) -> Self {
        Self {
            status: CorrelationStatus::Equal,
            units1: Some(units1),
            units2: Some(units2),
        }
    }

    /// Create a new Deleted sequence
    pub fn deleted(units1: Vec<ComparisonUnit>) -> Self {
        Self {
            status: CorrelationStatus::Deleted,
            units1: Some(units1),
            units2: None,
        }
    }

    /// Create a new Inserted sequence
    pub fn inserted(units2: Vec<ComparisonUnit>) -> Self {
        Self {
            status: CorrelationStatus::Inserted,
            units1: None,
            units2: Some(units2),
        }
    }

    /// Get length of units1
    pub fn len1(&self) -> usize {
        self.units1.as_ref().map(|u| u.len()).unwrap_or(0)
    }

    /// Get length of units2
    pub fn len2(&self) -> usize {
        self.units2.as_ref().map(|u| u.len()).unwrap_or(0)
    }
}

/// Main LCS algorithm entry point
///
/// Matches C# Lcs method (WmlComparer.cs:5779-5846)
///
/// Iteratively processes Unknown sequences until all are resolved.
pub fn lcs(
    units1: Vec<ComparisonUnit>,
    units2: Vec<ComparisonUnit>,
    settings: &WmlComparerSettings,
) -> Vec<CorrelatedSequence> {
    LCS_CALLS.fetch_add(1, Ordering::Relaxed);

    // Check for completely identical sources first (optimization)
    if let Some(result) = detect_identical_sources(&units1, &units2) {
        IDENTICAL_HITS.fetch_add(1, Ordering::Relaxed);
        return result;
    }

    // Check for completely unrelated sources (optimization)
    if let Some(result) = detect_unrelated_sources(&units1, &units2) {
        UNRELATED_HITS.fetch_add(1, Ordering::Relaxed);
        return result;
    }

    // Initialize with one Unknown sequence containing entire arrays
    let initial = CorrelatedSequence::unknown(units1, units2);
    let mut cs_list = vec![initial];

    loop {
        LCS_ITERATIONS.fetch_add(1, Ordering::Relaxed);

        // Find first Unknown sequence
        let unknown_idx = cs_list
            .iter()
            .position(|cs| cs.status == CorrelationStatus::Unknown);

        let Some(idx) = unknown_idx else {
            // All sequences resolved
            return cs_list;
        };

        // Extract the unknown sequence for processing
        let unknown = cs_list.remove(idx);

        // Set unids for matching single groups
        let unknown = set_after_unids(unknown);

        // Try ProcessCorrelatedHashes first (fastest)
        let new_sequences = if let (Some(units1), Some(units2)) =
            (unknown.units1.as_ref(), unknown.units2.as_ref())
        {
            if should_flatten_tabular_units(units1, units2) {
                let flat1 = flatten_units_for_unknown(units1);
                let flat2 = flatten_units_for_unknown(units2);
                vec![CorrelatedSequence::unknown(flat1, flat2)]
            } else if let Some(seqs) = split_on_tabular_span(units1, units2) {
                seqs
            } else if let Some(seqs) = process_correlated_hashes(&unknown, settings) {
                CORR_HASH_HITS.fetch_add(1, Ordering::Relaxed);
                seqs
            } else if !is_tabular_word_sequence(units1, units2) {
                if let Some(seqs) = find_common_at_beginning_and_end(&unknown, settings) {
                    BEGIN_END_HITS.fetch_add(1, Ordering::Relaxed);
                    seqs
                } else {
                    LCS_ALGO_HITS.fetch_add(1, Ordering::Relaxed);
                    do_lcs_algorithm(&unknown, settings)
                }
            } else {
                LCS_ALGO_HITS.fetch_add(1, Ordering::Relaxed);
                do_lcs_algorithm(&unknown, settings)
            }
        } else if let Some(seqs) = process_correlated_hashes(&unknown, settings) {
            CORR_HASH_HITS.fetch_add(1, Ordering::Relaxed);
            seqs
        } else if let Some(seqs) = find_common_at_beginning_and_end(&unknown, settings) {
            BEGIN_END_HITS.fetch_add(1, Ordering::Relaxed);
            seqs
        } else {
            LCS_ALGO_HITS.fetch_add(1, Ordering::Relaxed);
            do_lcs_algorithm(&unknown, settings)
        };

        // Insert new sequences at the position of the old unknown
        // (Reverse to maintain order when inserting at same position)
        for seq in new_sequences.into_iter().rev() {
            cs_list.insert(idx, seq);
        }
    }
}

// ============================================================================
// Trace functions - only compiled when "trace" feature is enabled
// ============================================================================

/// Extract text content from a ComparisonUnit for tracing/filtering
#[cfg(feature = "trace")]
fn extract_unit_text(unit: &ComparisonUnit) -> String {
    match unit {
        ComparisonUnit::Word(word) => word
            .atoms
            .iter()
            .map(|atom| atom.content_element.text_value())
            .collect::<Vec<_>>()
            .join(""),
        ComparisonUnit::Group(group) => extract_group_text(group),
    }
}

/// Extract text from a group recursively
#[cfg(feature = "trace")]
fn extract_group_text(group: &ComparisonUnitGroup) -> String {
    match &group.contents {
        ComparisonUnitGroupContents::Words(words) => words
            .iter()
            .flat_map(|word| word.atoms.iter())
            .map(|atom| atom.content_element.text_value())
            .collect::<Vec<_>>()
            .join(""),
        ComparisonUnitGroupContents::Groups(groups) => groups
            .iter()
            .map(extract_group_text)
            .collect::<Vec<_>>()
            .join(""),
    }
}

/// Extract text from a list of units
#[cfg(feature = "trace")]
#[allow(dead_code)]
fn extract_units_text(units: &[ComparisonUnit]) -> String {
    units
        .iter()
        .map(extract_unit_text)
        .collect::<Vec<_>>()
        .join("")
}

/// Convert a ComparisonUnit to a token string for tracing
#[cfg(feature = "trace")]
fn unit_to_token(unit: &ComparisonUnit) -> String {
    extract_unit_text(unit)
}

/// Generate trace output from correlated sequences
///
/// This captures the result of the LCS algorithm in a format suitable for debugging.
#[cfg(feature = "trace")]
pub fn generate_lcs_trace(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
    correlated: &[CorrelatedSequence],
    matched_text: String,
) -> LcsTraceOutput {
    // Extract tokens from input units
    let tokens1: Vec<String> = units1.iter().map(unit_to_token).collect();
    let tokens2: Vec<String> = units2.iter().map(unit_to_token).collect();

    // Generate raw operations from correlated sequences
    let mut raw_operations = Vec::new();
    let mut pos1 = 0usize;
    let mut pos2 = 0usize;

    for seq in correlated {
        match seq.status {
            CorrelationStatus::Equal => {
                if let Some(ref units) = seq.units1 {
                    for unit in units {
                        raw_operations.push(LcsTraceOperation {
                            op: "equal".to_string(),
                            tokens: vec![unit_to_token(unit)],
                            pos1: Some(pos1),
                            pos2: Some(pos2),
                        });
                        pos1 += 1;
                        pos2 += 1;
                    }
                }
            }
            CorrelationStatus::Deleted => {
                if let Some(ref units) = seq.units1 {
                    for unit in units {
                        raw_operations.push(LcsTraceOperation {
                            op: "delete".to_string(),
                            tokens: vec![unit_to_token(unit)],
                            pos1: Some(pos1),
                            pos2: None,
                        });
                        pos1 += 1;
                    }
                }
            }
            CorrelationStatus::Inserted => {
                if let Some(ref units) = seq.units2 {
                    for unit in units {
                        raw_operations.push(LcsTraceOperation {
                            op: "insert".to_string(),
                            tokens: vec![unit_to_token(unit)],
                            pos1: None,
                            pos2: Some(pos2),
                        });
                        pos2 += 1;
                    }
                }
            }
            CorrelationStatus::Unknown => {
                // Should not happen in final output
            }
        }
    }

    // Coalesce consecutive operations of the same type
    let coalesced_operations = coalesce_trace_operations(&raw_operations);

    // Calculate LCS length (count of equal operations)
    let lcs_length = raw_operations.iter().filter(|op| op.op == "equal").count();

    LcsTraceOutput {
        matched_text,
        tokens1,
        tokens2,
        raw_operations,
        coalesced_operations,
        lcs_length,
    }
}

/// Coalesce consecutive trace operations of the same type
#[cfg(feature = "trace")]
fn coalesce_trace_operations(ops: &[LcsTraceOperation]) -> Vec<LcsTraceOperation> {
    if ops.is_empty() {
        return Vec::new();
    }

    let mut result = Vec::new();
    let mut current = LcsTraceOperation {
        op: ops[0].op.clone(),
        tokens: ops[0].tokens.clone(),
        pos1: ops[0].pos1,
        pos2: ops[0].pos2,
    };

    for op in &ops[1..] {
        if op.op == current.op {
            // Same operation type - extend tokens
            current.tokens.extend(op.tokens.iter().cloned());
        } else {
            // Different operation - push current and start new
            result.push(current);
            current = LcsTraceOperation {
                op: op.op.clone(),
                tokens: op.tokens.clone(),
                pos1: op.pos1,
                pos2: op.pos2,
            };
        }
    }
    result.push(current);

    result
}

/// Information about a matched paragraph for tracing
#[cfg(feature = "trace")]
#[derive(Debug, Clone)]
pub struct MatchedParagraphInfo {
    /// The text content of the matched paragraph
    pub text: String,
    /// Index of the matched unit in the units array
    pub index: usize,
}

/// Check if any unit in the list contains text matching the filter
/// Returns the matched paragraph info if found
#[cfg(feature = "trace")]
pub fn units_match_filter(
    units: &[ComparisonUnit],
    settings: &WmlComparerSettings,
) -> Option<MatchedParagraphInfo> {
    if !settings.is_tracing_enabled() {
        return None;
    }

    // Check each unit (especially groups which represent paragraphs)
    for (index, unit) in units.iter().enumerate() {
        if let ComparisonUnit::Group(group) = unit {
            if group.group_type == ComparisonUnitGroupType::Paragraph {
                let text = extract_group_text(group);
                if settings.should_trace_paragraph(&text) {
                    return Some(MatchedParagraphInfo { text, index });
                }
            }
        }
    }

    None
}

/// Find the best matching paragraph in the other document
/// Uses text similarity to find the corresponding paragraph
#[cfg(feature = "trace")]
#[allow(dead_code)]
fn find_corresponding_paragraph(
    target_text: &str,
    units: &[ComparisonUnit],
    settings: &WmlComparerSettings,
) -> Option<MatchedParagraphInfo> {
    // First, try to find an exact or near-exact match using the same filter
    for (index, unit) in units.iter().enumerate() {
        if let ComparisonUnit::Group(group) = unit {
            if group.group_type == ComparisonUnitGroupType::Paragraph {
                let text = extract_group_text(group);
                // Check if this paragraph matches the filter (same section/prefix)
                if settings.should_trace_paragraph(&text) {
                    return Some(MatchedParagraphInfo { text, index });
                }
            }
        }
    }

    // Fallback: find paragraph with most text overlap
    let target_words: std::collections::HashSet<&str> = target_text.split_whitespace().collect();
    let mut best_match: Option<(usize, String, usize)> = None;

    for (index, unit) in units.iter().enumerate() {
        if let ComparisonUnit::Group(group) = unit {
            if group.group_type == ComparisonUnitGroupType::Paragraph {
                let text = extract_group_text(group);
                let text_words: std::collections::HashSet<&str> = text.split_whitespace().collect();
                let overlap = target_words.intersection(&text_words).count();

                if overlap > 0 {
                    if best_match.is_none() || overlap > best_match.as_ref().unwrap().2 {
                        best_match = Some((index, text, overlap));
                    }
                }
            }
        }
    }

    best_match.map(|(index, text, _)| MatchedParagraphInfo { text, index })
}

/// Extract words from a paragraph unit for focused comparison
#[cfg(feature = "trace")]
fn extract_words_from_unit(unit: &ComparisonUnit) -> Vec<ComparisonUnit> {
    match unit {
        ComparisonUnit::Group(group) => {
            match &group.contents {
                ComparisonUnitGroupContents::Words(words) => words
                    .iter()
                    .map(|w| ComparisonUnit::Word(w.clone()))
                    .collect(),
                ComparisonUnitGroupContents::Groups(groups) => {
                    // Recursively extract from nested groups
                    groups
                        .iter()
                        .flat_map(|g| extract_words_from_unit(&ComparisonUnit::Group(g.clone())))
                        .collect()
                }
            }
        }
        ComparisonUnit::Word(word) => vec![ComparisonUnit::Word(word.clone())],
    }
}

/// Generate a focused LCS trace for just the matched paragraph
/// This runs a separate LCS on just the paragraph's words for detailed debugging
#[cfg(feature = "trace")]
pub fn generate_focused_trace(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
    matched1: &MatchedParagraphInfo,
    matched2: Option<&MatchedParagraphInfo>,
    settings: &WmlComparerSettings,
) -> LcsTraceOutput {
    // Extract words from the matched paragraphs
    let para1 = &units1[matched1.index];
    let words1 = extract_words_from_unit(para1);

    let (words2, matched_text2) = if let Some(m2) = matched2 {
        let para2 = &units2[m2.index];
        (extract_words_from_unit(para2), m2.text.clone())
    } else {
        // No corresponding paragraph found - compare against empty
        (Vec::new(), String::new())
    };

    // Run LCS on just the paragraph words
    let correlated = lcs(words1.clone(), words2.clone(), settings);

    // Generate tokens from words
    let tokens1: Vec<String> = words1.iter().map(unit_to_token).collect();
    let tokens2: Vec<String> = words2.iter().map(unit_to_token).collect();

    // Generate operations from correlated sequences
    let mut raw_operations = Vec::new();
    let mut pos1 = 0usize;
    let mut pos2 = 0usize;

    for seq in &correlated {
        match seq.status {
            CorrelationStatus::Equal => {
                if let Some(ref units) = seq.units1 {
                    for unit in units {
                        raw_operations.push(LcsTraceOperation {
                            op: "equal".to_string(),
                            tokens: vec![unit_to_token(unit)],
                            pos1: Some(pos1),
                            pos2: Some(pos2),
                        });
                        pos1 += 1;
                        pos2 += 1;
                    }
                }
            }
            CorrelationStatus::Deleted => {
                if let Some(ref units) = seq.units1 {
                    for unit in units {
                        raw_operations.push(LcsTraceOperation {
                            op: "delete".to_string(),
                            tokens: vec![unit_to_token(unit)],
                            pos1: Some(pos1),
                            pos2: None,
                        });
                        pos1 += 1;
                    }
                }
            }
            CorrelationStatus::Inserted => {
                if let Some(ref units) = seq.units2 {
                    for unit in units {
                        raw_operations.push(LcsTraceOperation {
                            op: "insert".to_string(),
                            tokens: vec![unit_to_token(unit)],
                            pos1: None,
                            pos2: Some(pos2),
                        });
                        pos2 += 1;
                    }
                }
            }
            CorrelationStatus::Unknown => {}
        }
    }

    let coalesced_operations = coalesce_trace_operations(&raw_operations);
    let lcs_length = raw_operations.iter().filter(|op| op.op == "equal").count();

    LcsTraceOutput {
        matched_text: format!(
            "Para {}: {} | Para {}: {}",
            matched1.index,
            truncate_text(&matched1.text, 50),
            matched2.map(|m| m.index).unwrap_or(0),
            truncate_text(&matched_text2, 50)
        ),
        tokens1,
        tokens2,
        raw_operations,
        coalesced_operations,
        lcs_length,
    }
}

/// Truncate text for display
#[cfg(feature = "trace")]
fn truncate_text(text: &str, max_len: usize) -> String {
    if text.len() <= max_len {
        text.to_string()
    } else {
        format!("{}...", &text[..max_len])
    }
}

/// Set After UNIDs for matching single groups
///
/// Matches C# SetAfterUnids (WmlComparer.cs:5848-5936)
///
/// When both sides have a single group of the same type, propagate the UNIDs
/// from document 1 to document 2's ancestor elements. This enables proper
/// tree reconstruction later.
fn set_after_unids(mut unknown: CorrelatedSequence) -> CorrelatedSequence {
    let units1 = match &unknown.units1 {
        Some(u) if u.len() == 1 => u,
        _ => return unknown,
    };
    let units2 = match &mut unknown.units2 {
        Some(u) if u.len() == 1 => u,
        _ => return unknown,
    };

    // Both must be groups of the same type
    let group1 = match &units1[0] {
        ComparisonUnit::Group(g) => g,
        _ => return unknown,
    };
    let group2 = match &mut units2[0] {
        ComparisonUnit::Group(g) => g,
        _ => return unknown,
    };

    if group1.group_type != group2.group_type {
        return unknown;
    }

    // Get descendant atoms from both groups
    let atoms1 = group1.descendant_atoms();
    let atoms2 = group2.descendant_atoms();

    if atoms1.is_empty() || atoms2.is_empty() {
        return unknown;
    }

    // Determine which ancestor elements to include based on group type
    let take_through_name = match group1.group_type {
        ComparisonUnitGroupType::Paragraph => "p",
        ComparisonUnitGroupType::Table => "tbl",
        ComparisonUnitGroupType::Row => "tr",
        ComparisonUnitGroupType::Cell => "tc",
        ComparisonUnitGroupType::Textbox => "txbxContent",
    };

    // Collect relevant ancestors from first atom in group1
    let first_atom1 = atoms1[0];
    let mut relevant_ancestors = Vec::new();
    for ancestor in &first_atom1.ancestor_elements {
        relevant_ancestors.push(ancestor.unid.clone());
        if ancestor.local_name == take_through_name {
            break;
        }
    }

    // Generate missing UNIDs if needed (SDK 3.x compatibility)
    for unid in &mut relevant_ancestors {
        if unid.is_empty() {
            *unid = generate_unid();
        }
    }

    unknown
}

/// Process correlated hashes for quick matching
///
/// Matches C# ProcessCorrelatedHashes (WmlComparer.cs:5938-6146)
///
/// Uses pre-computed CorrelatedSHA1Hash values to find matching groups.
/// This is an optimization for paragraph/table/row-level matching.
fn process_correlated_hashes(
    unknown: &CorrelatedSequence,
    _settings: &WmlComparerSettings,
) -> Option<Vec<CorrelatedSequence>> {
    let units1 = unknown.units1.as_ref()?;
    let units2 = unknown.units2.as_ref()?;

    // Never attempt this optimization if there are less than 3 groups
    let max_depth = units1.len().min(units2.len());
    if max_depth < 3 {
        return None;
    }

    // Check that first elements are groups of appropriate types
    let first1 = units1.first()?.as_group()?;
    let first2 = units2.first()?.as_group()?;

    let valid_types = matches!(
        first1.group_type,
        ComparisonUnitGroupType::Paragraph
            | ComparisonUnitGroupType::Table
            | ComparisonUnitGroupType::Row
            | ComparisonUnitGroupType::Textbox
    ) && matches!(
        first2.group_type,
        ComparisonUnitGroupType::Paragraph
            | ComparisonUnitGroupType::Table
            | ComparisonUnitGroupType::Row
            | ComparisonUnitGroupType::Textbox
    );

    if !valid_types {
        return None;
    }

    // Find longest common sequence using CorrelatedSHA1Hash
    let mut best_length = 0usize;
    let mut best_atom_count = 0usize;
    let mut best_i1 = 0usize;
    let mut best_i2 = 0usize;

    // Optimization: Index units2 by correlated hash for O(1) lookup
    // Map hash -> list of indices in units2
    let mut units2_index: HashMap<&str, Vec<usize>> = HashMap::new();
    for (i2, unit) in units2.iter().enumerate() {
        if let ComparisonUnit::Group(g) = unit {
            if let Some(hash) = &g.correlated_sha1_hash {
                units2_index.entry(hash.as_str()).or_default().push(i2);
            }
        }
    }

    for i1 in 0..units1.len() {
        // Skip if we can't possibly beat best_length
        if units1.len() - i1 <= best_length {
            break;
        }

        let group1 = match &units1[i1] {
            ComparisonUnit::Group(g) => g,
            _ => continue,
        };

        let Some(hash1) = &group1.correlated_sha1_hash else {
            continue;
        };

        // Only check indices that have matching hash
        if let Some(candidates) = units2_index.get(hash1.as_str()) {
            for &i2 in candidates {
                // Optimization: Skip if we can't beat best_length
                if units2.len() - i2 <= best_length {
                    continue;
                }

                let mut seq_length = 0usize;
                let mut seq_atom_count = 0usize;
                let mut cur_i1 = i1;
                let mut cur_i2 = i2;

                while cur_i1 < units1.len() && cur_i2 < units2.len() {
                    let group1 = match &units1[cur_i1] {
                        ComparisonUnit::Group(g) => g,
                        _ => break,
                    };
                    let group2 = match &units2[cur_i2] {
                        ComparisonUnit::Group(g) => g,
                        _ => break,
                    };

                    // Match if same type and same correlated hash
                    // Note: We already know hash matches for first item from index,
                    // but we need to check type and subsequent items.
                    let matches = group1.group_type == group2.group_type
                        && group1.correlated_sha1_hash.is_some()
                        && group1.correlated_sha1_hash == group2.correlated_sha1_hash;

                    if matches {
                        seq_atom_count += group1.descendant_atom_count();
                        cur_i1 += 1;
                        cur_i2 += 1;
                        seq_length += 1;
                    } else {
                        break;
                    }
                }

                if seq_atom_count > best_atom_count {
                    best_length = seq_length;
                    best_atom_count = seq_atom_count;
                    best_i1 = i1;
                    best_i2 = i2;
                }
            }
        }
    }

    // Apply thresholds based on sequence length and atom count
    let do_correlation = if best_length == 1 {
        // Single group needs 16+ atoms on each side
        let atoms1 = units1[best_i1].descendant_content_atoms_count();
        let atoms2 = units2[best_i2].descendant_content_atoms_count();
        atoms1 > 16 && atoms2 > 16
    } else if best_length > 1 && best_length <= 3 {
        // 2-3 groups need 32+ atoms total on each side
        let atoms1: usize = units1[best_i1..best_i1 + best_length]
            .iter()
            .map(|u| u.descendant_content_atoms_count())
            .sum();
        let atoms2: usize = units2[best_i2..best_i2 + best_length]
            .iter()
            .map(|u| u.descendant_content_atoms_count())
            .sum();
        atoms1 > 32 && atoms2 > 32
    } else {
        // 4+ groups always correlate
        best_length > 3
    };

    if !do_correlation {
        return None;
    }

    // Build result sequences
    let mut result = Vec::new();

    // Handle prefix (before match)
    if best_i1 > 0 && best_i2 == 0 {
        result.push(CorrelatedSequence::deleted(units1[..best_i1].to_vec()));
    } else if best_i1 == 0 && best_i2 > 0 {
        result.push(CorrelatedSequence::inserted(units2[..best_i2].to_vec()));
    } else if best_i1 > 0 && best_i2 > 0 {
        result.push(CorrelatedSequence::unknown(
            units1[..best_i1].to_vec(),
            units2[..best_i2].to_vec(),
        ));
    }

    // Add matched groups - if sha1_hash matches (content identical), use Equal
    // Otherwise use Unknown for further processing of internal differences
    for i in 0..best_length {
        let u1 = &units1[best_i1 + i];
        let u2 = &units2[best_i2 + i];

        // If both content hashes match, the content is identical - no need to recurse
        let content_identical = match (u1, u2) {
            (ComparisonUnit::Group(g1), ComparisonUnit::Group(g2)) => g1.sha1_hash == g2.sha1_hash,
            _ => false,
        };

        if content_identical {
            result.push(CorrelatedSequence::equal(
                vec![u1.clone()],
                vec![u2.clone()],
            ));
        } else {
            result.push(CorrelatedSequence::unknown(
                vec![u1.clone()],
                vec![u2.clone()],
            ));
        }
    }

    // Handle suffix (after match)
    let end_i1 = best_i1 + best_length;
    let end_i2 = best_i2 + best_length;

    if end_i1 < units1.len() && end_i2 == units2.len() {
        result.push(CorrelatedSequence::deleted(units1[end_i1..].to_vec()));
    } else if end_i1 == units1.len() && end_i2 < units2.len() {
        result.push(CorrelatedSequence::inserted(units2[end_i2..].to_vec()));
    } else if end_i1 < units1.len() && end_i2 < units2.len() {
        result.push(CorrelatedSequence::unknown(
            units1[end_i1..].to_vec(),
            units2[end_i2..].to_vec(),
        ));
    }

    Some(result)
}

/// Split a comparison unit slice at the first paragraph mark
///
/// Matches C# SplitAtParagraphMark (WmlComparer.cs:4974-4995)
///
/// Finds the first comparison unit that starts with a paragraph mark (w:pPr)
/// and splits the slice into [0..i] and [i..end].
/// If no paragraph mark is found, returns a single-element vec containing the original slice.
///
/// # Returns
/// - `Vec<Vec<ComparisonUnit>>` with 1 element if no paragraph mark found
/// - `Vec<Vec<ComparisonUnit>>` with 2 elements if paragraph mark found at index i:
///   - First: units before the paragraph mark (0..i)
///   - Second: units from paragraph mark onwards (i..end)
fn split_at_paragraph_mark(units: &[ComparisonUnit]) -> Vec<Vec<ComparisonUnit>> {
    // C# WmlComparer.cs:4977-4982
    // for (i = 0; i < cua.Length; i++)
    // {
    //     var atom = cua[i].DescendantContentAtoms().FirstOrDefault();
    //     if (atom != null && atom.ContentElement.Name == W.pPr)
    //         break;
    // }
    for i in 0..units.len() {
        // Get the first descendant atom from this comparison unit
        let first_atom = units[i].descendant_atoms().first().cloned();
        if let Some(atom) = first_atom {
            if matches!(
                atom.content_element,
                ContentElement::ParagraphProperties { .. }
            ) {
                // C# WmlComparer.cs:4990-4994: Split at this position
                // Note: C# uses Take(i) then Skip(i), so first part is [0..i), second is [i..end]
                return vec![units[..i].to_vec(), units[i..].to_vec()];
            }
        }
    }

    // C# WmlComparer.cs:4983-4988: No paragraph mark found, return single element
    vec![units.to_vec()]
}

const TABULAR_TAB_THRESHOLD: usize = 3;
const TABULAR_WORD_TAB_RATIO_NUM: usize = 3; // 1/3 of tokens or more are tabs
const TABULAR_LCS_DP_THRESHOLD: usize = 200;

fn is_tab_atom(atom: &ComparisonUnitAtom) -> bool {
    matches!(
        atom.content_element,
        ContentElement::Tab | ContentElement::PositionalTab { .. }
    )
}

fn group_tab_count(group: &ComparisonUnitGroup) -> usize {
    group
        .descendant_atoms()
        .iter()
        .filter(|atom| is_tab_atom(atom))
        .count()
}

fn is_tab_only_unit(unit: &ComparisonUnit) -> bool {
    let Some(word) = unit.as_word() else {
        return false;
    };

    !word.atoms.is_empty()
        && word.atoms.iter().all(|atom| {
            matches!(
                atom.content_element,
                ContentElement::Tab | ContentElement::PositionalTab { .. }
            )
        })
}

fn is_paragraph_mark_unit(unit: &ComparisonUnit) -> bool {
    unit.as_word().is_some_and(|word| word.is_paragraph_mark())
}

fn group_has_text_content(group: &ComparisonUnitGroup) -> bool {
    group.descendant_atoms().iter().any(|atom| {
        matches!(&atom.content_element, ContentElement::Text(c) if !c.is_whitespace())
    })
}

fn group_is_tabular(group: &ComparisonUnitGroup) -> bool {
    group_tab_count(group) >= TABULAR_TAB_THRESHOLD
}

fn group_is_empty(group: &ComparisonUnitGroup) -> bool {
    !group_has_text_content(group) && group_tab_count(group) == 0
}

fn only_paragraph_groups(units: &[ComparisonUnit]) -> bool {
    units.iter().all(|unit| {
        matches!(
            unit,
            ComparisonUnit::Group(g) if g.group_type == ComparisonUnitGroupType::Paragraph
        )
    })
}

fn should_flatten_tabular_units(units1: &[ComparisonUnit], units2: &[ComparisonUnit]) -> bool {
    if !only_paragraph_groups(units1) || !only_paragraph_groups(units2) {
        return false;
    }

    let mut left_tabular = 0usize;
    let mut right_tabular = 0usize;
    let mut left_non_tabular = 0usize;
    let mut right_non_tabular = 0usize;

    for unit in units1 {
        if let Some(group) = unit.as_group() {
            if group_is_tabular(group) {
                left_tabular += 1;
            } else if !group_is_empty(group) {
                left_non_tabular += 1;
            }
        }
    }

    for unit in units2 {
        if let Some(group) = unit.as_group() {
            if group_is_tabular(group) {
                right_tabular += 1;
            } else if !group_is_empty(group) {
                right_non_tabular += 1;
            }
        }
    }

    left_tabular > 0
        && right_tabular > 0
        && left_non_tabular == 0
        && right_non_tabular == 0
}

#[derive(Clone, Copy)]
struct TabularSpan {
    start: usize,
    end: usize,
    tabular_count: usize,
}

fn find_tabular_span(units: &[ComparisonUnit]) -> Option<TabularSpan> {
    let mut best: Option<TabularSpan> = None;
    let mut current_start: Option<usize> = None;
    let mut current_tabular = 0usize;

    for (idx, unit) in units.iter().enumerate() {
        let Some(group) = unit.as_group() else {
            if let Some(start) = current_start.take() {
                if current_tabular > 0 {
                    let span = TabularSpan {
                        start,
                        end: idx,
                        tabular_count: current_tabular,
                    };
                    if best
                        .map(|b| (span.tabular_count, span.end - span.start)
                            > (b.tabular_count, b.end - b.start))
                        .unwrap_or(true)
                    {
                        best = Some(span);
                    }
                }
                current_tabular = 0;
            }
            continue;
        };
        if group.group_type != ComparisonUnitGroupType::Paragraph {
            if let Some(start) = current_start.take() {
                if current_tabular > 0 {
                    let span = TabularSpan {
                        start,
                        end: idx,
                        tabular_count: current_tabular,
                    };
                    if best
                        .map(|b| (span.tabular_count, span.end - span.start)
                            > (b.tabular_count, b.end - b.start))
                        .unwrap_or(true)
                    {
                        best = Some(span);
                    }
                }
                current_tabular = 0;
            }
            continue;
        }
        let is_tabular = group_is_tabular(group);
        let is_empty = group_is_empty(group);

        if is_tabular || is_empty {
            if current_start.is_none() {
                current_start = Some(idx);
                current_tabular = 0;
            }
            if is_tabular {
                current_tabular += 1;
            }
        } else if let Some(start) = current_start.take() {
            if current_tabular > 0 {
                let span = TabularSpan {
                    start,
                    end: idx,
                    tabular_count: current_tabular,
                };
                if best
                    .map(|b| (span.tabular_count, span.end - span.start)
                        > (b.tabular_count, b.end - b.start))
                    .unwrap_or(true)
                {
                    best = Some(span);
                }
            }
            current_tabular = 0;
        }
    }

    if let Some(start) = current_start {
        if current_tabular > 0 {
            let span = TabularSpan {
                start,
                end: units.len(),
                tabular_count: current_tabular,
            };
            if best
                .map(|b| (span.tabular_count, span.end - span.start)
                    > (b.tabular_count, b.end - b.start))
                .unwrap_or(true)
            {
                best = Some(span);
            }
        }
    }

    best
}

fn split_on_tabular_span(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Option<Vec<CorrelatedSequence>> {
    let span1 = find_tabular_span(units1)?;
    let span2 = find_tabular_span(units2)?;

    if span1.tabular_count < 2 || span2.tabular_count < 2 {
        return None;
    }

    let prefix_delta = if span1.start > span2.start {
        span1.start - span2.start
    } else {
        span2.start - span1.start
    };

    if prefix_delta > 1 {
        return None;
    }

    let mut sequences: Vec<CorrelatedSequence> = Vec::new();

    if span1.start > 0 || span2.start > 0 {
        sequences.push(CorrelatedSequence::unknown(
            units1[..span1.start].to_vec(),
            units2[..span2.start].to_vec(),
        ));
    }

    sequences.push(CorrelatedSequence::unknown(
        flatten_units_for_unknown(&units1[span1.start..span1.end]),
        flatten_units_for_unknown(&units2[span2.start..span2.end]),
    ));

    if span1.end < units1.len() || span2.end < units2.len() {
        sequences.push(CorrelatedSequence::unknown(
            units1[span1.end..].to_vec(),
            units2[span2.end..].to_vec(),
        ));
    }

    Some(sequences)
}

fn is_tabular_word_sequence(units1: &[ComparisonUnit], units2: &[ComparisonUnit]) -> bool {
    let all_words1 = units1.iter().all(|u| u.as_word().is_some());
    let all_words2 = units2.iter().all(|u| u.as_word().is_some());

    if !all_words1 || !all_words2 {
        return false;
    }

    let tab_only1 = units1.iter().filter(|u| is_tab_only_unit(u)).count();
    let tab_only2 = units2.iter().filter(|u| is_tab_only_unit(u)).count();

    let ratio1 = tab_only1 * TABULAR_WORD_TAB_RATIO_NUM >= units1.len();
    let ratio2 = tab_only2 * TABULAR_WORD_TAB_RATIO_NUM >= units2.len();

    ratio1 || ratio2
}

fn matchable_hash<'a>(unit: &'a ComparisonUnit, skip_tab_only: bool) -> Option<&'a str> {
    if skip_tab_only && (is_tab_only_unit(unit) || is_paragraph_mark_unit(unit)) {
        None
    } else {
        Some(unit.hash())
    }
}

/// Find common elements at beginning and end
///
/// Matches C# FindCommonAtBeginningAndEnd (WmlComparer.cs:4489-4972)
///
/// Quick check for matching prefix/suffix using SHA1Hash comparison.
fn find_common_at_beginning_and_end(
    unknown: &CorrelatedSequence,
    settings: &WmlComparerSettings,
) -> Option<Vec<CorrelatedSequence>> {
    let units1 = unknown.units1.as_ref()?;
    let units2 = unknown.units2.as_ref()?;

    let length_to_compare = units1.len().min(units2.len());
    if length_to_compare == 0 {
        return None;
    }

    // Count common at beginning
    let count_common_at_beginning = units1
        .iter()
        .zip(units2.iter())
        .take(length_to_compare)
        .take_while(|(u1, u2)| u1.hash() == u2.hash())
        .count();

    // Apply detail threshold
    let count_common_at_beginning = if count_common_at_beginning > 0 {
        let ratio = count_common_at_beginning as f64 / length_to_compare as f64;
        if ratio < settings.detail_threshold {
            0
        } else {
            count_common_at_beginning
        }
    } else {
        0
    };

    if count_common_at_beginning > 0 {
        let mut result = Vec::new();

        // Add Equal sequence for common prefix
        result.push(CorrelatedSequence::equal(
            units1[..count_common_at_beginning].to_vec(),
            units2[..count_common_at_beginning].to_vec(),
        ));

        // Handle remaining
        let remaining_left = units1.len() - count_common_at_beginning;
        let remaining_right = units2.len() - count_common_at_beginning;

        if remaining_left > 0 && remaining_right == 0 {
            // C# WmlComparer.cs:4536-4544
            result.push(CorrelatedSequence::deleted(
                units1[count_common_at_beginning..].to_vec(),
            ));
        } else if remaining_left == 0 && remaining_right > 0 {
            // C# WmlComparer.cs:4529-4534
            result.push(CorrelatedSequence::inserted(
                units2[count_common_at_beginning..].to_vec(),
            ));
        } else if remaining_left > 0 && remaining_right > 0 {
            // C# WmlComparer.cs:4546-4612: Paragraph-aware prefix splitting
            // Check if we're operating at word level and can do paragraph-aware splitting
            let first1 = units1[0].as_word();
            let first2 = units2[0].as_word();

            if first1.is_some() && first2.is_some() {
                // C# WmlComparer.cs:4551-4561: Comment from C# explains the logic:
                // if operating at the word level and
                //   if the last word on the left != pPr && last word on right != pPr
                //     then create an unknown for the rest of the paragraph, and create an unknown for the rest of the unknown
                //   if the last word on the left != pPr and last word on right == pPr
                //     then create deleted for the left, and create an unknown for the rest of the unknown
                //   if the last word on the left == pPr and last word on right != pPr
                //     then create inserted for the right, and create an unknown for the rest of the unknown
                //   if the last word on the left == pPr and last word on right == pPr
                //     then create an unknown for the rest of the unknown

                // C# WmlComparer.cs:4563-4571: Get remaining content after common prefix
                let remaining_in_left = units1[count_common_at_beginning..].to_vec();
                let remaining_in_right = units2[count_common_at_beginning..].to_vec();

                // C# WmlComparer.cs:4573-4574: Get last content atom from the last common element
                // Note: C# uses countCommonAtBeginning - 1 to get the last element of the common prefix
                let last_content_atom_left = units1[count_common_at_beginning - 1]
                    .descendant_atoms()
                    .first()
                    .cloned();
                let last_content_atom_right = units2[count_common_at_beginning - 1]
                    .descendant_atoms()
                    .first()
                    .cloned();

                // C# WmlComparer.cs:4576: Check if both last atoms are NOT paragraph properties
                let left_not_ppr = last_content_atom_left
                    .as_ref()
                    .map(|a| {
                        !matches!(
                            a.content_element,
                            ContentElement::ParagraphProperties { .. }
                        )
                    })
                    .unwrap_or(false);
                let right_not_ppr = last_content_atom_right
                    .as_ref()
                    .map(|a| {
                        !matches!(
                            a.content_element,
                            ContentElement::ParagraphProperties { .. }
                        )
                    })
                    .unwrap_or(false);

                if left_not_ppr && right_not_ppr {
                    // C# WmlComparer.cs:4578-4579: Split remaining content at paragraph marks
                    let split1 = split_at_paragraph_mark(&remaining_in_left);
                    let split2 = split_at_paragraph_mark(&remaining_in_right);

                    // C# WmlComparer.cs:4580-4588: Both have no split (no paragraph mark in either)
                    if split1.len() == 1 && split2.len() == 1 {
                        result.push(CorrelatedSequence::unknown(
                            split1.into_iter().next().unwrap(),
                            split2.into_iter().next().unwrap(),
                        ));
                        return Some(result);
                    }
                    // C# WmlComparer.cs:4589-4604: Both split at paragraph mark
                    else if split1.len() == 2 && split2.len() == 2 {
                        // First unknown: content before paragraph mark
                        let mut split1_iter = split1.into_iter();
                        let mut split2_iter = split2.into_iter();

                        result.push(CorrelatedSequence::unknown(
                            split1_iter.next().unwrap(),
                            split2_iter.next().unwrap(),
                        ));

                        // Second unknown: content from paragraph mark onwards
                        result.push(CorrelatedSequence::unknown(
                            split1_iter.next().unwrap(),
                            split2_iter.next().unwrap(),
                        ));

                        return Some(result);
                    }
                    // C# WmlComparer.cs:4605: Fall through to default case if split counts don't match
                }
            }

            // C# WmlComparer.cs:4608-4612: Default case - single unknown for all remaining
            result.push(CorrelatedSequence::unknown(
                units1[count_common_at_beginning..].to_vec(),
                units2[count_common_at_beginning..].to_vec(),
            ));
        }

        return Some(result);
    }

    // If no common at beginning, try common at end
    let mut count_common_at_end = units1
        .iter()
        .rev()
        .zip(units2.iter().rev())
        .take(length_to_compare)
        .take_while(|(u1, u2)| u1.hash() == u2.hash())
        .count();

    // Never start a common section with a paragraph mark (unless it's the only thing)
    while count_common_at_end > 1 {
        let first_common_idx1 = units1.len() - count_common_at_end;
        if let Some(word) = units1[first_common_idx1].as_word() {
            if word.is_paragraph_mark() {
                count_common_at_end -= 1;
                continue;
            }
        }
        break;
    }

    // Check if only paragraph mark (C# lines 4672-4726)
    let mut is_only_paragraph_mark = false;

    // C# lines 4673-4694: countCommonAtEnd == 1 case
    if count_common_at_end == 1 {
        let first_common_idx1 = units1.len() - count_common_at_end;
        if let Some(word) = units1[first_common_idx1].as_word() {
            if word.atoms.len() == 1 {
                if let Some(atom) = word.atoms.first() {
                    if matches!(
                        atom.content_element,
                        ContentElement::ParagraphProperties { .. }
                    ) {
                        is_only_paragraph_mark = true;
                    }
                }
            }
        }
    }

    // C# lines 4696-4726: countCommonAtEnd == 2 case
    if count_common_at_end == 2 {
        let first_common_idx1 = units1.len() - count_common_at_end;
        let second_common_idx1 = units1.len() - 1;

        if let (Some(first_word), Some(second_word)) = (
            units1[first_common_idx1].as_word(),
            units1[second_common_idx1].as_word(),
        ) {
            if first_word.atoms.len() == 1 && second_word.atoms.len() == 1 {
                if let (Some(_first_atom), Some(second_atom)) =
                    (first_word.atoms.first(), second_word.atoms.first())
                {
                    if matches!(
                        second_atom.content_element,
                        ContentElement::ParagraphProperties { .. }
                    ) {
                        is_only_paragraph_mark = true;
                    }
                }
            }
        }
    }

    // Apply detail threshold (unless it's just a paragraph mark)
    if !is_only_paragraph_mark && count_common_at_end > 0 {
        let ratio = count_common_at_end as f64 / length_to_compare as f64;
        if ratio < settings.detail_threshold {
            count_common_at_end = 0;
        }
    }

    // C# line 4734: If only paragraph mark, don't use it as common end
    if is_only_paragraph_mark {
        count_common_at_end = 0;
    }

    // C# line 4737-4738: If no common at end, return None
    if count_common_at_end == 0 {
        return None;
    }

    // C# lines 4740-4868: Handle "remaining in paragraph" logic
    // If countCommonAtEnd != 0, and if it contains a paragraph mark, then if there are
    // comparison units in the same paragraph before the common at end (in either version)
    // then we want to put all of those comparison units into a single unknown, where they
    // must be resolved against each other.

    let mut remaining_in_left_paragraph = 0usize;
    let mut remaining_in_right_paragraph = 0usize;

    // C# lines 4748-4753: Get common end sequence
    let common_end_start = units1.len() - count_common_at_end;
    let common_end_seq: Vec<_> = units1[common_end_start..].to_vec();

    // C# lines 4755-4795: Check if first of common end is a Word and contains paragraph marks
    if let Some(first_of_common) = common_end_seq.first() {
        if first_of_common.as_word().is_some() {
            // Check if any unit in common end seq has a paragraph mark (pPr)
            let has_paragraph_mark = common_end_seq.iter().any(|cu| {
                if let Some(word) = cu.as_word() {
                    if let Some(first_atom) = word.atoms.first() {
                        return matches!(
                            first_atom.content_element,
                            ContentElement::ParagraphProperties { .. }
                        );
                    }
                }
                false
            });

            if has_paragraph_mark {
                // C# lines 4767-4780: Calculate remainingInLeftParagraph
                // Count units before common end that are in the same paragraph (no pPr)
                remaining_in_left_paragraph = units1[..common_end_start]
                    .iter()
                    .rev()
                    .take_while(|cu| {
                        if let Some(word) = cu.as_word() {
                            if let Some(first_atom) = word.atoms.first() {
                                // Continue while NOT a paragraph mark
                                return !matches!(
                                    first_atom.content_element,
                                    ContentElement::ParagraphProperties { .. }
                                );
                            }
                            // No atoms means continue
                            return true;
                        }
                        // Not a word, stop
                        false
                    })
                    .count();

                // C# lines 4781-4794: Calculate remainingInRightParagraph
                let common_end_start2 = units2.len() - count_common_at_end;
                remaining_in_right_paragraph = units2[..common_end_start2]
                    .iter()
                    .rev()
                    .take_while(|cu| {
                        if let Some(word) = cu.as_word() {
                            if let Some(first_atom) = word.atoms.first() {
                                return !matches!(
                                    first_atom.content_element,
                                    ContentElement::ParagraphProperties { .. }
                                );
                            }
                            return true;
                        }
                        false
                    })
                    .count();
            }
        }
    }

    // C# lines 4798-4867: Build new sequence with proper splits
    let mut new_sequence = Vec::new();

    // C# lines 4800-4801: Calculate boundaries
    let before_common_paragraph_left =
        units1.len() - remaining_in_left_paragraph - count_common_at_end;
    let before_common_paragraph_right =
        units2.len() - remaining_in_right_paragraph - count_common_at_end;

    // C# lines 4803-4830: Handle "before common paragraph" segment
    if before_common_paragraph_left != 0 && before_common_paragraph_right == 0 {
        new_sequence.push(CorrelatedSequence::deleted(
            units1[..before_common_paragraph_left].to_vec(),
        ));
    } else if before_common_paragraph_left == 0 && before_common_paragraph_right != 0 {
        new_sequence.push(CorrelatedSequence::inserted(
            units2[..before_common_paragraph_right].to_vec(),
        ));
    } else if before_common_paragraph_left != 0 && before_common_paragraph_right != 0 {
        new_sequence.push(CorrelatedSequence::unknown(
            units1[..before_common_paragraph_left].to_vec(),
            units2[..before_common_paragraph_right].to_vec(),
        ));
    }
    // else both == 0: nothing to do

    // C# lines 4832-4859: Handle "remaining in paragraph" segment
    if remaining_in_left_paragraph != 0 && remaining_in_right_paragraph == 0 {
        new_sequence.push(CorrelatedSequence::deleted(
            units1[before_common_paragraph_left
                ..before_common_paragraph_left + remaining_in_left_paragraph]
                .to_vec(),
        ));
    } else if remaining_in_left_paragraph == 0 && remaining_in_right_paragraph != 0 {
        new_sequence.push(CorrelatedSequence::inserted(
            units2[before_common_paragraph_right
                ..before_common_paragraph_right + remaining_in_right_paragraph]
                .to_vec(),
        ));
    } else if remaining_in_left_paragraph != 0 && remaining_in_right_paragraph != 0 {
        new_sequence.push(CorrelatedSequence::unknown(
            units1[before_common_paragraph_left
                ..before_common_paragraph_left + remaining_in_left_paragraph]
                .to_vec(),
            units2[before_common_paragraph_right
                ..before_common_paragraph_right + remaining_in_right_paragraph]
                .to_vec(),
        ));
    }
    // else both == 0: nothing to do

    // C# lines 4861-4865: Add Equal sequence for common end
    new_sequence.push(CorrelatedSequence::equal(
        units1[units1.len() - count_common_at_end..].to_vec(),
        units2[units2.len() - count_common_at_end..].to_vec(),
    ));

    Some(new_sequence)
}

/// Full LCS algorithm with edge case handling
///
/// Matches C# DoLcsAlgorithm (WmlComparer.cs:6148-6724+)
///
/// This is the fallback when ProcessCorrelatedHashes and FindCommonAtBeginningAndEnd
/// don't find matches. It handles complex cases like mixed content types.
fn do_lcs_algorithm(
    unknown: &CorrelatedSequence,
    settings: &WmlComparerSettings,
) -> Vec<CorrelatedSequence> {
    let units1 = unknown.units1.as_ref();
    let units2 = unknown.units2.as_ref();

    // Handle empty cases
    match (units1, units2) {
        (Some(u1), Some(u2)) if u1.is_empty() && u2.is_empty() => {
            return Vec::new();
        }
        (Some(u1), _) if !u1.is_empty() && units2.map(|u| u.is_empty()).unwrap_or(true) => {
            return vec![CorrelatedSequence::deleted(u1.clone())];
        }
        (_, Some(u2)) if !u2.is_empty() && units1.map(|u| u.is_empty()).unwrap_or(true) => {
            return vec![CorrelatedSequence::inserted(u2.clone())];
        }
        (None, None) | (Some(_), None) | (None, Some(_)) => {
            // Handle malformed input
            let mut result = Vec::new();
            if let Some(u1) = units1 {
                if !u1.is_empty() {
                    result.push(CorrelatedSequence::deleted(u1.clone()));
                }
            }
            if let Some(u2) = units2 {
                if !u2.is_empty() {
                    result.push(CorrelatedSequence::inserted(u2.clone()));
                }
            }
            return result;
        }
        _ => {}
    }

    let units1 = units1.unwrap();
    let units2 = units2.unwrap();

    if only_paragraph_groups(units1) && only_paragraph_groups(units2) {
        let mut left_tabular = 0usize;
        let mut right_tabular = 0usize;
        let mut left_non_tabular = 0usize;
        let mut right_non_tabular = 0usize;
        for unit in units1 {
            if let Some(group) = unit.as_group() {
                if group_is_tabular(group) {
                    left_tabular += 1;
                } else if !group_is_empty(group) {
                    left_non_tabular += 1;
                }
            }
        }

        for unit in units2 {
            if let Some(group) = unit.as_group() {
                if group_is_tabular(group) {
                    right_tabular += 1;
                } else if !group_is_empty(group) {
                    right_non_tabular += 1;
                }
            }
        }

        if left_tabular > 0
            && right_tabular > 0
            && left_non_tabular == 0
            && right_non_tabular == 0
        {
            return flatten_and_create_unknown(units1, units2);
        }
    }

    // Find longest common subsequence using SHA1Hash
    let mut best_length = 0usize;
    let mut best_i1: isize = -1;
    let mut best_i2: isize = -1;

    let tabular_word_sequence = is_tabular_word_sequence(units1, units2);

    if tabular_word_sequence
        && units1.iter().all(|u| u.as_word().is_some())
        && units2.iter().all(|u| u.as_word().is_some())
        && units1.len() <= TABULAR_LCS_DP_THRESHOLD
        && units2.len() <= TABULAR_LCS_DP_THRESHOLD
    {
        return lcs_dp_for_tabular(units1, units2);
    }

    // Optimization: Index units2 by hash for O(1) lookup
    let mut units2_index: HashMap<&str, Vec<usize>> = HashMap::new();
    for (i2, unit) in units2.iter().enumerate() {
        if let Some(hash) = matchable_hash(unit, tabular_word_sequence) {
            units2_index.entry(hash).or_default().push(i2);
        }
    }

    for i1 in 0..units1.len().saturating_sub(best_length) {
        let Some(hash1) = matchable_hash(&units1[i1], tabular_word_sequence) else {
            continue;
        };

        if let Some(candidates) = units2_index.get(hash1) {
            for &i2 in candidates {
                // Optimization: Skip if we can't beat best_length
                if units2.len() - i2 <= best_length {
                    continue;
                }

                let mut seq_length = 0usize;
                let mut cur_i1 = i1;
                let mut cur_i2 = i2;

                while cur_i1 < units1.len() && cur_i2 < units2.len() {
                    let Some(cur_hash1) = matchable_hash(&units1[cur_i1], tabular_word_sequence)
                    else {
                        break;
                    };
                    let Some(cur_hash2) = matchable_hash(&units2[cur_i2], tabular_word_sequence)
                    else {
                        break;
                    };
                    if cur_hash1 == cur_hash2 {
                        cur_i1 += 1;
                        cur_i2 += 1;
                        seq_length += 1;
                    } else {
                        break;
                    }
                }

                if seq_length > best_length {
                    best_length = seq_length;
                    best_i1 = i1 as isize;
                    best_i2 = i2 as isize;
                }
            }
        }
    }

    // Never start a common section with a paragraph mark
    while best_length > 1 && best_i1 >= 0 {
        let first = &units1[best_i1 as usize];
        if let Some(word) = first.as_word() {
            if word.is_paragraph_mark() {
                best_length -= 1;
                if best_length == 0 {
                    best_i1 = -1;
                    best_i2 = -1;
                } else {
                    best_i1 += 1;
                    best_i2 += 1;
                }
                continue;
            }
        }
        break;
    }

    // Check if only paragraph mark
    let is_only_paragraph_mark = if best_length == 1 && best_i1 >= 0 {
        units1[best_i1 as usize]
            .as_word()
            .map(|w| w.is_paragraph_mark())
            .unwrap_or(false)
    } else {
        false
    };

    // Don't use empty or near-empty paragraph groups as anchor points
    // This prevents splitting similar paragraphs onto opposite sides of an empty anchor
    if best_length > 0 && best_i1 >= 0 {
        let all_groups = units1[best_i1 as usize..best_i1 as usize + best_length]
            .iter()
            .all(|u| u.as_group().is_some());
        if all_groups {
            // Check if any matched group has meaningful content (more than just whitespace/marks)
            let has_meaningful_content = units1[best_i1 as usize..best_i1 as usize + best_length]
                .iter()
                .any(|u| {
                    if let Some(g) = u.as_group() {
                        // Count atoms that are actual text content (not just paragraph marks)
                        let content_atoms = g.descendant_atoms().iter().filter(|a| {
                            matches!(&a.content_element, ContentElement::Text(c) if !c.is_whitespace())
                        }).count();
                        content_atoms > 0
                    } else {
                        false
                    }
                });
            if !has_meaningful_content {
                best_i1 = -1;
                best_i2 = -1;
                best_length = 0;
            }
        }
    }

    // Don't match just a single space character
    if best_length == 1 && best_i2 >= 0 {
        if let Some(word) = units2[best_i2 as usize].as_word() {
            if word.text() == " " {
                best_i1 = -1;
                best_i2 = -1;
                best_length = 0;
            }
        }
    }

    // C# lines 6295-6330: Don't match only word break characters
    if best_length > 0 && best_length <= 3 && best_i1 >= 0 {
        let common_seq: Vec<_> = units1[best_i1 as usize..best_i1 as usize + best_length].to_vec();
        // Check if all are ComparisonUnitWord
        let all_are_words = common_seq.iter().all(|cs| cs.as_word().is_some());
        if all_are_words {
            // Check if any word has content other than word split chars
            let has_content_other_than_split = common_seq.iter().any(|cs| {
                if let Some(word) = cs.as_word() {
                    // Check if any atom is not text
                    let has_non_text = word
                        .atoms
                        .iter()
                        .any(|atom| !matches!(atom.content_element, ContentElement::Text(_)));
                    if has_non_text {
                        return true;
                    }
                    // Check if text is not just word separator
                    let has_non_separator = word.atoms.iter().any(|atom| {
                        if let ContentElement::Text(c) = atom.content_element {
                            // Chinese/Japanese/Korean characters (CJK)
                            let is_cjk = (c as u32) >= 0x4e00 && (c as u32) <= 0x9fff;
                            if is_cjk {
                                return false; // CJK chars are word separators
                            }
                            // Check common word separators
                            !settings.is_word_separator(c)
                        } else {
                            true // Non-text atoms count as content
                        }
                    });
                    return has_non_separator;
                }
                true
            });
            if !has_content_other_than_split {
                best_i1 = -1;
                best_i2 = -1;
                best_length = 0;
            }
        }
    }

    // Don't match sequences consisting only of common stopwords
    // Words like "and", "shall", "the" appear in almost any sentence and
    // don't represent meaningful shared content worth using as anchors
    if best_length > 0 && best_length <= 5 && best_i1 >= 0 {
        let common_seq: Vec<_> = units1[best_i1 as usize..best_i1 as usize + best_length].to_vec();
        let all_are_words = common_seq.iter().all(|cs| cs.as_word().is_some());

        if all_are_words {
            // Extract the actual text content (lowercased) from the sequence
            // Each Word contains multiple atoms (characters), so we concatenate them per word
            let text_tokens: Vec<String> = common_seq
                .iter()
                .filter_map(|cs| cs.as_word())
                .map(|word| {
                    // Concatenate all text characters in this word
                    word.atoms
                        .iter()
                        .filter_map(|atom| {
                            if let ContentElement::Text(c) = atom.content_element {
                                Some(c.to_lowercase().to_string())
                            } else {
                                None
                            }
                        })
                        .collect::<String>()
                })
                .filter(|s| !s.trim().is_empty() && s.chars().any(|c| c.is_alphanumeric()))
                .collect();

            // Common stopwords that shouldn't anchor comparisons on their own
            const STOPWORDS: &[&str] = &[
                "a", "an", "the", "and", "or", "but", "nor", "for", "yet", "so", "is", "are",
                "was", "were", "be", "been", "being", "am", "has", "have", "had", "do", "does",
                "did", "shall", "will", "would", "could", "should", "may", "might", "must", "can",
                "to", "of", "in", "on", "at", "by", "with", "from", "into", "upon", "as", "if",
                "that", "this", "these", "those", "which", "who", "whom", "it", "its", "he", "she",
                "they", "we", "you", "i", "not", "no", "any", "all", "each", "every", "both",
                "either", "neither", "such", "other", "another", "same", "own",
            ];

            let all_stopwords = text_tokens
                .iter()
                .all(|token| STOPWORDS.contains(&token.as_str()));

            if all_stopwords && !text_tokens.is_empty() {
                best_i1 = -1;
                best_i2 = -1;
                best_length = 0;
            }
        }
    }

    // Apply detail threshold for text-only sequences
    // Skip if matched sequence contains structural elements (non-text atoms)
    if !is_only_paragraph_mark && best_length > 0 && best_i1 >= 0 {
        let all_words1 = units1.iter().all(|u| u.as_word().is_some());
        let all_words2 = units2.iter().all(|u| u.as_word().is_some());

        if all_words1 && all_words2 {
            // Check if matched sequence contains structural (non-text) elements
            // If so, don't discard - structural elements like tabs, footnotes should be preserved
            let matched_seq = &units1[best_i1 as usize..best_i1 as usize + best_length];
            let contains_structural = matched_seq.iter().any(|u| {
                if let Some(word) = u.as_word() {
                    word.atoms
                        .iter()
                        .any(|atom| !matches!(atom.content_element, ContentElement::Text(_)))
                } else {
                    false
                }
            });

            if !contains_structural {
                let max_len = if tabular_word_sequence {
                    let len1 = units1
                        .iter()
                        .filter(|u| matchable_hash(u, true).is_some())
                        .count();
                    let len2 = units2
                        .iter()
                        .filter(|u| matchable_hash(u, true).is_some())
                        .count();
                    len1.max(len2).max(1)
                } else {
                    units1.len().max(units2.len())
                };
                let ratio = best_length as f64 / max_len as f64;
                // Use a lower threshold for pure-text word-level matches
                // The main detail_threshold controls paragraph-level decisions (flatten vs block)
                // For word-level matches within a paragraph, we want to be more permissive
                // to find matches like "Landlord may request" (7 tokens in 148 = 4.7%)
                // which would fail the 15% threshold but are meaningful
                let text_match_threshold = settings.detail_threshold / 5.0; // 3% if default is 15%
                let has_strong_anchor = matched_seq.iter().any(is_strong_anchor_unit);
                if ratio < text_match_threshold && !has_strong_anchor {
                    best_i1 = -1;
                    best_i2 = -1;
                    best_length = 0;
                }
            }
        }
    }

    // If no match found, handle special cases
    if best_i1 < 0 || best_i2 < 0 {
        return handle_no_match_cases(units1, units2, settings);
    }

    // Build result with paragraph-aware prefix, match, and suffix
    // C# WmlComparer.cs lines 6927-7130
    let current_i1 = best_i1 as usize;
    let current_i2 = best_i2 as usize;
    let current_longest_common_sequence_length = best_length;
    let mut new_sequence = Vec::new();

    // ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // C# lines 6927-6987: Calculate paragraph boundaries
    // Here we have the longest common subsequence.
    // But it may start in the middle of a paragraph.
    // Therefore need to dispose of the content from the beginning of the longest common subsequence to the beginning of the paragraph.
    // This should be in a separate unknown region.
    // If currentLongestCommonSequenceLength != 0, and if it contains a paragraph mark, then if there are comparison units
    // in the same paragraph before the common (in either version) then we want to put all of those comparison units into
    // a single unknown, where they must be resolved against each other. We don't want those comparison units to go into
    // the middle unknown comparison unit.

    let mut remaining_in_left_paragraph = 0usize;
    let mut remaining_in_right_paragraph = 0usize;

    if current_longest_common_sequence_length != 0 {
        // C# lines 6940-6944: Get the common sequence
        let common_seq: Vec<_> =
            units1[current_i1..current_i1 + current_longest_common_sequence_length].to_vec();

        // C# lines 6946-6955: Check if first of common seq is a Word and contains paragraph marks
        if let Some(first_of_common) = common_seq.first() {
            if first_of_common.as_word().is_some() {
                // Are there any paragraph marks in the common seq?
                let has_paragraph_mark = common_seq.iter().any(|cu| {
                    if let Some(word) = cu.as_word() {
                        if let Some(first_atom) = word.atoms.first() {
                            return matches!(
                                first_atom.content_element,
                                ContentElement::ParagraphProperties { .. }
                            );
                        }
                    }
                    false
                });

                if has_paragraph_mark {
                    // C# lines 6957-6970: Calculate remainingInLeftParagraph
                    // Count units before LCS match that are in the same paragraph (no pPr)
                    remaining_in_left_paragraph = units1[..current_i1]
                        .iter()
                        .rev()
                        .take_while(|cu| {
                            if let Some(word) = cu.as_word() {
                                if let Some(first_atom) = word.atoms.first() {
                                    // Continue while NOT a paragraph mark
                                    return !matches!(
                                        first_atom.content_element,
                                        ContentElement::ParagraphProperties { .. }
                                    );
                                }
                                // No atoms means continue
                                return true;
                            }
                            // Not a word, stop
                            false
                        })
                        .count();

                    // C# lines 6971-6984: Calculate remainingInRightParagraph
                    remaining_in_right_paragraph = units2[..current_i2]
                        .iter()
                        .rev()
                        .take_while(|cu| {
                            if let Some(word) = cu.as_word() {
                                if let Some(first_atom) = word.atoms.first() {
                                    return !matches!(
                                        first_atom.content_element,
                                        ContentElement::ParagraphProperties { .. }
                                    );
                                }
                                return true;
                            }
                            false
                        })
                        .count();
                }
            }
        }
    }

    // C# lines 6989-6991: Calculate boundaries
    let count_before_current_paragraph_left = current_i1 - remaining_in_left_paragraph;
    let count_before_current_paragraph_right = current_i2 - remaining_in_right_paragraph;

    // C# lines 6992-7028: Handle content BEFORE the current paragraph
    // This is content that belongs to a previous paragraph
    if count_before_current_paragraph_left > 0 && count_before_current_paragraph_right == 0 {
        new_sequence.push(CorrelatedSequence::deleted(
            units1[..count_before_current_paragraph_left].to_vec(),
        ));
    } else if count_before_current_paragraph_left == 0 && count_before_current_paragraph_right > 0 {
        new_sequence.push(CorrelatedSequence::inserted(
            units2[..count_before_current_paragraph_right].to_vec(),
        ));
    } else if count_before_current_paragraph_left > 0 && count_before_current_paragraph_right > 0 {
        new_sequence.push(CorrelatedSequence::unknown(
            units1[..count_before_current_paragraph_left].to_vec(),
            units2[..count_before_current_paragraph_right].to_vec(),
        ));
    }
    // else both == 0: nothing to do (C# 7025-7028)

    // C# lines 7030-7069: Handle content WITHIN the current paragraph but before the LCS match
    if remaining_in_left_paragraph > 0 && remaining_in_right_paragraph == 0 {
        new_sequence.push(CorrelatedSequence::deleted(
            units1[count_before_current_paragraph_left
                ..count_before_current_paragraph_left + remaining_in_left_paragraph]
                .to_vec(),
        ));
    } else if remaining_in_left_paragraph == 0 && remaining_in_right_paragraph > 0 {
        new_sequence.push(CorrelatedSequence::inserted(
            units2[count_before_current_paragraph_right
                ..count_before_current_paragraph_right + remaining_in_right_paragraph]
                .to_vec(),
        ));
    } else if remaining_in_left_paragraph > 0 && remaining_in_right_paragraph > 0 {
        new_sequence.push(CorrelatedSequence::unknown(
            units1[count_before_current_paragraph_left
                ..count_before_current_paragraph_left + remaining_in_left_paragraph]
                .to_vec(),
            units2[count_before_current_paragraph_right
                ..count_before_current_paragraph_right + remaining_in_right_paragraph]
                .to_vec(),
        ));
    }
    // else both == 0: nothing to do (C# 7066-7069)

    // C# lines 7071-7081: Add the Equal sequence for the LCS match
    let middle_equal = CorrelatedSequence::equal(
        units1[current_i1..current_i1 + current_longest_common_sequence_length].to_vec(),
        units2[current_i2..current_i2 + current_longest_common_sequence_length].to_vec(),
    );
    new_sequence.push(middle_equal.clone());

    // C# lines 7084-7093: Calculate remaining content after LCS
    let end_i1 = current_i1 + current_longest_common_sequence_length;
    let end_i2 = current_i2 + current_longest_common_sequence_length;

    let remaining1: Vec<_> = units1[end_i1..].to_vec();
    let remaining2: Vec<_> = units2[end_i2..].to_vec();

    // C# lines 7095-7122: Post-LCS paragraph extension
    // Here is the point that we want to make a new unknown from this point to the end of the paragraph
    // that contains the equal parts. This will never hurt anything, and will in many cases result in
    // a better difference.
    if let Some(last_unit) = middle_equal.units1.as_ref().and_then(|u| u.last()) {
        if let Some(left_cuw) = last_unit.as_word() {
            // Get the last content atom from the word
            let last_atom = left_cuw.atoms.last();

            // If the middleEqual did not end with a paragraph mark (C# 7103)
            let ends_with_para = last_atom
                .map(|a| {
                    matches!(
                        a.content_element,
                        ContentElement::ParagraphProperties { .. }
                    )
                })
                .unwrap_or(false);

            if !ends_with_para {
                // C# lines 7105-7106: Find next paragraph marks in remaining content
                let idx1 = find_index_of_next_para_mark(&remaining1);
                let idx2 = find_index_of_next_para_mark(&remaining2);

                // C# lines 7108-7112: Create Unknown for content up to next paragraph mark
                if idx1 > 0 || idx2 > 0 {
                    new_sequence.push(CorrelatedSequence::unknown(
                        remaining1[..idx1].to_vec(),
                        remaining2[..idx2].to_vec(),
                    ));
                }

                // C# lines 7114-7118: Create Unknown for content after paragraph mark
                if idx1 < remaining1.len() || idx2 < remaining2.len() {
                    new_sequence.push(CorrelatedSequence::unknown(
                        remaining1[idx1..].to_vec(),
                        remaining2[idx2..].to_vec(),
                    ));
                }

                return new_sequence;
            }
        }
    }

    // C# lines 7124-7128: Default case - create single Unknown for all remaining content
    if !remaining1.is_empty() && remaining2.is_empty() {
        new_sequence.push(CorrelatedSequence::deleted(remaining1));
    } else if remaining1.is_empty() && !remaining2.is_empty() {
        new_sequence.push(CorrelatedSequence::inserted(remaining2));
    } else if !remaining1.is_empty() && !remaining2.is_empty() {
        new_sequence.push(CorrelatedSequence::unknown(remaining1, remaining2));
    }

    new_sequence
}

fn is_strong_anchor_unit(unit: &ComparisonUnit) -> bool {
    let Some(word) = unit.as_word() else {
        return false;
    };

    let mut text = String::new();
    for atom in word.atoms.iter() {
        if let ContentElement::Text(ch) = atom.content_element {
            text.push(ch);
        } else {
            return false;
        }
    }

    if text.len() < 3 {
        return false;
    }

    let mut has_alpha = false;
    let mut all_upper = true;
    for ch in text.chars() {
        if ch.is_ascii_alphabetic() {
            has_alpha = true;
            if !ch.is_ascii_uppercase() {
                all_upper = false;
            }
        }
    }

    if has_alpha && all_upper {
        return true;
    }

    text.chars().any(|ch| ch.is_ascii_digit())
}

/// Find index of next paragraph mark in comparison unit array
///
/// Matches C# FindIndexOfNextParaMark (WmlComparer.cs:7133-7143)
///
/// Returns the index of the first comparison unit that ends with a paragraph mark (w:pPr).
/// If no paragraph mark is found, returns the length of the array (meaning all remaining content
/// should be included).
fn find_index_of_next_para_mark(units: &[ComparisonUnit]) -> usize {
    for (i, unit) in units.iter().enumerate() {
        if let Some(word) = unit.as_word() {
            // Get the last atom from the word (C# uses DescendantContentAtoms().LastOrDefault())
            if let Some(last_atom) = word.atoms.last() {
                if matches!(
                    last_atom.content_element,
                    ContentElement::ParagraphProperties { .. }
                ) {
                    return i;
                }
            }
        }
    }
    // No paragraph mark found, return length (include all remaining content)
    units.len()
}

/// Handle cases where no LCS match was found
///
/// This handles complex document structures like mixed paragraphs and tables.
fn handle_no_match_cases(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
    _settings: &WmlComparerSettings,
) -> Vec<CorrelatedSequence> {
    // Count different group types
    let (tables1, rows1, _cells1, paras1, textboxes1, words1) = count_group_types(units1);
    let (tables2, rows2, _cells2, paras2, textboxes2, words2) = count_group_types(units2);

    let left_len = units1.len();
    let right_len = units2.len();

    // Handle mixed words/rows/textboxes
    let left_only_words_rows_textboxes = left_len == words1 + rows1 + textboxes1;
    let right_only_words_rows_textboxes = right_len == words2 + rows2 + textboxes2;

    if (words1 > 0 || words2 > 0)
        && (rows1 > 0 || rows2 > 0 || textboxes1 > 0 || textboxes2 > 0)
        && left_only_words_rows_textboxes
        && right_only_words_rows_textboxes
    {
        return handle_mixed_words_rows_textboxes(units1, units2);
    }

    // Handle mixed tables and paragraphs
    if tables1 > 0 && tables2 > 0 && paras1 > 0 && paras2 > 0 && (left_len > 1 || right_len > 1) {
        return handle_mixed_tables_paragraphs(units1, units2);
    }

    // Handle single tables with potential merged cells
    if tables1 == 1 && left_len == 1 && tables2 == 1 && right_len == 1 {
        if let Some(result) = do_lcs_algorithm_for_table(units1, units2, _settings) {
            return result;
        }
    }

    // Handle only paras/tables/textboxes - flatten and iterate
    let left_only_paras_tables = left_len == tables1 + paras1 + textboxes1;
    let right_only_paras_tables = right_len == tables2 + paras2 + textboxes2;

    if left_only_paras_tables && right_only_paras_tables {
        return flatten_and_create_unknown(units1, units2);
    }

    // Handle matching rows - flatten to cells
    if let (Some(first1), Some(first2)) = (units1.first(), units2.first()) {
        if let (Some(group1), Some(group2)) = (first1.as_group(), first2.as_group()) {
            if group1.group_type == ComparisonUnitGroupType::Row
                && group2.group_type == ComparisonUnitGroupType::Row
            {
                return handle_matching_rows(units1, units2);
            }
        }
    }

    // Handle matching cells - flatten cell contents
    // C# WmlComparer.cs lines 6771-6824
    if let (Some(first1), Some(first2)) = (units1.first(), units2.first()) {
        if let (Some(group1), Some(group2)) = (first1.as_group(), first2.as_group()) {
            if group1.group_type == ComparisonUnitGroupType::Cell
                && group2.group_type == ComparisonUnitGroupType::Cell
            {
                let mut result = Vec::new();

                let left_contents = group1.contents_as_units();
                let right_contents = group2.contents_as_units();

                result.push(CorrelatedSequence::unknown(left_contents, right_contents));

                let remainder_left: Vec<_> = units1.iter().skip(1).cloned().collect();
                let remainder_right: Vec<_> = units2.iter().skip(1).cloned().collect();

                if !remainder_left.is_empty() && remainder_right.is_empty() {
                    result.push(CorrelatedSequence::deleted(remainder_left));
                } else if remainder_left.is_empty() && !remainder_right.is_empty() {
                    result.push(CorrelatedSequence::inserted(remainder_right));
                } else if !remainder_left.is_empty() && !remainder_right.is_empty() {
                    result.push(CorrelatedSequence::unknown(remainder_left, remainder_right));
                }

                return result;
            }
        }
    }

    // C# WmlComparer.cs lines 6827-6869: Word/row conflict resolution
    if let (Some(_), Some(right_group)) = (
        units1.first().and_then(|u| u.as_word()),
        units2.first().and_then(|u| u.as_group()),
    ) {
        if right_group.group_type == ComparisonUnitGroupType::Row {
            return vec![
                CorrelatedSequence::inserted(units2.to_vec()),
                CorrelatedSequence::deleted(units1.to_vec()),
            ];
        }
    }

    if let (Some(left_group), Some(_)) = (
        units1.first().and_then(|u| u.as_group()),
        units2.first().and_then(|u| u.as_word()),
    ) {
        if left_group.group_type == ComparisonUnitGroupType::Row {
            return vec![
                CorrelatedSequence::deleted(units1.to_vec()),
                CorrelatedSequence::inserted(units2.to_vec()),
            ];
        }
    }

    // C# WmlComparer.cs lines 6871-6909: Paragraph mark priority logic
    // This determines the order of Deleted/Inserted sequences based on whether
    // each side ends with a paragraph mark (w:pPr).
    if !units1.is_empty() && !units2.is_empty() {
        // Get the last content atom from each side
        // C# equivalent: unknown.ComparisonUnitArray1.Select(cu => cu.DescendantContentAtoms().Last()).LastOrDefault()
        let last_atom_left = units1
            .iter()
            .filter_map(|cu| cu.descendant_atoms().last().cloned())
            .last();
        let last_atom_right = units2
            .iter()
            .filter_map(|cu| cu.descendant_atoms().last().cloned())
            .last();

        if let (Some(left), Some(right)) = (last_atom_left, last_atom_right) {
            let left_is_ppr = matches!(
                left.content_element,
                ContentElement::ParagraphProperties { .. }
            );
            let right_is_ppr = matches!(
                right.content_element,
                ContentElement::ParagraphProperties { .. }
            );

            if left_is_ppr && !right_is_ppr {
                // Left ends with pPr, right doesn't → Insert first, then Delete
                return vec![
                    CorrelatedSequence::inserted(units2.to_vec()),
                    CorrelatedSequence::deleted(units1.to_vec()),
                ];
            } else if !left_is_ppr && right_is_ppr {
                // Right ends with pPr, left doesn't → Delete first, then Insert
                return vec![
                    CorrelatedSequence::deleted(units1.to_vec()),
                    CorrelatedSequence::inserted(units2.to_vec()),
                ];
            }
        }
    }

    // Default: mark everything as deleted and inserted

    vec![
        CorrelatedSequence::deleted(units1.to_vec()),
        CorrelatedSequence::inserted(units2.to_vec()),
    ]
}

/// Count different group types in a unit list
fn count_group_types(units: &[ComparisonUnit]) -> (usize, usize, usize, usize, usize, usize) {
    let mut tables = 0;
    let mut rows = 0;
    let mut cells = 0;
    let mut paras = 0;
    let mut textboxes = 0;
    let mut words = 0;

    for unit in units {
        match unit {
            ComparisonUnit::Word(_) => words += 1,
            ComparisonUnit::Group(g) => match g.group_type {
                ComparisonUnitGroupType::Table => tables += 1,
                ComparisonUnitGroupType::Row => rows += 1,
                ComparisonUnitGroupType::Cell => cells += 1,
                ComparisonUnitGroupType::Paragraph => paras += 1,
                ComparisonUnitGroupType::Textbox => textboxes += 1,
            },
        }
    }

    (tables, rows, cells, paras, textboxes, words)
}

fn get_unit_type_key(u: &ComparisonUnit) -> &'static str {
    match u {
        ComparisonUnit::Word(_) => "Word",
        ComparisonUnit::Group(g) if g.group_type == ComparisonUnitGroupType::Row => "Row",
        ComparisonUnit::Group(g) if g.group_type == ComparisonUnitGroupType::Textbox => "Textbox",
        _ => "Other",
    }
}

fn group_units_by_type(units: &[ComparisonUnit]) -> Vec<(&'static str, Vec<ComparisonUnit>)> {
    if units.is_empty() {
        return Vec::new();
    }

    let mut result = Vec::new();
    let mut current_key = get_unit_type_key(&units[0]);
    let mut current_group = vec![units[0].clone()];

    for unit in units.iter().skip(1) {
        let key = get_unit_type_key(unit);
        if key == current_key {
            current_group.push(unit.clone());
        } else {
            result.push((current_key, current_group));
            current_key = key;
            current_group = vec![unit.clone()];
        }
    }

    result.push((current_key, current_group));
    result
}

fn handle_mixed_words_rows_textboxes(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Vec<CorrelatedSequence> {
    let grouped1 = group_units_by_type(units1);
    let grouped2 = group_units_by_type(units2);

    let mut result = Vec::new();
    let mut i1 = 0;
    let mut i2 = 0;

    while i1 < grouped1.len() && i2 < grouped2.len() {
        let (key1, items1) = &grouped1[i1];
        let (key2, items2) = &grouped2[i2];

        if key1 == key2 {
            result.push(CorrelatedSequence::unknown(items1.clone(), items2.clone()));
            i1 += 1;
            i2 += 1;
        } else if *key1 == "Word" {
            result.push(CorrelatedSequence::deleted(items1.clone()));
            i1 += 1;
        } else if *key2 == "Word" {
            result.push(CorrelatedSequence::inserted(items2.clone()));
            i2 += 1;
        } else {
            result.push(CorrelatedSequence::deleted(items1.clone()));
            i1 += 1;
        }
    }

    while i1 < grouped1.len() {
        let (_, items1) = &grouped1[i1];
        result.push(CorrelatedSequence::deleted(items1.clone()));
        i1 += 1;
    }

    while i2 < grouped2.len() {
        let (_, items2) = &grouped2[i2];
        result.push(CorrelatedSequence::inserted(items2.clone()));
        i2 += 1;
    }

    result
}

fn get_table_para_key(u: &ComparisonUnit) -> &'static str {
    match u.as_group().map(|g| g.group_type) {
        Some(ComparisonUnitGroupType::Table) => "Table",
        _ => "Para",
    }
}

fn group_units_table_para(units: &[ComparisonUnit]) -> Vec<(&'static str, Vec<ComparisonUnit>)> {
    if units.is_empty() {
        return Vec::new();
    }

    let mut result = Vec::new();
    let mut current_key = get_table_para_key(&units[0]);
    let mut current_group = vec![units[0].clone()];

    for unit in units.iter().skip(1) {
        let key = get_table_para_key(unit);
        if key == current_key {
            current_group.push(unit.clone());
        } else {
            result.push((current_key, current_group));
            current_key = key;
            current_group = vec![unit.clone()];
        }
    }

    result.push((current_key, current_group));
    result
}

fn handle_mixed_tables_paragraphs(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Vec<CorrelatedSequence> {
    let grouped1 = group_units_table_para(units1);
    let grouped2 = group_units_table_para(units2);

    let mut result = Vec::new();
    let mut i1 = 0;
    let mut i2 = 0;

    while i1 < grouped1.len() && i2 < grouped2.len() {
        let (key1, items1) = &grouped1[i1];
        let (key2, items2) = &grouped2[i2];

        if (key1 == &"Table" && key2 == &"Table") || (key1 == &"Para" && key2 == &"Para") {
            result.push(CorrelatedSequence::unknown(items1.clone(), items2.clone()));
            i1 += 1;
            i2 += 1;
        } else if key1 == &"Para" && key2 == &"Table" {
            result.push(CorrelatedSequence::deleted(items1.clone()));
            i1 += 1;
        } else if key1 == &"Table" && key2 == &"Para" {
            result.push(CorrelatedSequence::inserted(items2.clone()));
            i2 += 1;
        } else {
            i1 += 1;
            i2 += 1;
        }
    }

    while i1 < grouped1.len() {
        let (_, items1) = &grouped1[i1];
        result.push(CorrelatedSequence::deleted(items1.clone()));
        i1 += 1;
    }

    while i2 < grouped2.len() {
        let (_, items2) = &grouped2[i2];
        result.push(CorrelatedSequence::inserted(items2.clone()));
        i2 += 1;
    }

    result
}

/// Flatten groups and create a single Unknown sequence
fn flatten_and_create_unknown(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Vec<CorrelatedSequence> {
    let flattened1 = flatten_units_for_unknown(units1);
    let flattened2 = flatten_units_for_unknown(units2);

    vec![CorrelatedSequence::unknown(flattened1, flattened2)]
}

fn flatten_units_for_unknown(units: &[ComparisonUnit]) -> Vec<ComparisonUnit> {
    units
        .iter()
        .flat_map(|u| match u {
            ComparisonUnit::Group(g) => match &g.contents {
                super::comparison_unit::ComparisonUnitGroupContents::Words(words) => words
                    .iter()
                    .map(|w| ComparisonUnit::Word(w.clone()))
                    .collect::<Vec<_>>(),
                super::comparison_unit::ComparisonUnitGroupContents::Groups(groups) => {
                    flatten_group_children_for_unknown(g, groups)
                }
            },
            ComparisonUnit::Word(w) => vec![ComparisonUnit::Word(w.clone())],
        })
        .collect()
}

fn flatten_group_children_for_unknown(
    parent: &ComparisonUnitGroup,
    groups: &[ComparisonUnitGroup],
) -> Vec<ComparisonUnit> {
    let mut flattened = Vec::new();

    for group in groups {
        match &group.contents {
            super::comparison_unit::ComparisonUnitGroupContents::Words(words)
                if group.group_type == parent.group_type =>
            {
                flattened.extend(words.iter().map(|w| ComparisonUnit::Word(w.clone())));
            }
            _ => {
                flattened.push(ComparisonUnit::Group(group.clone()));
            }
        }
    }

    flattened
}

fn tabular_match_weight(unit1: &ComparisonUnit, unit2: &ComparisonUnit) -> Option<usize> {
    if is_paragraph_mark_unit(unit1) || is_paragraph_mark_unit(unit2) {
        return None;
    }
    if unit1.hash() != unit2.hash() {
        return None;
    }
    if is_tab_only_unit(unit1) && is_tab_only_unit(unit2) {
        return Some(0);
    }
    Some(1)
}

fn lcs_dp_for_tabular(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Vec<CorrelatedSequence> {
    let m = units1.len();
    let n = units2.len();

    let mut dp = vec![vec![0usize; n + 1]; m + 1];
    for i in 1..=m {
        for j in 1..=n {
            let mut best = dp[i - 1][j].max(dp[i][j - 1]);
            if let Some(weight) = tabular_match_weight(&units1[i - 1], &units2[j - 1]) {
                best = best.max(dp[i - 1][j - 1] + weight);
            }
            dp[i][j] = best;
        }
    }

    enum Op {
        Equal(ComparisonUnit, ComparisonUnit),
        Delete(ComparisonUnit),
        Insert(ComparisonUnit),
    }

    let mut ops = Vec::new();
    let mut i = m;
    let mut j = n;
    while i > 0 || j > 0 {
        if i > 0
            && j > 0
            && tabular_match_weight(&units1[i - 1], &units2[j - 1])
                .is_some_and(|w| dp[i][j] == dp[i - 1][j - 1] + w)
        {
            ops.push(Op::Equal(units1[i - 1].clone(), units2[j - 1].clone()));
            i -= 1;
            j -= 1;
        } else if j > 0 && (i == 0 || dp[i][j - 1] >= dp[i - 1][j]) {
            ops.push(Op::Insert(units2[j - 1].clone()));
            j -= 1;
        } else if i > 0 {
            ops.push(Op::Delete(units1[i - 1].clone()));
            i -= 1;
        }
    }
    ops.reverse();

    let mut sequences: Vec<CorrelatedSequence> = Vec::new();
    for op in ops {
        match op {
            Op::Equal(u1, u2) => {
                if let Some(last) = sequences.last_mut() {
                    if last.status == CorrelationStatus::Equal {
                        if let Some(ref mut left) = last.units1 {
                            left.push(u1);
                        }
                        if let Some(ref mut right) = last.units2 {
                            right.push(u2);
                        }
                        continue;
                    }
                }
                sequences.push(CorrelatedSequence::equal(vec![u1], vec![u2]));
            }
            Op::Delete(u1) => {
                if let Some(last) = sequences.last_mut() {
                    if last.status == CorrelationStatus::Deleted {
                        if let Some(ref mut left) = last.units1 {
                            left.push(u1);
                        }
                        continue;
                    }
                }
                sequences.push(CorrelatedSequence::deleted(vec![u1]));
            }
            Op::Insert(u2) => {
                if let Some(last) = sequences.last_mut() {
                    if last.status == CorrelationStatus::Inserted {
                        if let Some(ref mut right) = last.units2 {
                            right.push(u2);
                        }
                        continue;
                    }
                }
                sequences.push(CorrelatedSequence::inserted(vec![u2]));
            }
        }
    }

    sequences
}

/// Handle matching rows by flattening to cells
fn handle_matching_rows(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Vec<CorrelatedSequence> {
    // Get first row from each side
    let row1 = match units1.first() {
        Some(ComparisonUnit::Group(g)) => g,
        _ => return vec![],
    };
    let row2 = match units2.first() {
        Some(ComparisonUnit::Group(g)) => g,
        _ => return vec![],
    };

    // Extract cells
    let cells1: Vec<_> = match &row1.contents {
        super::comparison_unit::ComparisonUnitGroupContents::Groups(groups) => groups
            .iter()
            .map(|g| ComparisonUnit::Group(g.clone()))
            .collect(),
        _ => return vec![],
    };
    let cells2: Vec<_> = match &row2.contents {
        super::comparison_unit::ComparisonUnitGroupContents::Groups(groups) => groups
            .iter()
            .map(|g| ComparisonUnit::Group(g.clone()))
            .collect(),
        _ => return vec![],
    };

    let mut result = Vec::new();
    let max_len = cells1.len().max(cells2.len());

    for i in 0..max_len {
        match (cells1.get(i), cells2.get(i)) {
            (Some(c1), Some(c2)) => {
                result.push(CorrelatedSequence::unknown(
                    vec![c1.clone()],
                    vec![c2.clone()],
                ));
            }
            (Some(c1), None) => {
                result.push(CorrelatedSequence::deleted(vec![c1.clone()]));
            }
            (None, Some(c2)) => {
                result.push(CorrelatedSequence::inserted(vec![c2.clone()]));
            }
            (None, None) => {}
        }
    }

    // Handle remaining rows
    if units1.len() > 1 || units2.len() > 1 {
        let remaining1: Vec<_> = units1.iter().skip(1).cloned().collect();
        let remaining2: Vec<_> = units2.iter().skip(1).cloned().collect();

        if !remaining1.is_empty() && remaining2.is_empty() {
            result.push(CorrelatedSequence::deleted(remaining1));
        } else if remaining1.is_empty() && !remaining2.is_empty() {
            result.push(CorrelatedSequence::inserted(remaining2));
        } else if !remaining1.is_empty() && !remaining2.is_empty() {
            result.push(CorrelatedSequence::unknown(remaining1, remaining2));
        }
    }

    result
}

enum DescendantAtomsFrame<'a> {
    Atoms(std::slice::Iter<'a, ComparisonUnitAtom>),
    Words(std::slice::Iter<'a, ComparisonUnitWord>),
    Groups(std::slice::Iter<'a, ComparisonUnitGroup>),
}

struct DescendantAtomsIter<'a> {
    stack: Vec<DescendantAtomsFrame<'a>>,
}

impl<'a> DescendantAtomsIter<'a> {
    fn new(unit: &'a ComparisonUnit) -> Self {
        let mut iter = Self {
            stack: Vec::with_capacity(8),
        };
        iter.push_unit(unit);
        iter
    }

    fn push_unit(&mut self, unit: &'a ComparisonUnit) {
        match unit {
            ComparisonUnit::Word(word) => {
                self.stack
                    .push(DescendantAtomsFrame::Atoms(word.atoms.iter()));
            }
            ComparisonUnit::Group(group) => self.push_group(group),
        }
    }

    fn push_group(&mut self, group: &'a ComparisonUnitGroup) {
        match &group.contents {
            ComparisonUnitGroupContents::Words(words) => {
                self.stack.push(DescendantAtomsFrame::Words(words.iter()));
            }
            ComparisonUnitGroupContents::Groups(groups) => {
                self.stack.push(DescendantAtomsFrame::Groups(groups.iter()));
            }
        }
    }
}

impl<'a> Iterator for DescendantAtomsIter<'a> {
    type Item = &'a ComparisonUnitAtom;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            let frame = self.stack.last_mut()?;
            match frame {
                DescendantAtomsFrame::Atoms(iter) => {
                    if let Some(atom) = iter.next() {
                        return Some(atom);
                    }
                    self.stack.pop();
                }
                DescendantAtomsFrame::Words(iter) => {
                    if let Some(word) = iter.next() {
                        self.stack
                            .push(DescendantAtomsFrame::Atoms(word.atoms.iter()));
                    } else {
                        self.stack.pop();
                    }
                }
                DescendantAtomsFrame::Groups(iter) => {
                    if let Some(group) = iter.next() {
                        self.push_group(group);
                    } else {
                        self.stack.pop();
                    }
                }
            }
        }
    }
}

/// Flatten correlated sequences back to a list of atoms with appropriate correlation status
///
/// This function takes the output of the LCS algorithm (correlated sequences) and
/// produces a flat list of ComparisonUnitAtom with the correct correlation status
/// set on each atom.
///
/// # Arguments
/// * `correlated` - Slice of CorrelatedSequence from the LCS algorithm
///
/// # Returns
/// A Vec of ComparisonUnitAtom with correlation_status set appropriately:
/// - Equal atoms come from matching content (uses units1)
/// - Deleted atoms come from units1 only
/// - Inserted atoms come from units2 only
/// - Unknown atoms get atoms from both sides with Unknown status
pub fn flatten_to_atoms(correlated: &[CorrelatedSequence]) -> Vec<ComparisonUnitAtom> {
    fn needs_before_ancestor_elements(atom: &ComparisonUnitAtom) -> bool {
        const VML_RELATED_ELEMENTS: &[&str] = &[
            "pict",
            "shape",
            "rect",
            "group",
            "shapetype",
            "oval",
            "line",
            "arc",
            "curve",
            "polyline",
            "roundrect",
        ];

        let is_ppr = matches!(
            atom.content_element,
            ContentElement::ParagraphProperties { .. }
        );
        let mut is_in_textbox = false;
        let mut is_vml = false;

        for ancestor in &atom.ancestor_elements {
            if ancestor.local_name == "txbxContent" {
                is_in_textbox = true;
            }
            if VML_RELATED_ELEMENTS.contains(&ancestor.local_name.as_str()) {
                is_vml = true;
            }
            if is_in_textbox && is_vml {
                break;
            }
        }

        is_ppr || is_in_textbox || is_vml
    }

    fn count_units(units: &[ComparisonUnit]) -> usize {
        units
            .iter()
            .map(|u| u.descendant_content_atoms_count())
            .sum()
    }

    let mut total_atoms = 0usize;
    for seq in correlated {
        match seq.status {
            CorrelationStatus::Equal => {
                if let Some(units2) = &seq.units2 {
                    total_atoms += count_units(units2);
                }
            }
            CorrelationStatus::Deleted => {
                if let Some(units1) = &seq.units1 {
                    total_atoms += count_units(units1);
                }
            }
            CorrelationStatus::Inserted => {
                if let Some(units2) = &seq.units2 {
                    total_atoms += count_units(units2);
                }
            }
            CorrelationStatus::Unknown => {
                if let Some(units1) = &seq.units1 {
                    total_atoms += count_units(units1);
                }
                if let Some(units2) = &seq.units2 {
                    total_atoms += count_units(units2);
                }
            }
        }
    }

    let mut result = Vec::with_capacity(total_atoms);

    for seq in correlated {
        match seq.status {
            CorrelationStatus::Equal => {
                // For Equal status, get atoms from units1 (original) and units2 (modified)
                // In C#, it uses units2 as the basis but preserves link to units1
                if let (Some(units1), Some(units2)) = (&seq.units1, &seq.units2) {
                    for (u1, u2) in units1.iter().zip(units2.iter()) {
                        let mut atoms1 = DescendantAtomsIter::new(u1);
                        let mut atoms2 = DescendantAtomsIter::new(u2);

                        while let (Some(a1), Some(a2)) = (atoms1.next(), atoms2.next()) {
                            let mut cloned = a2.clone();
                            cloned.correlation_status = ComparisonCorrelationStatus::Equal;
                            cloned.content_element_before = Some(a1.content_element.clone());
                            cloned.formatting_signature_before = a1.formatting_signature.clone();
                            if needs_before_ancestor_elements(&cloned) {
                                cloned.ancestor_elements_before =
                                    Some(a1.ancestor_elements.clone());
                            }
                            cloned.part_before = Some(a1.part_name.clone());
                            result.push(cloned);
                        }
                    }
                }
            }
            CorrelationStatus::Deleted => {
                // Deleted content comes from units1 only
                if let Some(units) = &seq.units1 {
                    for unit in units {
                        for atom in DescendantAtomsIter::new(unit) {
                            let mut cloned = atom.clone();
                            cloned.correlation_status = ComparisonCorrelationStatus::Deleted;
                            result.push(cloned);
                        }
                    }
                }
            }
            CorrelationStatus::Inserted => {
                // Inserted content comes from units2 only
                if let Some(units) = &seq.units2 {
                    for unit in units {
                        for atom in DescendantAtomsIter::new(unit) {
                            let mut cloned = atom.clone();
                            cloned.correlation_status = ComparisonCorrelationStatus::Inserted;
                            result.push(cloned);
                        }
                    }
                }
            }
            CorrelationStatus::Unknown => {
                // Unknown status indicates modified content
                // Get atoms from both sides with Unknown status
                // First add deleted atoms from units1
                if let Some(units) = &seq.units1 {
                    for unit in units {
                        for atom in DescendantAtomsIter::new(unit) {
                            let mut cloned = atom.clone();
                            cloned.correlation_status = ComparisonCorrelationStatus::Deleted;
                            result.push(cloned);
                        }
                    }
                }
                // Then add inserted atoms from units2
                if let Some(units) = &seq.units2 {
                    for unit in units {
                        for atom in DescendantAtomsIter::new(unit) {
                            let mut cloned = atom.clone();
                            cloned.correlation_status = ComparisonCorrelationStatus::Inserted;
                            result.push(cloned);
                        }
                    }
                }
            }
        }
    }

    result
}

/// Detect completely identical sources (optimization)
///
/// If both sides have the same groups with identical sha1_hash at each position,
/// we can immediately return all Equal sequences without expensive LCS computation.
fn detect_identical_sources(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Option<Vec<CorrelatedSequence>> {
    // Must have same length
    if units1.len() != units2.len() {
        return None;
    }

    // Skip for very small inputs (not worth the overhead)
    if units1.len() < 3 {
        return None;
    }

    // Check that all sha1_hashes match at corresponding positions
    let all_match = units1
        .iter()
        .zip(units2.iter())
        .all(|(u1, u2)| match (u1, u2) {
            (ComparisonUnit::Group(g1), ComparisonUnit::Group(g2)) => g1.sha1_hash == g2.sha1_hash,
            (ComparisonUnit::Word(w1), ComparisonUnit::Word(w2)) => w1.sha1_hash == w2.sha1_hash,
            _ => false,
        });

    if !all_match {
        return None;
    }

    // All units are identical - return Equal sequences
    let result: Vec<_> = units1
        .iter()
        .zip(units2.iter())
        .map(|(u1, u2)| CorrelatedSequence::equal(vec![u1.clone()], vec![u2.clone()]))
        .collect();

    Some(result)
}

/// Detect completely unrelated sources (optimization)
///
/// Matches C# DetectUnrelatedSources (WmlComparer.cs:5745-5774)
///
/// If both sides have >3 groups and no common SHA1 hashes, mark everything
/// as deleted/inserted to avoid expensive LCS computation.
fn detect_unrelated_sources(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Option<Vec<CorrelatedSequence>> {
    let groups1: Vec<_> = units1.iter().filter_map(|u| u.as_group()).collect();
    let groups2: Vec<_> = units2.iter().filter_map(|u| u.as_group()).collect();

    if groups1.len() <= 3 || groups2.len() <= 3 {
        return None;
    }

    let hashes1: Vec<_> = groups1.iter().map(|g| &g.sha1_hash).collect();
    let hashes2: Vec<_> = groups2.iter().map(|g| &g.sha1_hash).collect();

    let has_intersection = hashes1.iter().any(|h1| hashes2.contains(h1));

    if has_intersection {
        return None;
    }

    Some(vec![
        CorrelatedSequence::deleted(units1.to_vec()),
        CorrelatedSequence::inserted(units2.to_vec()),
    ])
}

/// LCS algorithm for table structures
///
/// Matches C# DoLcsAlgorithmForTable (WmlComparer.cs:7145-7255)
///
/// Handles tables with merged cells by comparing structure hashes.
fn do_lcs_algorithm_for_table(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
    _settings: &WmlComparerSettings,
) -> Option<Vec<CorrelatedSequence>> {
    let tbl_group1 = units1.first()?.as_group()?;
    let tbl_group2 = units2.first()?.as_group()?;

    if tbl_group1.group_type != ComparisonUnitGroupType::Table
        || tbl_group2.group_type != ComparisonUnitGroupType::Table
    {
        return None;
    }

    let rows1 = match &tbl_group1.contents {
        super::comparison_unit::ComparisonUnitGroupContents::Groups(groups) => groups,
        _ => return None,
    };
    let rows2 = match &tbl_group2.contents {
        super::comparison_unit::ComparisonUnitGroupContents::Groups(groups) => groups,
        _ => return None,
    };

    if rows1.len() == rows2.len() {
        let all_rows_match = rows1.iter().zip(rows2.iter()).all(|(r1, r2)| {
            r1.correlated_sha1_hash.is_some() && r1.correlated_sha1_hash == r2.correlated_sha1_hash
        });

        if all_rows_match {
            let sequences: Vec<_> = rows1
                .iter()
                .zip(rows2.iter())
                .map(|(r1, r2)| {
                    CorrelatedSequence::unknown(
                        vec![ComparisonUnit::Group(r1.clone())],
                        vec![ComparisonUnit::Group(r2.clone())],
                    )
                })
                .collect();
            return Some(sequences);
        }
    }

    let left_contains_merged = check_table_has_merged_cells(tbl_group1);
    let right_contains_merged = check_table_has_merged_cells(tbl_group2);

    if left_contains_merged || right_contains_merged {
        if tbl_group1.structure_sha1_hash.is_some()
            && tbl_group1.structure_sha1_hash == tbl_group2.structure_sha1_hash
        {
            let sequences: Vec<_> = rows1
                .iter()
                .zip(rows2.iter())
                .map(|(r1, r2)| {
                    CorrelatedSequence::unknown(
                        vec![ComparisonUnit::Group(r1.clone())],
                        vec![ComparisonUnit::Group(r2.clone())],
                    )
                })
                .collect();
            return Some(sequences);
        }

        let flattened1: Vec<_> = rows1
            .iter()
            .map(|r| ComparisonUnit::Group(r.clone()))
            .collect();
        let flattened2: Vec<_> = rows2
            .iter()
            .map(|r| ComparisonUnit::Group(r.clone()))
            .collect();

        return Some(vec![
            CorrelatedSequence::deleted(flattened1),
            CorrelatedSequence::inserted(flattened2),
        ]);
    }

    None
}

/// Check if a table contains merged cells
///
/// Examines all descendant atoms in the table group to check if any have
/// ancestors with merged cell properties (vMerge or gridSpan).
/// This corresponds to the C# check (WmlComparer.cs:7197-7203):
/// ```csharp
/// var leftContainsMerged = tblElement1
///     .Descendants()
///     .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);
/// ```
/// In our Rust implementation, merged cell status is computed during atom
/// building and stored in `AncestorInfo.has_merged_cells`.
fn check_table_has_merged_cells(table_group: &super::comparison_unit::ComparisonUnitGroup) -> bool {
    // Get all descendant atoms from the table
    for atom in table_group.descendant_atoms() {
        // Check if any ancestor (specifically tc elements) has merged cells
        for ancestor in &atom.ancestor_elements {
            if ancestor.has_merged_cells {
                return true;
            }
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::wml::comparison_unit::{
        ComparisonUnitAtom, ComparisonUnitGroup, ComparisonUnitWord, ContentElement,
    };

    fn make_word(text: &str) -> ComparisonUnitWord {
        let settings = WmlComparerSettings::default();
        let atoms: Vec<_> = text
            .chars()
            .map(|c| ComparisonUnitAtom::new(ContentElement::Text(c), vec![], "main", &settings))
            .collect();
        ComparisonUnitWord::new(atoms)
    }

    fn make_para(text: &str) -> ComparisonUnitGroup {
        let word = make_word(text);
        ComparisonUnitGroup::from_words(vec![word], ComparisonUnitGroupType::Paragraph, 0)
    }

    #[test]
    fn test_lcs_identical() {
        let units1 = vec![ComparisonUnit::Word(make_word("hello"))];
        let units2 = vec![ComparisonUnit::Word(make_word("hello"))];
        let settings = WmlComparerSettings::default();

        let result = lcs(units1, units2, &settings);

        assert_eq!(result.len(), 1);
        assert_eq!(result[0].status, CorrelationStatus::Equal);
    }

    #[test]
    fn test_lcs_deletion() {
        let units1 = vec![
            ComparisonUnit::Word(make_word("hello")),
            ComparisonUnit::Word(make_word("world")),
        ];
        let units2 = vec![ComparisonUnit::Word(make_word("hello"))];
        let settings = WmlComparerSettings::default();

        let result = lcs(units1, units2, &settings);

        // Should have Equal for "hello" and Deleted for "world"
        let has_equal = result.iter().any(|s| s.status == CorrelationStatus::Equal);
        let has_deleted = result
            .iter()
            .any(|s| s.status == CorrelationStatus::Deleted);
        assert!(has_equal);
        assert!(has_deleted);
    }

    #[test]
    fn test_lcs_insertion() {
        let units1 = vec![ComparisonUnit::Word(make_word("hello"))];
        let units2 = vec![
            ComparisonUnit::Word(make_word("hello")),
            ComparisonUnit::Word(make_word("world")),
        ];
        let settings = WmlComparerSettings::default();

        let result = lcs(units1, units2, &settings);

        let has_equal = result.iter().any(|s| s.status == CorrelationStatus::Equal);
        let has_inserted = result
            .iter()
            .any(|s| s.status == CorrelationStatus::Inserted);
        assert!(has_equal);
        assert!(has_inserted);
    }

    #[test]
    fn test_lcs_completely_different() {
        let units1 = vec![ComparisonUnit::Word(make_word("abc"))];
        let units2 = vec![ComparisonUnit::Word(make_word("xyz"))];
        let settings = WmlComparerSettings::default();

        let result = lcs(units1, units2, &settings);

        let has_deleted = result
            .iter()
            .any(|s| s.status == CorrelationStatus::Deleted);
        let has_inserted = result
            .iter()
            .any(|s| s.status == CorrelationStatus::Inserted);
        assert!(has_deleted);
        assert!(has_inserted);
    }

    #[test]
    fn test_process_correlated_hashes_too_few_groups() {
        // With less than 3 groups, should return None
        let units1 = vec![ComparisonUnit::Group(make_para("hello"))];
        let units2 = vec![ComparisonUnit::Group(make_para("hello"))];
        let unknown = CorrelatedSequence::unknown(units1, units2);
        let settings = WmlComparerSettings::default();

        let result = process_correlated_hashes(&unknown, &settings);
        assert!(result.is_none());
    }

    #[test]
    fn test_find_common_at_beginning() {
        let units1 = vec![
            ComparisonUnit::Word(make_word("hello")),
            ComparisonUnit::Word(make_word("world")),
            ComparisonUnit::Word(make_word("foo")),
        ];
        let units2 = vec![
            ComparisonUnit::Word(make_word("hello")),
            ComparisonUnit::Word(make_word("world")),
            ComparisonUnit::Word(make_word("bar")),
        ];
        let unknown = CorrelatedSequence::unknown(units1, units2);
        let settings = WmlComparerSettings::default();

        let result = find_common_at_beginning_and_end(&unknown, &settings);
        assert!(result.is_some());

        let sequences = result.unwrap();
        assert!(sequences
            .iter()
            .any(|s| s.status == CorrelationStatus::Equal));
    }

    #[test]
    fn test_find_common_at_end() {
        let units1 = vec![
            ComparisonUnit::Word(make_word("foo")),
            ComparisonUnit::Word(make_word("hello")),
            ComparisonUnit::Word(make_word("world")),
        ];
        let units2 = vec![
            ComparisonUnit::Word(make_word("bar")),
            ComparisonUnit::Word(make_word("hello")),
            ComparisonUnit::Word(make_word("world")),
        ];
        let unknown = CorrelatedSequence::unknown(units1, units2);
        let settings = WmlComparerSettings::default();

        let result = find_common_at_beginning_and_end(&unknown, &settings);
        assert!(result.is_some());

        let sequences = result.unwrap();
        assert!(sequences
            .iter()
            .any(|s| s.status == CorrelationStatus::Equal));
    }

    #[test]
    fn test_flatten_to_atoms_basic() {
        let word1 = make_word("hello");
        let word2 = make_word("world");

        let correlated = vec![
            CorrelatedSequence::equal(
                vec![ComparisonUnit::Word(word1.clone())],
                vec![ComparisonUnit::Word(word1)],
            ),
            CorrelatedSequence::deleted(vec![ComparisonUnit::Word(word2)]),
        ];

        let atoms = flatten_to_atoms(&correlated);

        assert_eq!(atoms.len(), 10);

        let equal_count = atoms
            .iter()
            .filter(|a| a.correlation_status == ComparisonCorrelationStatus::Equal)
            .count();
        let deleted_count = atoms
            .iter()
            .filter(|a| a.correlation_status == ComparisonCorrelationStatus::Deleted)
            .count();

        assert_eq!(equal_count, 5);
        assert_eq!(deleted_count, 5);
    }

    #[test]
    fn test_flatten_to_atoms_inserted() {
        let word = make_word("new");
        let correlated = vec![CorrelatedSequence::inserted(vec![ComparisonUnit::Word(
            word,
        )])];

        let atoms = flatten_to_atoms(&correlated);

        assert_eq!(atoms.len(), 3);
        assert!(atoms
            .iter()
            .all(|a| a.correlation_status == ComparisonCorrelationStatus::Inserted));
    }

    #[test]
    fn test_flatten_to_atoms_empty() {
        let correlated: Vec<CorrelatedSequence> = vec![];
        let atoms = flatten_to_atoms(&correlated);
        assert!(atoms.is_empty());
    }

    /// Helper to create a word that represents a paragraph mark (w:pPr)
    fn make_para_mark() -> ComparisonUnitWord {
        let settings = WmlComparerSettings::default();
        let atoms = vec![ComparisonUnitAtom::new(
            ContentElement::ParagraphProperties {
                element_xml: String::new(),
            },
            vec![],
            "main",
            &settings,
        )];
        ComparisonUnitWord::new(atoms)
    }

    #[test]
    fn test_split_at_paragraph_mark_no_para_mark() {
        // Test case: No paragraph mark in the units
        // Should return a single element containing all units
        let units = vec![
            ComparisonUnit::Word(make_word("hello")),
            ComparisonUnit::Word(make_word("world")),
        ];

        let result = split_at_paragraph_mark(&units);

        assert_eq!(result.len(), 1);
        assert_eq!(result[0].len(), 2);
    }

    #[test]
    fn test_split_at_paragraph_mark_with_para_mark_at_end() {
        // Test case: Paragraph mark at the end
        // Should return [hello, world] and [pPr]
        let units = vec![
            ComparisonUnit::Word(make_word("hello")),
            ComparisonUnit::Word(make_word("world")),
            ComparisonUnit::Word(make_para_mark()),
        ];

        let result = split_at_paragraph_mark(&units);

        assert_eq!(result.len(), 2);
        assert_eq!(result[0].len(), 2); // hello, world
        assert_eq!(result[1].len(), 1); // pPr
    }

    #[test]
    fn test_split_at_paragraph_mark_with_para_mark_in_middle() {
        // Test case: Paragraph mark in the middle
        // Should return [hello] and [pPr, world]
        let units = vec![
            ComparisonUnit::Word(make_word("hello")),
            ComparisonUnit::Word(make_para_mark()),
            ComparisonUnit::Word(make_word("world")),
        ];

        let result = split_at_paragraph_mark(&units);

        assert_eq!(result.len(), 2);
        assert_eq!(result[0].len(), 1); // hello
        assert_eq!(result[1].len(), 2); // pPr, world
    }

    #[test]
    fn test_split_at_paragraph_mark_para_mark_first() {
        // Test case: Paragraph mark at the beginning
        // Should return [] and [pPr, hello, world]
        let units = vec![
            ComparisonUnit::Word(make_para_mark()),
            ComparisonUnit::Word(make_word("hello")),
            ComparisonUnit::Word(make_word("world")),
        ];

        let result = split_at_paragraph_mark(&units);

        assert_eq!(result.len(), 2);
        assert_eq!(result[0].len(), 0); // empty before pPr
        assert_eq!(result[1].len(), 3); // pPr, hello, world
    }

    #[test]
    fn test_split_at_paragraph_mark_empty_input() {
        // Test case: Empty input
        // Should return a single empty vec
        let units: Vec<ComparisonUnit> = vec![];

        let result = split_at_paragraph_mark(&units);

        assert_eq!(result.len(), 1);
        assert_eq!(result[0].len(), 0);
    }
}
