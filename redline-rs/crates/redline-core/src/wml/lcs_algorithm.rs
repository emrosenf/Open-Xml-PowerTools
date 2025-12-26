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
    ComparisonCorrelationStatus, ComparisonUnit, ComparisonUnitAtom, ComparisonUnitGroupType,
    generate_unid,
};
use super::settings::WmlComparerSettings;

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
    // Initialize with one Unknown sequence containing entire arrays
    let initial = CorrelatedSequence::unknown(units1, units2);
    let mut cs_list = vec![initial];

    loop {
        // Find first Unknown sequence
        let unknown_idx = cs_list
            .iter()
            .position(|cs| cs.status == CorrelationStatus::Unknown);

        let Some(idx) = unknown_idx else {
            // No more Unknown sequences - we're done
            return cs_list;
        };

        // Extract the unknown sequence for processing
        let unknown = cs_list.remove(idx);

        let unknown = set_after_unids(unknown);

        // Try ProcessCorrelatedHashes first (fastest)
        let new_sequences = process_correlated_hashes(&unknown, settings)
            // Then try FindCommonAtBeginningAndEnd
            .or_else(|| find_common_at_beginning_and_end(&unknown, settings))
            // Finally fall back to DoLcsAlgorithm
            .unwrap_or_else(|| do_lcs_algorithm(&unknown, settings));

        // Insert new sequences at the position of the old unknown
        // (Reverse to maintain order when inserting at same position)
        for seq in new_sequences.into_iter().rev() {
            cs_list.insert(idx, seq);
        }
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
    ) && matches!(
        first2.group_type,
        ComparisonUnitGroupType::Paragraph
            | ComparisonUnitGroupType::Table
            | ComparisonUnitGroupType::Row
    );

    if !valid_types {
        return None;
    }

    // Find longest common sequence using CorrelatedSHA1Hash
    let mut best_length = 0usize;
    let mut best_atom_count = 0usize;
    let mut best_i1 = 0usize;
    let mut best_i2 = 0usize;

    for i1 in 0..units1.len() {
        for i2 in 0..units2.len() {
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

    // Apply thresholds based on sequence length and atom count
    let do_correlation = if best_length == 1 {
        // Single group needs 16+ atoms on each side
        let atoms1 = units1[best_i1].descendant_atoms().len();
        let atoms2 = units2[best_i2].descendant_atoms().len();
        atoms1 > 16 && atoms2 > 16
    } else if best_length > 1 && best_length <= 3 {
        // 2-3 groups need 32+ atoms total on each side
        let atoms1: usize = units1[best_i1..best_i1 + best_length]
            .iter()
            .map(|u| u.descendant_atoms().len())
            .sum();
        let atoms2: usize = units2[best_i2..best_i2 + best_length]
            .iter()
            .map(|u| u.descendant_atoms().len())
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
        result.push(CorrelatedSequence::deleted(
            units1[..best_i1].to_vec(),
        ));
    } else if best_i1 == 0 && best_i2 > 0 {
        result.push(CorrelatedSequence::inserted(
            units2[..best_i2].to_vec(),
        ));
    } else if best_i1 > 0 && best_i2 > 0 {
        result.push(CorrelatedSequence::unknown(
            units1[..best_i1].to_vec(),
            units2[..best_i2].to_vec(),
        ));
    }

    // Add matched groups as individual Unknown sequences (for further processing)
    for i in 0..best_length {
        result.push(CorrelatedSequence::unknown(
            vec![units1[best_i1 + i].clone()],
            vec![units2[best_i2 + i].clone()],
        ));
    }

    // Handle suffix (after match)
    let end_i1 = best_i1 + best_length;
    let end_i2 = best_i2 + best_length;

    if end_i1 < units1.len() && end_i2 == units2.len() {
        result.push(CorrelatedSequence::deleted(
            units1[end_i1..].to_vec(),
        ));
    } else if end_i1 == units1.len() && end_i2 < units2.len() {
        result.push(CorrelatedSequence::inserted(
            units2[end_i2..].to_vec(),
        ));
    } else if end_i1 < units1.len() && end_i2 < units2.len() {
        result.push(CorrelatedSequence::unknown(
            units1[end_i1..].to_vec(),
            units2[end_i2..].to_vec(),
        ));
    }

    Some(result)
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
            result.push(CorrelatedSequence::deleted(
                units1[count_common_at_beginning..].to_vec(),
            ));
        } else if remaining_left == 0 && remaining_right > 0 {
            result.push(CorrelatedSequence::inserted(
                units2[count_common_at_beginning..].to_vec(),
            ));
        } else if remaining_left > 0 && remaining_right > 0 {
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

    // Check if only paragraph mark
    let is_only_paragraph_mark = if count_common_at_end == 1 {
        let first_common_idx1 = units1.len() - count_common_at_end;
        units1[first_common_idx1]
            .as_word()
            .map(|w| w.is_paragraph_mark())
            .unwrap_or(false)
    } else {
        false
    };

    // Apply detail threshold (unless it's just a paragraph mark)
    if !is_only_paragraph_mark && count_common_at_end > 0 {
        let ratio = count_common_at_end as f64 / length_to_compare as f64;
        if ratio < settings.detail_threshold {
            count_common_at_end = 0;
        }
    }

    if count_common_at_end > 0 {
        let mut result = Vec::new();

        let prefix_len1 = units1.len() - count_common_at_end;
        let prefix_len2 = units2.len() - count_common_at_end;

        // Handle prefix (different content)
        if prefix_len1 > 0 && prefix_len2 == 0 {
            result.push(CorrelatedSequence::deleted(
                units1[..prefix_len1].to_vec(),
            ));
        } else if prefix_len1 == 0 && prefix_len2 > 0 {
            result.push(CorrelatedSequence::inserted(
                units2[..prefix_len2].to_vec(),
            ));
        } else if prefix_len1 > 0 && prefix_len2 > 0 {
            result.push(CorrelatedSequence::unknown(
                units1[..prefix_len1].to_vec(),
                units2[..prefix_len2].to_vec(),
            ));
        }

        // Add Equal sequence for common suffix
        result.push(CorrelatedSequence::equal(
            units1[prefix_len1..].to_vec(),
            units2[prefix_len2..].to_vec(),
        ));

        return Some(result);
    }

    None
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

    // Find longest common subsequence using SHA1Hash
    let mut best_length = 0usize;
    let mut best_i1: isize = -1;
    let mut best_i2: isize = -1;

    for i1 in 0..units1.len().saturating_sub(best_length) {
        for i2 in 0..units2.len().saturating_sub(best_length) {
            let mut seq_length = 0usize;
            let mut cur_i1 = i1;
            let mut cur_i2 = i2;

            while cur_i1 < units1.len() && cur_i2 < units2.len() {
                if units1[cur_i1].hash() == units2[cur_i2].hash() {
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

    // Apply detail threshold for text-only sequences
    if !is_only_paragraph_mark && best_length > 0 && best_i1 >= 0 {
        let all_words1 = units1.iter().all(|u| u.as_word().is_some());
        let all_words2 = units2.iter().all(|u| u.as_word().is_some());

        if all_words1 && all_words2 {
            let max_len = units1.len().max(units2.len());
            let ratio = best_length as f64 / max_len as f64;
            if ratio < settings.detail_threshold {
                best_i1 = -1;
                best_i2 = -1;
                best_length = 0;
            }
        }
    }

    // If no match found, handle special cases
    if best_i1 < 0 || best_i2 < 0 {
        return handle_no_match_cases(units1, units2, settings);
    }

    // Build result with prefix, match, and suffix
    let best_i1 = best_i1 as usize;
    let best_i2 = best_i2 as usize;
    let mut result = Vec::new();

    // Prefix
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

    // Match
    result.push(CorrelatedSequence::equal(
        units1[best_i1..best_i1 + best_length].to_vec(),
        units2[best_i2..best_i2 + best_length].to_vec(),
    ));

    // Suffix
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

    result
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
        // Could implement DoLcsAlgorithmForTable here
        // For now, just mark as changed
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
            result.push(CorrelatedSequence::unknown(
                items1.clone(),
                items2.clone(),
            ));
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
    // Extract contents from groups
    let flattened1: Vec<_> = units1
        .iter()
        .flat_map(|u| match u {
            ComparisonUnit::Group(g) => match &g.contents {
                super::comparison_unit::ComparisonUnitGroupContents::Words(words) => words
                    .iter()
                    .map(|w| ComparisonUnit::Word(w.clone()))
                    .collect::<Vec<_>>(),
                super::comparison_unit::ComparisonUnitGroupContents::Groups(groups) => groups
                    .iter()
                    .map(|g| ComparisonUnit::Group(g.clone()))
                    .collect::<Vec<_>>(),
            },
            ComparisonUnit::Word(w) => vec![ComparisonUnit::Word(w.clone())],
        })
        .collect();

    let flattened2: Vec<_> = units2
        .iter()
        .flat_map(|u| match u {
            ComparisonUnit::Group(g) => match &g.contents {
                super::comparison_unit::ComparisonUnitGroupContents::Words(words) => words
                    .iter()
                    .map(|w| ComparisonUnit::Word(w.clone()))
                    .collect::<Vec<_>>(),
                super::comparison_unit::ComparisonUnitGroupContents::Groups(groups) => groups
                    .iter()
                    .map(|g| ComparisonUnit::Group(g.clone()))
                    .collect::<Vec<_>>(),
            },
            ComparisonUnit::Word(w) => vec![ComparisonUnit::Word(w.clone())],
        })
        .collect();

    vec![CorrelatedSequence::unknown(flattened1, flattened2)]
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
        super::comparison_unit::ComparisonUnitGroupContents::Groups(groups) => {
            groups.iter().map(|g| ComparisonUnit::Group(g.clone())).collect()
        }
        _ => return vec![],
    };
    let cells2: Vec<_> = match &row2.contents {
        super::comparison_unit::ComparisonUnitGroupContents::Groups(groups) => {
            groups.iter().map(|g| ComparisonUnit::Group(g.clone())).collect()
        }
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
    let mut result = Vec::new();

    for seq in correlated {
        match seq.status {
            CorrelationStatus::Equal => {
                // For Equal status, get atoms from units1 (both sides are equivalent)
                if let Some(units) = &seq.units1 {
                    for unit in units {
                        for atom in unit.descendant_atoms() {
                            let mut cloned = atom.clone();
                            cloned.correlation_status = ComparisonCorrelationStatus::Equal;
                            result.push(cloned);
                        }
                    }
                }
            }
            CorrelationStatus::Deleted => {
                // Deleted content comes from units1 only
                if let Some(units) = &seq.units1 {
                    for unit in units {
                        for atom in unit.descendant_atoms() {
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
                        for atom in unit.descendant_atoms() {
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
                        for atom in unit.descendant_atoms() {
                            let mut cloned = atom.clone();
                            cloned.correlation_status = ComparisonCorrelationStatus::Deleted;
                            result.push(cloned);
                        }
                    }
                }
                // Then add inserted atoms from units2
                if let Some(units) = &seq.units2 {
                    for unit in units {
                        for atom in unit.descendant_atoms() {
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::wml::comparison_unit::{
        ComparisonUnitAtom, ComparisonUnitGroup, ComparisonUnitWord, ContentElement,
    };

    fn make_word(text: &str) -> ComparisonUnitWord {
        let atoms: Vec<_> = text
            .chars()
            .map(|c| ComparisonUnitAtom::new(ContentElement::Text(c), vec![], "main"))
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
        let has_deleted = result.iter().any(|s| s.status == CorrelationStatus::Deleted);
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
        let has_inserted = result.iter().any(|s| s.status == CorrelationStatus::Inserted);
        assert!(has_equal);
        assert!(has_inserted);
    }

    #[test]
    fn test_lcs_completely_different() {
        let units1 = vec![ComparisonUnit::Word(make_word("abc"))];
        let units2 = vec![ComparisonUnit::Word(make_word("xyz"))];
        let settings = WmlComparerSettings::default();

        let result = lcs(units1, units2, &settings);

        let has_deleted = result.iter().any(|s| s.status == CorrelationStatus::Deleted);
        let has_inserted = result.iter().any(|s| s.status == CorrelationStatus::Inserted);
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
        assert!(sequences.iter().any(|s| s.status == CorrelationStatus::Equal));
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
        assert!(sequences.iter().any(|s| s.status == CorrelationStatus::Equal));
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
}
