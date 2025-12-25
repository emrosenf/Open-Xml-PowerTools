//! Longest Common Subsequence (LCS) algorithm for document comparison
//!
//! This is a port of the LCS algorithm from Open-Xml-PowerTools WmlComparer.
//! The algorithm finds the longest common contiguous subsequence between two
//! arrays of comparison units, then recursively processes the non-matching
//! portions.
//!
//! Key insight: This is NOT the classic LCS that finds non-contiguous matches.
//! This finds the longest contiguous matching run, which is simpler but requires
//! recursive application to find all matches.

use std::fmt;

/// Correlation status indicating how content relates between documents
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CorrelationStatus {
    /// Content appears in both documents at this position
    Equal,
    /// Content was deleted from the original document
    Deleted,
    /// Content was inserted in the modified document
    Inserted,
    /// Not yet determined - needs further processing
    Unknown,
}

impl fmt::Display for CorrelationStatus {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CorrelationStatus::Equal => write!(f, "Equal"),
            CorrelationStatus::Deleted => write!(f, "Deleted"),
            CorrelationStatus::Inserted => write!(f, "Inserted"),
            CorrelationStatus::Unknown => write!(f, "Unknown"),
        }
    }
}

/// Trait for items that can be compared using LCS
/// Items must provide a hash for equality comparison
pub trait Hashable {
    fn hash(&self) -> &str;
}

/// A correlated sequence showing how portions of two arrays relate
#[derive(Debug, Clone)]
pub struct CorrelatedSequence<T: Hashable + Clone> {
    pub status: CorrelationStatus,
    /// Items from the first (original) array, None if inserted
    pub items1: Option<Vec<T>>,
    /// Items from the second (modified) array, None if deleted
    pub items2: Option<Vec<T>>,
}

impl<T: Hashable + Clone> CorrelatedSequence<T> {
    /// Create a new Equal sequence
    pub fn equal(items1: Vec<T>, items2: Vec<T>) -> Self {
        Self {
            status: CorrelationStatus::Equal,
            items1: Some(items1),
            items2: Some(items2),
        }
    }

    /// Create a new Deleted sequence
    pub fn deleted(items1: Vec<T>) -> Self {
        Self {
            status: CorrelationStatus::Deleted,
            items1: Some(items1),
            items2: None,
        }
    }

    /// Create a new Inserted sequence
    pub fn inserted(items2: Vec<T>) -> Self {
        Self {
            status: CorrelationStatus::Inserted,
            items1: None,
            items2: Some(items2),
        }
    }
}

/// Type alias for the skip anchor predicate
pub type SkipPredicate = Box<dyn Fn(&str) -> bool + Send + Sync>;

/// Settings for the LCS algorithm
#[derive(Default)]
pub struct LcsSettings {
    /// Minimum length for a match to be considered valid.
    /// Helps avoid matching insignificant content like single spaces.
    /// Default: 1
    pub min_match_length: usize,

    /// Threshold (0-1) for minimum match ratio relative to total length.
    /// If the longest match is less than this ratio of the max length,
    /// it's considered not a real match. Default: 0.0 (any match accepted)
    pub detail_threshold: f64,

    /// Optional predicate to skip certain items from being match anchors.
    /// Useful for skipping paragraph marks, whitespace-only content, etc.
    /// The predicate receives the hash of the item.
    pub should_skip_as_anchor: Option<SkipPredicate>,
}

impl LcsSettings {
    /// Create new settings with defaults
    pub fn new() -> Self {
        Self::default()
    }

    /// Create settings with a skip predicate
    pub fn with_skip_predicate<F>(skip: F) -> Self
    where
        F: Fn(&str) -> bool + Send + Sync + 'static,
    {
        Self {
            min_match_length: 1,
            detail_threshold: 0.0,
            should_skip_as_anchor: Some(Box::new(skip)),
        }
    }

    /// Set minimum match length
    pub fn min_match_length(mut self, len: usize) -> Self {
        self.min_match_length = len;
        self
    }

    /// Set detail threshold
    pub fn detail_threshold(mut self, threshold: f64) -> Self {
        self.detail_threshold = threshold;
        self
    }
}

/// Result of finding the longest match
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MatchResult {
    /// Start index in the first array
    pub i1: usize,
    /// Start index in the second array
    pub i2: usize,
    /// Length of the match
    pub length: usize,
}

/// Find the longest common contiguous subsequence between two arrays.
///
/// This is an O(n*m) algorithm where n and m are the lengths of the input arrays.
/// It finds the longest run of consecutive items with matching hashes.
///
/// # Returns
/// The start indices and length of the match, or None if no valid match found
pub fn find_longest_match<T: Hashable>(
    items1: &[T],
    items2: &[T],
    settings: &LcsSettings,
) -> Option<MatchResult> {
    let mut best_length = 0usize;
    let mut best_i1 = 0usize;
    let mut best_i2 = 0usize;

    // Optimization: don't search positions where we can't possibly find
    // a longer match than what we already have
    for i1 in 0..items1.len().saturating_sub(best_length) {
        for i2 in 0..items2.len().saturating_sub(best_length) {
            // Count consecutive matches starting at this position
            let mut match_length = 0usize;
            let mut cur_i1 = i1;
            let mut cur_i2 = i2;

            while cur_i1 < items1.len()
                && cur_i2 < items2.len()
                && items1[cur_i1].hash() == items2[cur_i2].hash()
            {
                match_length += 1;
                cur_i1 += 1;
                cur_i2 += 1;
            }

            if match_length > best_length {
                best_length = match_length;
                best_i1 = i1;
                best_i2 = i2;
            }
        }
    }

    // Apply minimum length filter
    if best_length < settings.min_match_length {
        return None;
    }

    // Skip matches that start with items that shouldn't be anchors
    if let Some(ref skip) = settings.should_skip_as_anchor {
        while best_length > 0 && skip(items1[best_i1].hash()) {
            best_i1 += 1;
            best_i2 += 1;
            best_length -= 1;
        }
    }

    if best_length == 0 {
        return None;
    }

    // Apply detail threshold filter
    if settings.detail_threshold > 0.0 {
        let max_len = items1.len().max(items2.len());
        if max_len > 0 && (best_length as f64 / max_len as f64) < settings.detail_threshold {
            return None;
        }
    }

    Some(MatchResult {
        i1: best_i1,
        i2: best_i2,
        length: best_length,
    })
}

/// Compute the LCS-based correlation between two arrays.
///
/// This recursively finds matches and builds a list of correlated sequences
/// showing which parts are equal, deleted, or inserted.
///
/// # Arguments
/// * `items1` - The original array
/// * `items2` - The modified array
/// * `settings` - Optional settings for match thresholds
///
/// # Returns
/// Array of correlated sequences in order
pub fn compute_correlation<T: Hashable + Clone>(
    items1: &[T],
    items2: &[T],
    settings: &LcsSettings,
) -> Vec<CorrelatedSequence<T>> {
    // Handle empty array cases
    if items1.is_empty() && items2.is_empty() {
        return vec![];
    }

    if items1.is_empty() {
        return vec![CorrelatedSequence::inserted(items2.to_vec())];
    }

    if items2.is_empty() {
        return vec![CorrelatedSequence::deleted(items1.to_vec())];
    }

    // Find longest match
    let Some(match_result) = find_longest_match(items1, items2, settings) else {
        // No match found - everything is different
        return vec![
            CorrelatedSequence::deleted(items1.to_vec()),
            CorrelatedSequence::inserted(items2.to_vec()),
        ];
    };

    let mut result = Vec::new();

    // Process items before the match
    if match_result.i1 > 0 || match_result.i2 > 0 {
        let before = compute_correlation(
            &items1[..match_result.i1],
            &items2[..match_result.i2],
            settings,
        );
        result.extend(before);
    }

    // Add the matching portion
    result.push(CorrelatedSequence::equal(
        items1[match_result.i1..match_result.i1 + match_result.length].to_vec(),
        items2[match_result.i2..match_result.i2 + match_result.length].to_vec(),
    ));

    // Process items after the match
    let after_i1 = match_result.i1 + match_result.length;
    let after_i2 = match_result.i2 + match_result.length;
    if after_i1 < items1.len() || after_i2 < items2.len() {
        let after = compute_correlation(&items1[after_i1..], &items2[after_i2..], settings);
        result.extend(after);
    }

    result
}

/// Flatten a list of correlated sequences, merging adjacent sequences
/// of the same status.
pub fn flatten_correlation<T: Hashable + Clone>(
    sequences: Vec<CorrelatedSequence<T>>,
) -> Vec<CorrelatedSequence<T>> {
    if sequences.is_empty() {
        return vec![];
    }

    let mut result = Vec::new();
    let mut iter = sequences.into_iter();
    let mut current = iter.next().unwrap();

    for next in iter {
        if next.status == current.status {
            // Merge adjacent sequences of same status
            if let (Some(ref mut items1), Some(ref items1_next)) =
                (&mut current.items1, &next.items1)
            {
                items1.extend(items1_next.iter().cloned());
            }
            if let (Some(ref mut items2), Some(ref items2_next)) =
                (&mut current.items2, &next.items2)
            {
                items2.extend(items2_next.iter().cloned());
            }
        } else {
            result.push(current);
            current = next;
        }
    }

    result.push(current);
    result
}

/// Simple diff result for text comparison
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DiffType {
    Equal,
    Insert,
    Delete,
}

/// Result of a text diff operation
#[derive(Debug, Clone)]
pub struct DiffResult {
    pub diff_type: DiffType,
    pub value: String,
}

impl DiffResult {
    pub fn equal(value: String) -> Self {
        Self {
            diff_type: DiffType::Equal,
            value,
        }
    }

    pub fn insert(value: String) -> Self {
        Self {
            diff_type: DiffType::Insert,
            value,
        }
    }

    pub fn delete(value: String) -> Self {
        Self {
            diff_type: DiffType::Delete,
            value,
        }
    }
}

/// A simple hashable string wrapper for text diffing
#[derive(Debug, Clone)]
struct HashableString(String);

impl Hashable for HashableString {
    fn hash(&self) -> &str {
        &self.0
    }
}

/// Simple text diff using LCS algorithm.
/// Useful for character-by-character or word-by-word comparison.
///
/// # Arguments
/// * `text1` - Original text
/// * `text2` - Modified text
/// * `split_pattern` - How to split text into units (None = by character)
///
/// # Returns
/// Array of diff results
pub fn diff_text(text1: &str, text2: &str, split_pattern: Option<&str>) -> Vec<DiffResult> {
    // Split texts into units
    let units1: Vec<HashableString> = match split_pattern {
        None => text1.chars().map(|c| HashableString(c.to_string())).collect(),
        Some(pattern) => text1
            .split(pattern)
            .filter(|s| !s.is_empty())
            .map(|s| HashableString(s.to_string()))
            .collect(),
    };

    let units2: Vec<HashableString> = match split_pattern {
        None => text2.chars().map(|c| HashableString(c.to_string())).collect(),
        Some(pattern) => text2
            .split(pattern)
            .filter(|s| !s.is_empty())
            .map(|s| HashableString(s.to_string()))
            .collect(),
    };

    let settings = LcsSettings::default();
    let correlation = compute_correlation(&units1, &units2, &settings);

    let joiner = split_pattern.unwrap_or("");

    let mut results = Vec::new();

    for seq in correlation {
        match seq.status {
            CorrelationStatus::Equal => {
                if let Some(items) = seq.items1 {
                    let value = items.iter().map(|i| i.0.as_str()).collect::<Vec<_>>().join(joiner);
                    results.push(DiffResult::equal(value));
                }
            }
            CorrelationStatus::Deleted => {
                if let Some(items) = seq.items1 {
                    let value = items.iter().map(|i| i.0.as_str()).collect::<Vec<_>>().join(joiner);
                    results.push(DiffResult::delete(value));
                }
            }
            CorrelationStatus::Inserted => {
                if let Some(items) = seq.items2 {
                    let value = items.iter().map(|i| i.0.as_str()).collect::<Vec<_>>().join(joiner);
                    results.push(DiffResult::insert(value));
                }
            }
            CorrelationStatus::Unknown => {}
        }
    }

    results
}

#[cfg(test)]
mod tests {
    use super::*;

    #[derive(Debug, Clone)]
    struct TestItem {
        hash: String,
    }

    impl Hashable for TestItem {
        fn hash(&self) -> &str {
            &self.hash
        }
    }

    fn items(hashes: &[&str]) -> Vec<TestItem> {
        hashes
            .iter()
            .map(|h| TestItem { hash: h.to_string() })
            .collect()
    }

    #[test]
    fn test_find_longest_match_identical() {
        let items1 = items(&["a", "b", "c"]);
        let items2 = items(&["a", "b", "c"]);
        let settings = LcsSettings::default();

        let result = find_longest_match(&items1, &items2, &settings).unwrap();
        assert_eq!(result.i1, 0);
        assert_eq!(result.i2, 0);
        assert_eq!(result.length, 3);
    }

    #[test]
    fn test_find_longest_match_with_diff() {
        let items1 = items(&["a", "b", "c", "d"]);
        let items2 = items(&["x", "b", "c", "y"]);
        let settings = LcsSettings::default();

        let result = find_longest_match(&items1, &items2, &settings).unwrap();
        assert_eq!(result.i1, 1);
        assert_eq!(result.i2, 1);
        assert_eq!(result.length, 2);
    }

    #[test]
    fn test_find_longest_match_no_match() {
        let items1 = items(&["a", "b", "c"]);
        let items2 = items(&["x", "y", "z"]);
        let settings = LcsSettings::default();

        let result = find_longest_match(&items1, &items2, &settings);
        assert!(result.is_none());
    }

    #[test]
    fn test_find_longest_match_empty() {
        let items1: Vec<TestItem> = vec![];
        let items2 = items(&["a", "b"]);
        let settings = LcsSettings::default();

        let result = find_longest_match(&items1, &items2, &settings);
        assert!(result.is_none());
    }

    #[test]
    fn test_compute_correlation_identical() {
        let items1 = items(&["a", "b", "c"]);
        let items2 = items(&["a", "b", "c"]);
        let settings = LcsSettings::default();

        let result = compute_correlation(&items1, &items2, &settings);
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].status, CorrelationStatus::Equal);
    }

    #[test]
    fn test_compute_correlation_insertion() {
        let items1 = items(&["a", "c"]);
        let items2 = items(&["a", "b", "c"]);
        let settings = LcsSettings::default();

        let result = compute_correlation(&items1, &items2, &settings);
        // Should be: Equal(a), Inserted(b), Equal(c)
        assert!(result.len() >= 2);

        let has_insert = result.iter().any(|s| s.status == CorrelationStatus::Inserted);
        let has_equal = result.iter().any(|s| s.status == CorrelationStatus::Equal);
        assert!(has_insert);
        assert!(has_equal);
    }

    #[test]
    fn test_compute_correlation_deletion() {
        let items1 = items(&["a", "b", "c"]);
        let items2 = items(&["a", "c"]);
        let settings = LcsSettings::default();

        let result = compute_correlation(&items1, &items2, &settings);
        // Should be: Equal(a), Deleted(b), Equal(c)
        let has_delete = result.iter().any(|s| s.status == CorrelationStatus::Deleted);
        let has_equal = result.iter().any(|s| s.status == CorrelationStatus::Equal);
        assert!(has_delete);
        assert!(has_equal);
    }

    #[test]
    fn test_compute_correlation_empty_first() {
        let items1: Vec<TestItem> = vec![];
        let items2 = items(&["a", "b"]);
        let settings = LcsSettings::default();

        let result = compute_correlation(&items1, &items2, &settings);
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].status, CorrelationStatus::Inserted);
    }

    #[test]
    fn test_compute_correlation_empty_second() {
        let items1 = items(&["a", "b"]);
        let items2: Vec<TestItem> = vec![];
        let settings = LcsSettings::default();

        let result = compute_correlation(&items1, &items2, &settings);
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].status, CorrelationStatus::Deleted);
    }

    #[test]
    fn test_flatten_merges_adjacent_same_status() {
        let sequences = vec![
            CorrelatedSequence::deleted(items(&["a"])),
            CorrelatedSequence::deleted(items(&["b"])),
            CorrelatedSequence::equal(items(&["c"]), items(&["c"])),
        ];

        let result = flatten_correlation(sequences);
        assert_eq!(result.len(), 2);
        assert_eq!(result[0].status, CorrelationStatus::Deleted);
        assert_eq!(result[0].items1.as_ref().unwrap().len(), 2);
        assert_eq!(result[1].status, CorrelationStatus::Equal);
    }

    #[test]
    fn test_diff_text_character_level() {
        let result = diff_text("hello", "hallo", None);

        // Should detect the change from 'e' to 'a'
        let has_delete = result.iter().any(|d| d.diff_type == DiffType::Delete);
        let has_insert = result.iter().any(|d| d.diff_type == DiffType::Insert);
        let has_equal = result.iter().any(|d| d.diff_type == DiffType::Equal);

        assert!(has_equal);
        // Either we have delete/insert for 'e'/'a', or the algorithm finds a different path
        assert!(has_delete || has_insert);
    }

    #[test]
    fn test_diff_text_word_level() {
        let result = diff_text("hello world", "hello there", Some(" "));

        // Should have equal("hello"), delete("world"), insert("there")
        assert!(result.len() >= 2);

        let equal_hello = result
            .iter()
            .any(|d| d.diff_type == DiffType::Equal && d.value == "hello");
        assert!(equal_hello);
    }

    #[test]
    fn test_diff_text_identical() {
        let result = diff_text("hello", "hello", None);

        assert_eq!(result.len(), 1);
        assert_eq!(result[0].diff_type, DiffType::Equal);
        assert_eq!(result[0].value, "hello");
    }

    #[test]
    fn test_diff_text_completely_different() {
        let result = diff_text("abc", "xyz", None);

        let has_delete = result.iter().any(|d| d.diff_type == DiffType::Delete);
        let has_insert = result.iter().any(|d| d.diff_type == DiffType::Insert);

        assert!(has_delete);
        assert!(has_insert);
    }

    #[test]
    fn test_detail_threshold() {
        let items1 = items(&["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"]);
        let items2 = items(&["x", "y", "z", "d", "w"]);
        
        // With high threshold, small match shouldn't count
        let settings = LcsSettings {
            min_match_length: 1,
            detail_threshold: 0.5, // Require 50% match
            should_skip_as_anchor: None,
        };

        let result = find_longest_match(&items1, &items2, &settings);
        // The match "d" is only 1/10 = 10%, less than 50% threshold
        assert!(result.is_none());
    }

    #[test]
    fn test_min_match_length() {
        let items1 = items(&["a", "b", "c"]);
        let items2 = items(&["x", "b", "y"]);

        // Require at least 2 items to match
        let settings = LcsSettings {
            min_match_length: 2,
            detail_threshold: 0.0,
            should_skip_as_anchor: None,
        };

        let result = find_longest_match(&items1, &items2, &settings);
        assert!(result.is_none());
    }
}
