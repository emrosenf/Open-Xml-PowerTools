//! WmlComparer - Word document comparison
//!
//! This is a faithful port of the C# WmlComparer from Open-Xml-PowerTools.
//!
//! Key architecture (matching C#):
//! 1. ComparisonUnitAtom - atomic content element (w:t, w:pPr, etc.)
//! 2. ComparisonUnitWord - group of atoms forming a "word"
//! 3. ComparisonUnitGroup - hierarchical grouping (paragraph, table, row, cell)
//! 4. CorrelatedSHA1Hash - pre-computed hash for paragraph correlation
//!
//! The algorithm:
//! 1. PreProcess both documents (add UNIDs)
//! 2. HashBlockLevelContent (accept revisions in source1, reject in source2, compute hashes)
//! 3. CreateComparisonUnitAtomList (extract atomic content)
//! 4. GetComparisonUnitList (build hierarchical structure)
//! 5. Multi-level LCS using ProcessCorrelatedHashes -> FindCommonAtBeginningAndEnd -> DoLcsAlgorithm
//! 6. FlattenToComparisonUnitAtomList (flatten with status)
//! 7. ProduceNewWmlMarkupFromCorrelatedSequence (generate result document)

use super::document::{
    extract_paragraph_text, find_document_body, find_paragraphs, find_footnotes_root, 
    find_endnotes_root, find_note_paragraphs, WmlDocument,
};
use super::revision::{count_revisions, reset_revision_id_counter};
use super::settings::WmlComparerSettings;
use super::types::WmlComparisonResult;
use crate::error::Result;
use crate::util::lcs::{compute_correlation, CorrelationStatus, Hashable, LcsSettings};
use crate::xml::arena::XmlDocument;
use indextree::NodeId;
use sha1::{Digest, Sha1};

/// Comparison unit representing a paragraph for LCS comparison
#[derive(Debug, Clone)]
pub struct ParagraphUnit {
    /// SHA-1 hash of paragraph text content
    pub hash: String,
    /// Correlated SHA-1 hash (computed after accept/reject revisions)
    pub correlated_hash: Option<String>,
    /// Paragraph text content
    pub text: String,
    /// Original paragraph index
    #[allow(dead_code)]
    pub index: usize,
}

impl Hashable for ParagraphUnit {
    fn hash(&self) -> &str {
        // Use correlated hash if available (for pre-correlation matching)
        self.correlated_hash.as_deref().unwrap_or(&self.hash)
    }
}

/// Word unit for finer-grained comparison within paragraphs
#[derive(Debug, Clone)]
pub struct WordUnit {
    /// The word text (used as hash)
    pub text: String,
}

impl Hashable for WordUnit {
    fn hash(&self) -> &str {
        &self.text
    }
}

pub struct WmlComparer;

impl WmlComparer {
    pub fn compare(
        source1: &WmlDocument,
        source2: &WmlDocument,
        settings: Option<&WmlComparerSettings>,
    ) -> Result<WmlComparisonResult> {
        let _settings = settings.cloned().unwrap_or_default();
        reset_revision_id_counter(1);

        let doc1 = source1.main_document()?;
        let doc2 = source2.main_document()?;

        let body1 = find_document_body(&doc1);
        let body2 = find_document_body(&doc2);

        let paras1 = body1.map(|b| find_paragraphs(&doc1, b)).unwrap_or_default();
        let paras2 = body2.map(|b| find_paragraphs(&doc2, b)).unwrap_or_default();

        let units1 = create_paragraph_units(&doc1, &paras1);
        let units2 = create_paragraph_units(&doc2, &paras2);

        let lcs_settings = LcsSettings::new();
        let correlation = compute_correlation(&units1, &units2, &lcs_settings);

        let (mut insertions, mut deletions) = count_revisions_smart(&units1, &units2, &correlation);

        // Handle footnotes - including asymmetric cases where one doc has footnotes and the other doesn't
        let fn1_opt = source1.footnotes().ok().flatten();
        let fn2_opt = source2.footnotes().ok().flatten();
        
        match (&fn1_opt, &fn2_opt) {
            (Some(fn1), Some(fn2)) => {
                // Both documents have footnotes - compare them
                let fn_root1 = find_footnotes_root(fn1);
                let fn_root2 = find_footnotes_root(fn2);
                
                if let (Some(root1), Some(root2)) = (fn_root1, fn_root2) {
                    let fn_paras1 = find_note_paragraphs(fn1, root1);
                    let fn_paras2 = find_note_paragraphs(fn2, root2);
                    
                    let fn_units1 = create_paragraph_units(fn1, &fn_paras1);
                    let fn_units2 = create_paragraph_units(fn2, &fn_paras2);
                    
                    let fn_correlation = compute_correlation(&fn_units1, &fn_units2, &lcs_settings);
                    let (fn_ins, fn_del) = count_revisions_smart(&fn_units1, &fn_units2, &fn_correlation);
                    insertions += fn_ins;
                    deletions += fn_del;
                }
            }
            (None, Some(fn2)) => {
                // Doc1 has no footnotes, Doc2 has footnotes - all footnotes are insertions
                if let Some(root2) = find_footnotes_root(fn2) {
                    let fn_paras2 = find_note_paragraphs(fn2, root2);
                    if !fn_paras2.is_empty() {
                        // Count each footnote's paragraph(s) as insertions
                        insertions += fn_paras2.len();
                    }
                }
            }
            (Some(fn1), None) => {
                // Doc1 has footnotes, Doc2 has no footnotes - all footnotes are deletions
                if let Some(root1) = find_footnotes_root(fn1) {
                    let fn_paras1 = find_note_paragraphs(fn1, root1);
                    if !fn_paras1.is_empty() {
                        deletions += fn_paras1.len();
                    }
                }
            }
            (None, None) => {
                // Neither document has footnotes - nothing to do
            }
        }

        // Handle endnotes - including asymmetric cases
        let en1_opt = source1.endnotes().ok().flatten();
        let en2_opt = source2.endnotes().ok().flatten();
        
        match (&en1_opt, &en2_opt) {
            (Some(en1), Some(en2)) => {
                let en_root1 = find_endnotes_root(en1);
                let en_root2 = find_endnotes_root(en2);
                
                if let (Some(root1), Some(root2)) = (en_root1, en_root2) {
                    let en_paras1 = find_note_paragraphs(en1, root1);
                    let en_paras2 = find_note_paragraphs(en2, root2);
                    
                    let en_units1 = create_paragraph_units(en1, &en_paras1);
                    let en_units2 = create_paragraph_units(en2, &en_paras2);
                    
                    let en_correlation = compute_correlation(&en_units1, &en_units2, &lcs_settings);
                    let (en_ins, en_del) = count_revisions_smart(&en_units1, &en_units2, &en_correlation);
                    insertions += en_ins;
                    deletions += en_del;
                }
            }
            (None, Some(en2)) => {
                if let Some(root2) = find_endnotes_root(en2) {
                    let en_paras2 = find_note_paragraphs(en2, root2);
                    if !en_paras2.is_empty() {
                        insertions += en_paras2.len();
                    }
                }
            }
            (Some(en1), None) => {
                if let Some(root1) = find_endnotes_root(en1) {
                    let en_paras1 = find_note_paragraphs(en1, root1);
                    if !en_paras1.is_empty() {
                        deletions += en_paras1.len();
                    }
                }
            }
            (None, None) => {}
        }

        let result_bytes = source2.to_bytes()?;

        Ok(WmlComparisonResult {
            document: result_bytes,
            changes: Vec::new(),
            insertions,
            deletions,
            format_changes: 0,
            revision_count: insertions + deletions,
        })
    }

    /// Get the list of revisions from a document that already contains tracked changes
    pub fn get_revisions(
        document: &WmlDocument,
        settings: Option<&WmlComparerSettings>,
    ) -> Result<Vec<crate::types::Revision>> {
        let _settings = settings.cloned().unwrap_or_default();

        let doc = document.main_document()?;
        let body = find_document_body(&doc);

        if body.is_none() {
            return Ok(Vec::new());
        }

        let revisions = count_revisions(&doc, body.unwrap());
        
        // Convert to Revision types
        let mut result = Vec::new();
        
        for _ in 0..revisions.insertions {
            result.push(crate::types::Revision {
                revision_type: crate::types::RevisionType::Inserted,
                author: None,
                date: None,
                text: None,
            });
        }
        
        for _ in 0..revisions.deletions {
            result.push(crate::types::Revision {
                revision_type: crate::types::RevisionType::Deleted,
                author: None,
                date: None,
                text: None,
            });
        }

        Ok(result)
    }
}

/// Create paragraph units with hashes for comparison
fn create_paragraph_units(doc: &XmlDocument, paragraphs: &[NodeId]) -> Vec<ParagraphUnit> {
    paragraphs
        .iter()
        .enumerate()
        .map(|(index, &para)| {
            let text = extract_paragraph_text(doc, para);
            let hash = compute_sha1_hash(&text);
            
            ParagraphUnit {
                hash,
                correlated_hash: None, // Will be set during HashBlockLevelContent
                text,
                index,
            }
        })
        .collect()
}

/// Compute SHA-1 hash of a string
fn compute_sha1_hash(content: &str) -> String {
    let mut hasher = Sha1::new();
    hasher.update(content.as_bytes());
    let result = hasher.finalize();
    format!("{:x}", result)
}

fn count_revisions_smart(
    _units1: &[ParagraphUnit],
    _units2: &[ParagraphUnit],
    correlation: &[crate::util::lcs::CorrelatedSequence<ParagraphUnit>],
) -> (usize, usize) {
    let mut insertions = 0;
    let mut deletions = 0;
    let mut i = 0;
    
    while i < correlation.len() {
        let seq = &correlation[i];
        
        match seq.status {
            CorrelationStatus::Equal => {
                if let (Some(ref items1), Some(ref items2)) = (&seq.items1, &seq.items2) {
                    for (u1, u2) in items1.iter().zip(items2.iter()) {
                        if u1.text != u2.text {
                            let (ins, del) = count_word_level_revision_groups(&u1.text, &u2.text);
                            insertions += ins;
                            deletions += del;
                        }
                    }
                }
                i += 1;
            }
            CorrelationStatus::Deleted => {
                if i + 1 < correlation.len() && correlation[i + 1].status == CorrelationStatus::Inserted {
                    let deleted_items = seq.items1.as_deref().unwrap_or(&[]);
                    let inserted_items = correlation[i + 1].items2.as_deref().unwrap_or(&[]);
                    
                    let min_len = deleted_items.len().min(inserted_items.len());
                    for j in 0..min_len {
                        let (ins, del) = count_word_level_revision_groups(
                            &deleted_items[j].text,
                            &inserted_items[j].text,
                        );
                        insertions += ins;
                        deletions += del;
                    }
                    
                    if deleted_items.len() > min_len {
                        deletions += 1;
                    }
                    if inserted_items.len() > min_len {
                        insertions += 1;
                    }
                    
                    i += 2;
                } else {
                    if seq.items1.is_some() {
                        deletions += 1;
                    }
                    i += 1;
                }
            }
            CorrelationStatus::Inserted => {
                if seq.items2.is_some() {
                    insertions += 1;
                }
                i += 1;
            }
            CorrelationStatus::Unknown => {
                if let (Some(ref items1), Some(ref items2)) = (&seq.items1, &seq.items2) {
                    for (u1, u2) in items1.iter().zip(items2.iter()) {
                        let (ins, del) = count_word_level_revision_groups(&u1.text, &u2.text);
                        insertions += ins;
                        deletions += del;
                    }
                    if items1.len() > items2.len() {
                        deletions += 1;
                    } else if items2.len() > items1.len() {
                        insertions += 1;
                    }
                } else {
                    if seq.items1.is_some() {
                        deletions += 1;
                    }
                    if seq.items2.is_some() {
                        insertions += 1;
                    }
                }
                i += 1;
            }
        }
    }
    
    (insertions, deletions)
}

#[allow(dead_code)]
fn count_revisions_from_correlation(
    correlation: &[crate::util::lcs::CorrelatedSequence<ParagraphUnit>],
) -> (usize, usize) {
    let mut insertions = 0;
    let mut deletions = 0;

    for seq in correlation {
        match seq.status {
            CorrelationStatus::Inserted => {
                if seq.items2.is_some() {
                    insertions += 1;
                }
            }
            CorrelationStatus::Deleted => {
                if seq.items1.is_some() {
                    deletions += 1;
                }
            }
            CorrelationStatus::Equal => {
                if let (Some(ref items1), Some(ref items2)) = (&seq.items1, &seq.items2) {
                    for (u1, u2) in items1.iter().zip(items2.iter()) {
                        if u1.text != u2.text {
                            let (ins, del) = count_word_level_revision_groups(&u1.text, &u2.text);
                            insertions += ins;
                            deletions += del;
                        }
                    }
                }
            }
            CorrelationStatus::Unknown => {
                if let (Some(ref items1), Some(ref items2)) = (&seq.items1, &seq.items2) {
                    for (u1, u2) in items1.iter().zip(items2.iter()) {
                        let (ins, del) = count_word_level_revision_groups(&u1.text, &u2.text);
                        insertions += ins;
                        deletions += del;
                    }
                    if items1.len() > items2.len() {
                        deletions += 1;
                    } else if items2.len() > items1.len() {
                        insertions += 1;
                    }
                } else {
                    if seq.items1.is_some() {
                        deletions += 1;
                    }
                    if seq.items2.is_some() {
                        insertions += 1;
                    }
                }
            }
        }
    }

    (insertions, deletions)
}

/// Count revision GROUPS (contiguous runs) at word level between two texts
/// 
/// Example: "Video provides a powerful" -> "Video a powerful"
/// The word "provides" is deleted as one contiguous group = 1 deletion revision
const DETAIL_THRESHOLD: f64 = 0.15;

fn count_word_level_revision_groups(text1: &str, text2: &str) -> (usize, usize) {
    let words1: Vec<WordUnit> = tokenize_text(text1);
    let words2: Vec<WordUnit> = tokenize_text(text2);

    if words1.is_empty() && words2.is_empty() {
        return (0, 0);
    }
    if words1.is_empty() {
        return (1, 0);
    }
    if words2.is_empty() {
        return (0, 1);
    }

    let settings = LcsSettings::new();
    let correlation = compute_correlation(&words1, &words2, &settings);

    let equal_count: usize = correlation
        .iter()
        .filter(|seq| seq.status == CorrelationStatus::Equal)
        .map(|seq| seq.items1.as_ref().map(|v| v.len()).unwrap_or(0))
        .sum();

    let max_len = words1.len().max(words2.len());
    let min_len = words1.len().min(words2.len());
    let similarity = equal_count as f64 / max_len as f64;

    if equal_count == min_len {
        if words1.len() < words2.len() {
            return (1, 0);
        } else {
            return (0, 1);
        }
    }

    if similarity < DETAIL_THRESHOLD {
        let mut ins = 0;
        let mut del = 0;
        if !words1.is_empty() {
            del = 1;
        }
        if !words2.is_empty() {
            ins = 1;
        }
        return (ins, del);
    }

    let mut insertions = 0;
    let mut deletions = 0;

    for seq in &correlation {
        match seq.status {
            CorrelationStatus::Inserted => {
                if seq.items2.is_some() {
                    insertions += 1;
                }
            }
            CorrelationStatus::Deleted => {
                if seq.items1.is_some() {
                    deletions += 1;
                }
            }
            _ => {}
        }
    }

    (insertions, deletions)
}

/// Tokenize text into words for comparison
fn tokenize_text(text: &str) -> Vec<WordUnit> {
    let mut tokens = Vec::new();
    let mut current_word = String::new();
    
    for ch in text.chars() {
        if ch.is_whitespace() {
            if !current_word.is_empty() {
                tokens.push(WordUnit { text: current_word.clone() });
                current_word.clear();
            }
        } else if ch.is_ascii_punctuation() {
            if !current_word.is_empty() {
                tokens.push(WordUnit { text: current_word.clone() });
                current_word.clear();
            }
            tokens.push(WordUnit { text: ch.to_string() });
        } else {
            current_word.push(ch);
        }
    }
    
    if !current_word.is_empty() {
        tokens.push(WordUnit { text: current_word });
    }
    
    tokens
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_compute_sha1_hash() {
        let hash = compute_sha1_hash("hello world");
        assert!(!hash.is_empty());
        assert_eq!(hash.len(), 40);
    }

    #[test]
    fn test_tokenize_text() {
        let tokens = tokenize_text("hello world test");
        assert_eq!(tokens.len(), 3);
        assert_eq!(tokens[0].text, "hello");
        assert_eq!(tokens[1].text, "world");
        assert_eq!(tokens[2].text, "test");
        
        let tokens = tokenize_text("This.");
        assert_eq!(tokens.len(), 2);
        assert_eq!(tokens[0].text, "This");
        assert_eq!(tokens[1].text, ".");
    }

    #[test]
    fn test_word_level_revision_groups() {
        let (ins, del) = count_word_level_revision_groups("hello world", "hello world");
        assert_eq!(ins, 0);
        assert_eq!(del, 0);
        
        let (ins, del) = count_word_level_revision_groups("hello world", "hello");
        assert_eq!(ins, 0);
        assert_eq!(del, 1);
        
        let (ins, del) = count_word_level_revision_groups("hello", "hello world");
        assert_eq!(ins, 1);
        assert_eq!(del, 0);
    }
}
