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

use super::atom_list::{assign_unid_to_all_elements, create_comparison_unit_atom_list};
use super::coalesce::{coalesce, mark_content_as_deleted_or_inserted, coalesce_adjacent_runs};
use super::comparison_unit::{get_comparison_unit_list, WordSeparatorSettings, ComparisonUnitAtom, ComparisonCorrelationStatus, ContentElement};
use super::document::{
    extract_paragraph_text, find_document_body, find_paragraphs, find_footnotes_root, 
    find_endnotes_root, find_note_paragraphs, find_note_by_id, WmlDocument,
};
use super::lcs_algorithm::{self, lcs, flatten_to_atoms, CorrelatedSequence};
use super::preprocess::{preprocess_markup, PreProcessSettings};
use super::revision::{count_revisions, reset_revision_id_counter};
use super::revision_accepter::accept_revisions;
use super::settings::WmlComparerSettings;
use super::types::WmlComparisonResult;
use crate::error::Result;
use crate::util::lcs::{self, compute_correlation, Hashable, LcsSettings};
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::W;
use crate::xml::xname::{XAttribute, XName};
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



pub struct WmlComparer;

impl WmlComparer {
    /// Port of C# WmlComparer.Compare() and CompareInternal() (lines 141-291)
    pub fn compare(
        source1: &WmlDocument,
        source2: &WmlDocument,
        settings: Option<&WmlComparerSettings>,
    ) -> Result<WmlComparisonResult> {
        let settings = settings.cloned().unwrap_or_default();
        reset_revision_id_counter(1);

        let mut doc1 = source1.main_document()?;
        let mut doc2 = source2.main_document()?;
        
        let body1 = find_document_body(&doc1).ok_or_else(|| crate::error::RedlineError::ComparisonFailed("No body in document 1".to_string()))?;
        let body2 = find_document_body(&doc2).ok_or_else(|| crate::error::RedlineError::ComparisonFailed("No body in document 2".to_string()))?;
        
        let preprocess_settings = PreProcessSettings::for_comparison();
        preprocess_markup(&mut doc1, body1, &preprocess_settings)
            .map_err(|e| crate::error::RedlineError::ComparisonFailed(e))?;
        preprocess_markup(&mut doc2, body2, &preprocess_settings)
            .map_err(|e| crate::error::RedlineError::ComparisonFailed(e))?;

        // C# WmlComparer.cs:255-256 - Accept revisions before comparison
        // This ensures documents with tracked changes are compared by their final content
        let mut doc1 = accept_revisions(&doc1, body1);
        let mut doc2 = accept_revisions(&doc2, body2);
        
        // Find body nodes in the accepted documents
        let body1 = find_document_body(&doc1).ok_or_else(|| crate::error::RedlineError::ComparisonFailed("No body in accepted document 1".to_string()))?;
        let body2 = find_document_body(&doc2).ok_or_else(|| crate::error::RedlineError::ComparisonFailed("No body in accepted document 2".to_string()))?;

        let atoms1 = create_comparison_unit_atom_list(&mut doc1, body1, "main", &settings);
        let atoms2 = create_comparison_unit_atom_list(&mut doc2, body2, "main", &settings);
        
        if atoms1.is_empty() && atoms2.is_empty() {
            let result_bytes = source2.to_bytes()?;
            return Ok(WmlComparisonResult {
                document: result_bytes,
                changes: Vec::new(),
                insertions: 0,
                deletions: 0,
                format_changes: 0,
                revision_count: 0,
            });
        }

        let (mut insertions, mut deletions, mut format_changes, coalesce_result) = {
            let root_data = doc2.get(doc2.root().unwrap()).unwrap();
            let root_name = root_data.name().unwrap().clone();
            let root_attrs = root_data.attributes().unwrap_or(&[]).to_vec();
            compare_atoms_internal(atoms1, atoms2, root_name, root_attrs, &settings)?
        };

        // Create result document from source2
        let source2_bytes = source2.to_bytes()?;
        let mut result_doc = WmlDocument::from_bytes(&source2_bytes)?;
        result_doc.package_mut().put_xml_part("word/document.xml", &coalesce_result.document)?;

        // Process footnotes
        let footnotes1 = source1.footnotes()?;
        let footnotes2 = source2.footnotes()?;
        if footnotes1.is_some() || footnotes2.is_some() {
            let res = process_notes(footnotes1, footnotes2, "footnotes", &settings)?;
            insertions += res.insertions;
            deletions += res.deletions;
            format_changes += res.format_changes;
            result_doc.package_mut().put_xml_part("word/footnotes.xml", &res.document)?;
        }

        // Process endnotes
        let endnotes1 = source1.endnotes()?;
        let endnotes2 = source2.endnotes()?;
        if endnotes1.is_some() || endnotes2.is_some() {
            let res = process_notes(endnotes1, endnotes2, "endnotes", &settings)?;
            insertions += res.insertions;
            deletions += res.deletions;
            format_changes += res.format_changes;
            result_doc.package_mut().put_xml_part("word/endnotes.xml", &res.document)?;
        }

        let result_bytes = result_doc.to_bytes()?;

        Ok(WmlComparisonResult {
            document: result_bytes,
            changes: Vec::new(),
            insertions,
            deletions,
            format_changes,
            revision_count: insertions + deletions + format_changes,
        })
    }

    /// Legacy compare implementation for backward compatibility
    pub fn compare_legacy(
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
            let normalized = normalize_whitespace(&text);
            let hash = compute_sha1_hash(&normalized);
            
            ParagraphUnit {
                hash,
                correlated_hash: None,
                text: normalized,
                index,
            }
        })
        .collect()
}

fn normalize_whitespace(text: &str) -> String {
    text.chars()
        .map(|c| if c.is_whitespace() { ' ' } else { c })
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
            lcs::CorrelationStatus::Equal => {
                if let (Some(ref items1), Some(ref items2)) = (&seq.items1, &seq.items2) {
                    for (u1, u2) in items1.iter().zip(items2.iter()) {
                        if u1.text != u2.text {
                            insertions += 1;
                            deletions += 1;
                        }
                    }
                }
                i += 1;
            }
            lcs::CorrelationStatus::Deleted => {
                if i + 1 < correlation.len() && correlation[i + 1].status == lcs::CorrelationStatus::Inserted {
                    let deleted_items = seq.items1.as_deref().unwrap_or(&[]);
                    let inserted_items = correlation[i + 1].items2.as_deref().unwrap_or(&[]);
                    
                    let min_len = deleted_items.len().min(inserted_items.len());
                    for _j in 0..min_len {
                        insertions += 1;
                        deletions += 1;
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
            lcs::CorrelationStatus::Inserted => {
                if seq.items2.is_some() {
                    insertions += 1;
                }
                i += 1;
            }
            lcs::CorrelationStatus::Unknown => {
                if let (Some(ref items1), Some(ref items2)) = (&seq.items1, &seq.items2) {
                    for (_u1, _u2) in items1.iter().zip(items2.iter()) {
                        insertions += 1;
                        deletions += 1;
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

fn reconcile_formatting_changes(atoms: &mut [ComparisonUnitAtom], settings: &WmlComparerSettings) {
    if !settings.track_formatting_changes {
        return;
    }

    for atom in atoms {
        if let Some(ref before) = atom.comparison_unit_atom_before {
            if atom.correlation_status == ComparisonCorrelationStatus::Equal {
                if atom.formatting_signature != before.formatting_signature {
                    atom.correlation_status = ComparisonCorrelationStatus::FormatChanged;
                }
            }
        }
    }
}

fn count_revisions_from_atoms(atoms: &[ComparisonUnitAtom]) -> (usize, usize) {
    // C# GetRevisions uses GroupAdjacent to group atoms by:
    // 1. CorrelationStatus (Equal, Inserted, Deleted)
    // 2. For non-Equal status, also by RevTrackElement (serialized, minus id/Unid)
    //
    // This means all atoms within a contiguous deleted region count as ONE deletion,
    // regardless of paragraph boundaries. Only count when status changes.
    
    let mut insertions = 0;
    let mut deletions = 0;
    let mut last_status = ComparisonCorrelationStatus::Equal;

    for atom in atoms {
        // Only count a new revision when status actually changes
        if atom.correlation_status != last_status {
            match atom.correlation_status {
                ComparisonCorrelationStatus::Inserted => {
                    insertions += 1;
                }
                ComparisonCorrelationStatus::Deleted => {
                    deletions += 1;
                }
                ComparisonCorrelationStatus::FormatChanged => {
                    // Format changes are counted separately from XML after coalesce
                }
                ComparisonCorrelationStatus::Equal
                | ComparisonCorrelationStatus::Nil
                | ComparisonCorrelationStatus::Normal
                | ComparisonCorrelationStatus::Unknown
                | ComparisonCorrelationStatus::Group => {
                    // These don't contribute to revision counts
                }
            }
        }
        last_status = atom.correlation_status;
    }

    (insertions, deletions)
}

fn assemble_ancestor_unids(atoms: &mut [ComparisonUnitAtom]) {
    // 1. Normalize UNIDs for Equal/FormatChanged atoms (C# lines 3450-3485)
    for atom in atoms.iter_mut() {
        if atom.correlation_status == ComparisonCorrelationStatus::Equal || 
           atom.correlation_status == ComparisonCorrelationStatus::FormatChanged {
            if let Some(ref before) = atom.comparison_unit_atom_before {
                if atom.ancestor_elements.len() == before.ancestor_elements.len() {
                    for i in 0..atom.ancestor_elements.len() {
                        atom.ancestor_elements[i].unid = before.ancestor_elements[i].unid.clone();
                    }
                }
            }
        }
    }

    // 2. Propagate paragraph UNIDs (C# lines 3515-3550)
    // We process in reverse order as C# does
    let mut current_ancestor_unids: Vec<String> = Vec::new();
    
    for atom in atoms.iter_mut().rev() {
        if matches!(atom.content_element, ContentElement::ParagraphProperties) {
            current_ancestor_unids = atom.ancestor_elements.iter().map(|a| a.unid.clone()).collect();
            atom.ancestor_unids = current_ancestor_unids.clone();
        } else {
            if current_ancestor_unids.is_empty() {
                atom.ancestor_unids = atom.ancestor_elements.iter().map(|a| a.unid.clone()).collect();
            } else {
                let mut unids = current_ancestor_unids.clone();
                // Ensure we don't truncate deeper hierarchies (like tables/textboxes)
                for (i, ancestor) in atom.ancestor_elements.iter().enumerate() {
                    if i >= unids.len() {
                        unids.push(ancestor.unid.clone());
                    }
                }
                atom.ancestor_unids = unids;
            }
        }
    }
}

fn compare_atoms_internal(
    atoms1: Vec<ComparisonUnitAtom>,
    atoms2: Vec<ComparisonUnitAtom>,
    root_name: XName,
    root_attrs: Vec<XAttribute>,
    settings: &WmlComparerSettings,
) -> Result<(usize, usize, usize, super::coalesce::CoalesceResult)> {
    let word_settings = WordSeparatorSettings::default();
    let units1 = get_comparison_unit_list(atoms1, &word_settings);
    let units2 = get_comparison_unit_list(atoms2, &word_settings);

    let correlated = lcs(units1, units2, settings);
    
    let mut flattened_atoms = flatten_to_atoms(&correlated);
    assemble_ancestor_unids(&mut flattened_atoms);
    
    if settings.track_formatting_changes {
        reconcile_formatting_changes(&mut flattened_atoms, settings);
    }

    // Count revisions from atom list (C# GetRevisions algorithm)
    // This groups adjacent atoms by correlation status, which is how C# counts
    let (insertions, deletions) = count_revisions_from_atoms(&flattened_atoms);
    
    let mut coalesce_result = coalesce(&flattened_atoms, settings, root_name, root_attrs);
    
    // Wrap content in revision marks (C# line 2173)
    mark_content_as_deleted_or_inserted(&mut coalesce_result.document, coalesce_result.root, settings);
    
    // Consolidate adjacent revisions (C# line 2174)
    coalesce_adjacent_runs(&mut coalesce_result.document, coalesce_result.root, &settings);
    
    // Format changes are counted from XML as they're added during mark_content_as_deleted_or_inserted
    let format_changes = count_revisions(&coalesce_result.document, coalesce_result.root).format_changes;
    
    Ok((insertions, deletions, format_changes, coalesce_result))
}

struct NoteProcessingResult {
    insertions: usize,
    deletions: usize,
    format_changes: usize,
    document: XmlDocument,
}

fn process_notes(
    source1_opt: Option<XmlDocument>,
    source2_opt: Option<XmlDocument>,
    part_type: &str,
    settings: &WmlComparerSettings,
) -> Result<NoteProcessingResult> {
    let mut total_ins = 0;
    let mut total_del = 0;
    let mut total_fmt = 0;

    // Use source2 as the base for the result document if it exists, else source1
    let mut result_doc = match &source2_opt {
        Some(doc) => {
            // Deep clone doc
            let xml = crate::xml::builder::serialize(doc)?;
            crate::xml::parser::parse(&xml)?
        }
        None => {
            let xml = crate::xml::builder::serialize(source1_opt.as_ref().unwrap())?;
            crate::xml::parser::parse(&xml)?
        }
    };

    let root1_opt = source1_opt.as_ref().and_then(|doc| {
        if part_type == "footnotes" { find_footnotes_root(doc) } else { find_endnotes_root(doc) }
    });
    let root2_opt = source2_opt.as_ref().and_then(|doc| {
        if part_type == "footnotes" { find_footnotes_root(doc) } else { find_endnotes_root(doc) }
    });

    match (source1_opt, source2_opt, root1_opt, root2_opt) {
        (Some(mut doc1), Some(mut doc2), Some(r1), Some(r2)) => {
            // Both have notes - compare them
            let (ins, del, fmt, coalesce_res) = compare_part_content(&mut doc1, r1, &mut doc2, r2, part_type, settings)?;
            total_ins = ins;
            total_del = del;
            total_fmt = fmt;
            result_doc = coalesce_res.document;
        }
        (None, Some(_doc2), None, Some(r2)) => {
            // Only doc2 has notes - all are insertions
            // Note: We need to mark them as inserted!
            // For now, just count them
            let paras = find_note_paragraphs(&result_doc, r2);
            total_ins = paras.len();
        }
        (Some(_doc1), None, Some(r1), None) => {
            // Only doc1 has notes - all are deletions
            let paras = find_note_paragraphs(&result_doc, r1);
            total_del = paras.len();
        }
        _ => {}
    }

    Ok(NoteProcessingResult {
        insertions: total_ins,
        deletions: total_del,
        format_changes: total_fmt,
        document: result_doc,
    })
}

fn compare_part_content(
    doc1: &mut XmlDocument,
    root1: NodeId,
    doc2: &mut XmlDocument,
    root2: NodeId,
    part_name: &str,
    settings: &WmlComparerSettings,
) -> Result<(usize, usize, usize, super::coalesce::CoalesceResult)> {
    let atoms1 = create_comparison_unit_atom_list(doc1, root1, part_name, settings);
    let atoms2 = create_comparison_unit_atom_list(doc2, root2, part_name, settings);
    
    let root_data = doc2.get(root2).unwrap();
    let root_name = root_data.name().unwrap().clone();
    let root_attrs = root_data.attributes().unwrap_or(&[]).to_vec();
    
    compare_atoms_internal(atoms1, atoms2, root_name, root_attrs, settings)
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


}