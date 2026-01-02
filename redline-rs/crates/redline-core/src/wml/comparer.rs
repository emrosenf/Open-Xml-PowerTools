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

use super::atom_list::create_comparison_unit_atom_list;
use super::change_event::detect_format_changes;
use super::coalesce::{
    coalesce, coalesce_adjacent_runs, mark_content_as_deleted_or_inserted, strip_pt_attributes,
};
use super::comments::{add_comments_to_package, extract_comments_data, merge_comments};
use super::comparison_unit::{
    get_comparison_unit_list, ComparisonCorrelationStatus, ComparisonUnitAtom, ContentElement,
    WordSeparatorSettings,
};
use super::document::{
    extract_paragraph_text, find_document_body, find_endnotes_root, find_footnotes_root,
    find_note_by_id, find_note_paragraphs, find_paragraphs, WmlDocument,
};
use super::extract_changes::extract_changes_from_document;
use super::lcs_algorithm::{flatten_to_atoms, lcs};
#[cfg(feature = "trace")]
use super::lcs_algorithm::{generate_focused_trace, units_match_filter};
use super::preprocess::{
    preprocess_markup, repair_unids_after_revision_acceptance, PreProcessSettings,
};
use super::revision::{count_revisions, reset_revision_id_counter};
use super::revision_accepter::accept_revisions;
use super::settings::WmlComparerSettings;
use super::types::WmlComparisonResult;
use crate::error::{RedlineError, Result};
use crate::util::lcs::{self, compute_correlation, Hashable, LcsSettings};
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::W;
use crate::xml::node::XmlNodeData;
use crate::xml::xname::{XAttribute, XName};
use chrono::Utc;
use indextree::NodeId;
use sha1::{Digest, Sha1};
use std::collections::HashMap;

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

/// Filter out xmlns attributes from node data
fn filter_xmlns_attrs(data: &XmlNodeData) -> XmlNodeData {
    match data {
        XmlNodeData::Element { name, attributes } => {
            let filtered_attrs: Vec<XAttribute> = attributes
                .iter()
                .filter(|attr| {
                    // Keep attribute if it's NOT an xmlns declaration
                    // (xmlns without prefix or xmlns:prefix)
                    let is_xmlns = (attr.name.namespace.is_none()
                        && attr.name.local_name == "xmlns")
                        || attr.name.namespace.as_deref() == Some("http://www.w3.org/2000/xmlns/")
                        || attr.name.local_name.starts_with("xmlns");
                    !is_xmlns
                })
                .cloned()
                .collect();
            XmlNodeData::Element {
                name: name.clone(),
                attributes: filtered_attrs,
            }
        }
        other => other.clone(),
    }
}

/// Helper to clone sectPr children that we want to preserve
fn clone_sect_pr_children(doc: &XmlDocument, sect_pr: NodeId) -> Vec<XmlNodeData> {
    let mut children = Vec::new();
    for sect_child in doc.children(sect_pr) {
        if let Some(child_data) = doc.get(sect_child) {
            if let Some(child_name) = child_data.name() {
                // Copy essential sectPr elements including header/footer references
                let local = &child_name.local_name;
                if local == "type"
                    || local == "pgSz"
                    || local == "pgMar"
                    || local == "cols"
                    || local == "titlePg"
                    || local == "noEndnote"
                    || local == "docGrid"
                    || local == "bidi"
                    || local == "headerReference"
                    || local == "footerReference"
                    || local == "endnotePr"
                {
                    // Filter xmlns from child elements too
                    children.push(filter_xmlns_attrs(child_data));
                }
            }
        }
    }
    children
}

/// Helper to clone sectPr attributes, filtering out xmlns declarations
fn clone_sect_pr_attrs(doc: &XmlDocument, sect_pr: NodeId) -> Vec<XAttribute> {
    doc.get(sect_pr)
        .and_then(|d| d.attributes())
        .map(|a| {
            a.iter()
                .filter(|attr| {
                    let is_xmlns = (attr.name.namespace.is_none()
                        && attr.name.local_name == "xmlns")
                        || attr.name.namespace.as_deref() == Some("http://www.w3.org/2000/xmlns/")
                        || attr.name.local_name.starts_with("xmlns");
                    !is_xmlns
                })
                .cloned()
                .collect()
        })
        .unwrap_or_default()
}

/// Extract sectPr (section properties) from document body
/// C# WmlComparer.cs lines 2030-2035: saves sectPr before comparison
fn extract_sect_pr(doc: &XmlDocument, body: NodeId) -> Option<(Vec<XAttribute>, Vec<XmlNodeData>)> {
    // Find sectPr that is a direct child of body
    for child in doc.children(body) {
        if let Some(data) = doc.get(child) {
            if let Some(name) = data.name() {
                if name == &W::sectPr() {
                    let attrs = clone_sect_pr_attrs(doc, child);
                    let children = clone_sect_pr_children(doc, child);
                    return Some((attrs, children));
                }
            }
        }
    }
    None
}

/// Extract sectPr from the last paragraph's pPr (contains headerReference/footerReference)
/// This is where Word stores the main section's header/footer references
fn extract_ppr_sect_pr(
    doc: &XmlDocument,
    body: NodeId,
) -> Option<(Vec<XAttribute>, Vec<XmlNodeData>)> {
    // Find the last paragraph that has a sectPr inside its pPr
    let mut last_ppr_sect_pr: Option<NodeId> = None;

    for child in doc.children(body) {
        if let Some(data) = doc.get(child) {
            if let Some(name) = data.name() {
                if name == &W::p() {
                    // Look for pPr child
                    for p_child in doc.children(child) {
                        if let Some(p_child_data) = doc.get(p_child) {
                            if let Some(p_child_name) = p_child_data.name() {
                                if p_child_name == &W::pPr() {
                                    // Look for sectPr inside pPr
                                    for ppr_child in doc.children(p_child) {
                                        if let Some(ppr_child_data) = doc.get(ppr_child) {
                                            if let Some(ppr_child_name) = ppr_child_data.name() {
                                                if ppr_child_name == &W::sectPr() {
                                                    last_ppr_sect_pr = Some(ppr_child);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    last_ppr_sect_pr.map(|sect_pr| {
        let attrs = clone_sect_pr_attrs(doc, sect_pr);
        let children = clone_sect_pr_children(doc, sect_pr);
        (attrs, children)
    })
}

/// Add sectPr to result document body
/// C# WmlComparer.cs lines 2201-2216: removes existing sectPr and adds saved one back
fn add_sect_pr_to_document(
    doc: &mut XmlDocument,
    saved_sect_pr: Option<(Vec<XAttribute>, Vec<XmlNodeData>)>,
) {
    let body = {
        let root = match doc.root() {
            Some(r) => r,
            None => return,
        };
        doc.children(root).find(|&child| {
            doc.get(child)
                .and_then(|d| d.name())
                .map(|n| n == &W::body())
                .unwrap_or(false)
        })
    };
    let body = match body {
        Some(b) => b,
        None => return,
    };

    // Remove any existing sectPr elements from body (C# line 2201)
    let sect_prs_to_remove: Vec<NodeId> = doc
        .children(body)
        .filter(|&child| {
            doc.get(child)
                .and_then(|d| d.name())
                .map(|n| n == &W::sectPr())
                .unwrap_or(false)
        })
        .collect();
    for node in sect_prs_to_remove {
        doc.detach(node);
    }

    // Add saved sectPr if present (C# lines 2204-2216)
    if let Some((attrs, children)) = saved_sect_pr {
        let sect_pr = doc.new_node(XmlNodeData::element_with_attrs(W::sectPr(), attrs));
        for child_data in children {
            doc.add_child(sect_pr, child_data);
        }
        doc.reparent(body, sect_pr);
    }
}

/// Find a relationship ID by target filename in a relationships XML
fn find_relationship_id_by_target(rels_content: &[u8], target: &str) -> Option<String> {
    let content = std::str::from_utf8(rels_content).ok()?;
    for line in content.split("<Relationship") {
        if line.contains(&format!("Target=\"{}\"", target)) {
            // Extract Id attribute
            if let Some(id_start) = line.find("Id=\"") {
                let start = id_start + 4;
                let rest = &line[start..];
                if let Some(end) = rest.find('"') {
                    return Some(rest[..end].to_string());
                }
            }
        }
    }
    None
}

/// Fix headerReference and footerReference relationship IDs in the document
/// This is necessary because relationship IDs from source1 may not match source2's package
fn fix_header_footer_relationship_ids(
    doc: &mut XmlDocument,
    package: &crate::package::OoxmlPackage,
) {
    // Get the document.xml.rels content
    let rels_content = match package.get_part("word/_rels/document.xml.rels") {
        Some(c) => c,
        None => return,
    };

    // Find the correct rIds for header and footer
    let header_rid = find_relationship_id_by_target(rels_content, "header1.xml");
    let footer_rid = find_relationship_id_by_target(rels_content, "footer1.xml");

    // If we don't have the rIds, don't try to fix
    if header_rid.is_none() && footer_rid.is_none() {
        return;
    }

    // Find all sectPr elements and update their headerReference/footerReference
    let root = match doc.root() {
        Some(r) => r,
        None => return,
    };

    // Collect all nodes that need updating (to avoid borrow issues)
    let mut header_refs_to_update: Vec<NodeId> = Vec::new();
    let mut footer_refs_to_update: Vec<NodeId> = Vec::new();

    for node_id in doc.descendants(root) {
        if let Some(data) = doc.get(node_id) {
            if let Some(name) = data.name() {
                if name.local_name == "headerReference" {
                    header_refs_to_update.push(node_id);
                } else if name.local_name == "footerReference" {
                    footer_refs_to_update.push(node_id);
                }
            }
        }
    }

    // Update headerReference elements
    if let Some(ref rid) = header_rid {
        for node_id in header_refs_to_update {
            if let Some(data) = doc.get_mut(node_id) {
                if let XmlNodeData::Element { attributes, .. } = data {
                    // Find and update the r:id attribute
                    for attr in attributes.iter_mut() {
                        if attr.name.local_name == "id" &&
                           attr.name.namespace.as_deref() == Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships") {
                            attr.value = rid.clone();
                        }
                    }
                }
            }
        }
    }

    // Update footerReference elements
    if let Some(ref rid) = footer_rid {
        for node_id in footer_refs_to_update {
            if let Some(data) = doc.get_mut(node_id) {
                if let XmlNodeData::Element { attributes, .. } = data {
                    // Find and update the r:id attribute
                    for attr in attributes.iter_mut() {
                        if attr.name.local_name == "id" &&
                           attr.name.namespace.as_deref() == Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships") {
                            attr.value = rid.clone();
                        }
                    }
                }
            }
        }
    }
}

/// Add sectPr to the last paragraph's pPr (for headerReference/footerReference)
/// This ensures Word can find the headers/footers for the document
fn add_ppr_sect_pr_to_document(
    doc: &mut XmlDocument,
    saved_ppr_sect_pr: Option<(Vec<XAttribute>, Vec<XmlNodeData>)>,
) {
    let saved = match saved_ppr_sect_pr {
        Some(s) => s,
        None => return,
    };

    let body = {
        let root = match doc.root() {
            Some(r) => r,
            None => return,
        };
        doc.children(root).find(|&child| {
            doc.get(child)
                .and_then(|d| d.name())
                .map(|n| n == &W::body())
                .unwrap_or(false)
        })
    };
    let body = match body {
        Some(b) => b,
        None => return,
    };

    // Find the last paragraph
    let mut last_p: Option<NodeId> = None;
    for child in doc.children(body) {
        if let Some(data) = doc.get(child) {
            if let Some(name) = data.name() {
                if name == &W::p() {
                    last_p = Some(child);
                }
            }
        }
    }

    let last_p = match last_p {
        Some(p) => p,
        None => return,
    };

    // Find or create pPr in the last paragraph
    let mut ppr: Option<NodeId> = None;
    for child in doc.children(last_p) {
        if let Some(data) = doc.get(child) {
            if let Some(name) = data.name() {
                if name == &W::pPr() {
                    ppr = Some(child);
                    break;
                }
            }
        }
    }

    let ppr = match ppr {
        Some(p) => p,
        None => {
            // Create pPr if it doesn't exist
            let new_ppr = doc.new_node(XmlNodeData::element(W::pPr()));
            // Insert pPr as first child of paragraph
            let children: Vec<NodeId> = doc.children(last_p).collect();
            for child in &children {
                doc.detach(*child);
            }
            doc.reparent(last_p, new_ppr);
            for child in children {
                doc.reparent(last_p, child);
            }
            new_ppr
        }
    };

    // Remove any existing sectPr from pPr
    let sect_prs_to_remove: Vec<NodeId> = doc
        .children(ppr)
        .filter(|&child| {
            doc.get(child)
                .and_then(|d| d.name())
                .map(|n| n == &W::sectPr())
                .unwrap_or(false)
        })
        .collect();
    for node in sect_prs_to_remove {
        doc.detach(node);
    }

    // Add the saved sectPr to pPr
    let (attrs, children) = saved;
    let sect_pr = doc.new_node(XmlNodeData::element_with_attrs(W::sectPr(), attrs));
    for child_data in children {
        doc.add_child(sect_pr, child_data);
    }
    doc.reparent(ppr, sect_pr);
}

pub struct WmlComparer;

impl WmlComparer {
    /// Port of C# WmlComparer.Compare() and CompareInternal() (lines 141-291)
    pub fn compare(
        source1: &WmlDocument,
        source2: &WmlDocument,
        settings: Option<&WmlComparerSettings>,
    ) -> Result<WmlComparisonResult> {
        // Timing is only available on native platforms, not WASM
        #[cfg(not(target_arch = "wasm32"))]
        let timing_enabled = std::env::var("WML_TIMING").is_ok();
        #[cfg(not(target_arch = "wasm32"))]
        let t0 = std::time::Instant::now();

        let mut settings = settings.cloned().unwrap_or_default();

        // C# WmlComparer.cs lines 171-176: Extract author/date from source2 if not set
        // This matches MS Word behavior where revision metadata comes from the modified document
        Self::resolve_revision_metadata(&mut settings, source2);

        reset_revision_id_counter(1);

        let mut doc1 = source1.main_document()?;
        let mut doc2 = source2.main_document()?;

        let body1 = find_document_body(&doc1).ok_or_else(|| {
            crate::error::RedlineError::ComparisonFailed("No body in document 1".to_string())
        })?;
        let body2 = find_document_body(&doc2).ok_or_else(|| {
            crate::error::RedlineError::ComparisonFailed("No body in document 2".to_string())
        })?;

        // Save sectPr from doc1 before comparison (C# WmlComparer.cs lines 2030-2035)
        // sectPr defines page layout (margins, size, etc.) and must be preserved in output
        let saved_sect_pr = extract_sect_pr(&doc1, body1);
        // Also save the pPr-level sectPr which contains headerReference/footerReference
        // This is critical for Word to locate header/footer parts
        let saved_ppr_sect_pr = extract_ppr_sect_pr(&doc1, body1);

        let preprocess_settings = PreProcessSettings::for_comparison();
        preprocess_markup(&mut doc1, body1, &preprocess_settings)
            .map_err(|e| crate::error::RedlineError::ComparisonFailed(e))?;
        preprocess_markup(&mut doc2, body2, &preprocess_settings)
            .map_err(|e| crate::error::RedlineError::ComparisonFailed(e))?;
        #[cfg(not(target_arch = "wasm32"))]
        if timing_enabled {
            eprintln!("  preprocess: {:?}", t0.elapsed());
        }

        // C# WmlComparer.cs:255-256 - Accept revisions before comparison
        // This ensures documents with tracked changes are compared by their final content
        // IMPORTANT: Pass the document root, not just the body, to preserve full document structure
        let doc1_root = doc1.root().ok_or_else(|| {
            crate::error::RedlineError::ComparisonFailed("No root in document 1".to_string())
        })?;
        let doc2_root = doc2.root().ok_or_else(|| {
            crate::error::RedlineError::ComparisonFailed("No root in document 2".to_string())
        })?;

        let mut doc1 = accept_revisions(&doc1, doc1_root);
        let mut doc2 = accept_revisions(&doc2, doc2_root);

        // Find body nodes in the accepted documents
        let body1 = find_document_body(&doc1).ok_or_else(|| {
            crate::error::RedlineError::ComparisonFailed(
                "No body in accepted document 1".to_string(),
            )
        })?;
        let body2 = find_document_body(&doc2).ok_or_else(|| {
            crate::error::RedlineError::ComparisonFailed(
                "No body in accepted document 2".to_string(),
            )
        })?;

        // C# WmlComparer.cs:270-279 - Repair UNIDs after revision acceptance
        // The accept_revisions function may create new elements (e.g., by unwrapping w:ins)
        // that don't have pt:Unid attributes. Re-assign UNIDs to ensure all elements have them.
        // This is critical for proper run boundary preservation during coalesce grouping.
        repair_unids_after_revision_acceptance(&mut doc1, body1)
            .map_err(|e| crate::error::RedlineError::ComparisonFailed(e))?;
        repair_unids_after_revision_acceptance(&mut doc2, body2)
            .map_err(|e| crate::error::RedlineError::ComparisonFailed(e))?;
        #[cfg(not(target_arch = "wasm32"))]
        if timing_enabled {
            eprintln!("  accept_revisions: {:?}", t0.elapsed());
        }

        let atoms1 = create_comparison_unit_atom_list(&mut doc1, body1, "main", &settings);
        let atoms2 = create_comparison_unit_atom_list(&mut doc2, body2, "main", &settings);
        #[cfg(not(target_arch = "wasm32"))]
        if timing_enabled {
            eprintln!("  create_atoms: {:?}", t0.elapsed());
        }

        if atoms1.is_empty() && atoms2.is_empty() {
            let result_bytes = source2.to_bytes()?;
            return Ok(WmlComparisonResult {
                document: result_bytes,
                changes: Vec::new(),
                insertions: 0,
                deletions: 0,
                format_changes: 0,
                revision_count: 0,
                lcs_traces: None,
            });
        }

        let (
            mut insertions,
            mut deletions,
            mut format_changes,
            mut coalesce_result,
            correlated_atoms,
            lcs_traces,
        ) = {
            let root_id = doc2.root().ok_or_else(|| {
                crate::error::RedlineError::ComparisonFailed(
                    "No root in document 2 for atoms comparison".to_string(),
                )
            })?;
            let root_data = doc2.get(root_id).ok_or_else(|| {
                crate::error::RedlineError::ComparisonFailed("Failed to get root data".to_string())
            })?;
            let root_name = root_data
                .name()
                .ok_or_else(|| {
                    crate::error::RedlineError::ComparisonFailed("Root has no name".to_string())
                })?
                .clone();
            let root_attrs = root_data.attributes().unwrap_or(&[]).to_vec();
            compare_atoms_internal(atoms1, atoms2, root_name, root_attrs, &settings)?
        };
        #[cfg(not(target_arch = "wasm32"))]
        if timing_enabled {
            eprintln!("  compare_atoms: {:?}", t0.elapsed());
        }

        // Add sectPr back to result document (C# WmlComparer.cs lines 2201-2216)
        // This restores page layout that was saved from doc1
        add_sect_pr_to_document(&mut coalesce_result.document, saved_sect_pr);
        // Also add the pPr-level sectPr to the last paragraph (contains headerReference/footerReference)
        add_ppr_sect_pr_to_document(&mut coalesce_result.document, saved_ppr_sect_pr);

        // Order elements per OOXML standard (C# WmlComparer.cs line 2196)
        // This is critical - OOXML requires elements in specific order within pPr, rPr, etc.
        super::order_elements_per_standard(&mut coalesce_result.document);

        // Create result document from source2
        let source2_bytes = source2.to_bytes()?;
        let mut result_doc = WmlDocument::from_bytes(&source2_bytes)?;

        // Fix header/footer relationship IDs before putting the document
        // The sectPr we copied from source1 may have different rIds than source2's package
        fix_header_footer_relationship_ids(&mut coalesce_result.document, result_doc.package());

        result_doc
            .package_mut()
            .put_xml_part("word/document.xml", &coalesce_result.document)?;

        // Process footnotes - collect references from correlated atoms
        let footnote_refs = collect_note_references(&correlated_atoms, "footnoteReference");
        let footnotes1 = source1.footnotes()?;
        let footnotes2 = source2.footnotes()?;
        if !footnote_refs.is_empty() && (footnotes1.is_some() || footnotes2.is_some()) {
            let mut res = process_notes(
                footnotes1,
                footnotes2,
                "footnotes",
                &footnote_refs,
                &settings,
            )?;
            insertions += res.insertions;
            deletions += res.deletions;
            format_changes += res.format_changes;
            // Order elements per OOXML standard (ensures pPr comes first in paragraphs, etc.)
            super::order_elements_per_standard(&mut res.document);
            result_doc
                .package_mut()
                .put_xml_part("word/footnotes.xml", &res.document)?;
        }

        // Process endnotes - collect references from correlated atoms
        let endnote_refs = collect_note_references(&correlated_atoms, "endnoteReference");
        let endnotes1 = source1.endnotes()?;
        let endnotes2 = source2.endnotes()?;
        if !endnote_refs.is_empty() && (endnotes1.is_some() || endnotes2.is_some()) {
            let mut res =
                process_notes(endnotes1, endnotes2, "endnotes", &endnote_refs, &settings)?;
            insertions += res.insertions;
            deletions += res.deletions;
            format_changes += res.format_changes;
            // Order elements per OOXML standard (ensures pPr comes first in paragraphs, etc.)
            super::order_elements_per_standard(&mut res.document);
            result_doc
                .package_mut()
                .put_xml_part("word/endnotes.xml", &res.document)?;
        }

        // NOTE: Comments from source documents are intentionally NOT preserved.
        // WmlComparer (C#) throws away commentRangeStart/End markers during comparison
        // and doesn't preserve existing comments. Including comments.xml without
        // corresponding markers in document.xml creates an invalid OOXML package
        // that MS Word cannot open. To match C# behavior, we skip comment handling.
        //
        // Future enhancement: If we want to preserve comments, we need to also
        // preserve commentRangeStart, commentRangeEnd, and commentReference elements
        // in the document body by NOT throwing them away during atom list creation.

        // Extract structured change data from the result document's revision markup
        // This walks the document looking for w:ins, w:del, w:rPrChange elements
        let changes = {
            // Get the document.xml from the result package to extract changes
            let doc_xml = result_doc.main_document()?;
            if let Some(body) = find_document_body(&doc_xml) {
                extract_changes_from_document(
                    &doc_xml,
                    body,
                    settings.author_for_revisions.as_deref(),
                    settings.date_time_for_revisions.as_deref(),
                )
            } else {
                Vec::new()
            }
        };

        let result_bytes = result_doc.to_bytes()?;
        #[cfg(not(target_arch = "wasm32"))]
        if timing_enabled {
            eprintln!("  finalize: {:?}", t0.elapsed());
        }

        Ok(WmlComparisonResult {
            document: result_bytes,
            changes,
            insertions,
            deletions,
            format_changes,
            revision_count: insertions + deletions + format_changes,
            lcs_traces,
        })
    }

    /// Resolve author and date/time for revisions from the modified document's core properties.
    /// C# WmlComparer.cs lines 155-176: ExtractAuthorFromDocument and CompareInternal
    ///
    /// Priority for author:
    /// 1. Explicitly set in settings (author_for_revisions)
    /// 2. cp:lastModifiedBy from source2's docProps/core.xml
    /// 3. dc:creator from source2's docProps/core.xml
    /// 4. "Unknown" as fallback
    ///
    /// Priority for date:
    /// 1. Explicitly set in settings (date_time_for_revisions)
    /// 2. dcterms:modified from source2's docProps/core.xml
    /// 3. Current time as fallback
    fn resolve_revision_metadata(settings: &mut WmlComparerSettings, source2: &WmlDocument) {
        let core_props = source2.package().get_core_properties();

        // Resolve author if not explicitly set
        if settings.author_for_revisions.is_none() {
            settings.author_for_revisions = core_props
                .last_modified_by
                .or(core_props.creator)
                .or_else(|| Some("Unknown".to_string()));
        }

        // Resolve date if not explicitly set
        if settings.date_time_for_revisions.is_none() {
            settings.date_time_for_revisions = core_props
                .modified
                .or_else(|| Some(Utc::now().format("%Y-%m-%dT%H:%M:%SZ").to_string()));
        }
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
                    let (fn_ins, fn_del) =
                        count_revisions_smart(&fn_units1, &fn_units2, &fn_correlation);
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
                    let (en_ins, en_del) =
                        count_revisions_smart(&en_units1, &en_units2, &en_correlation);
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
            lcs_traces: None,
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
                if i + 1 < correlation.len()
                    && correlation[i + 1].status == lcs::CorrelationStatus::Inserted
                {
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

/// Debug helper: count pt:Status attributes in the document
fn count_pt_status(doc: &XmlDocument, root: NodeId, status_value: &str) -> usize {
    use crate::wml::coalesce::pt_status;
    let mut count = 0;
    for node in std::iter::once(root).chain(doc.descendants(root)) {
        if let Some(data) = doc.get(node) {
            if let Some(attrs) = data.attributes() {
                if attrs
                    .iter()
                    .any(|a| a.name == pt_status() && a.value == status_value)
                {
                    count += 1;
                }
            }
        }
    }
    count
}

/// Debug helper: count w:del and w:ins elements
fn count_del_ins(doc: &XmlDocument, root: NodeId) -> (usize, usize) {
    use crate::xml::namespaces::W;
    let mut del_count = 0;
    let mut ins_count = 0;
    for node in std::iter::once(root).chain(doc.descendants(root)) {
        if let Some(data) = doc.get(node) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) {
                    if name.local_name == "del" {
                        del_count += 1;
                    } else if name.local_name == "ins" {
                        ins_count += 1;
                    }
                }
            }
        }
    }
    (del_count, ins_count)
}

fn reconcile_formatting_changes(atoms: &mut [ComparisonUnitAtom], settings: &WmlComparerSettings) {
    detect_format_changes(atoms, settings);
}

/// Suppress deletions that represent content that "moved" to another location.
///
/// When text moves from one paragraph to another (e.g., a standalone paragraph is merged
/// into another paragraph), the comparison shows:
/// 1. The text as EQUAL in its new location (matched during LCS)
/// 2. The text as DELETED from its old location (the standalone paragraph)
///
/// This creates confusing output where the same text appears twice - once unchanged and
/// once as deleted. MS Word handles this by suppressing the deletion when the text
/// appears elsewhere as unchanged.
///
/// Algorithm:
/// 1. Extract all text sequences that appear as EQUAL
/// 2. Extract all text sequences that appear as DELETED
/// 3. For significant DELETED sequences (>= 50 chars) that appear in EQUAL text,
///    filter them out of the atom list
fn suppress_moved_deletions(
    mut atoms: Vec<ComparisonUnitAtom>,
    settings: &WmlComparerSettings,
) -> Vec<ComparisonUnitAtom> {
    use crate::util::group_adjacent;

    const MIN_TEXT_LENGTH: usize = 50; // Minimum length for move detection

    // Group atoms by correlation status to find runs of INSERTED and DELETED atoms
    let groups = group_adjacent(atoms.iter().enumerate(), |item| item.1.correlation_status);

    // Collect INSERTED and DELETED runs with their normalized text and indices
    let mut inserted_runs: Vec<(Vec<usize>, String)> = Vec::new();
    let mut deleted_runs: Vec<(Vec<usize>, String)> = Vec::new();

    for group in &groups {
        if group.is_empty() {
            continue;
        }

        let (_, first_atom) = group[0];
        let status = first_atom.correlation_status;

        if status != ComparisonCorrelationStatus::Inserted
            && status != ComparisonCorrelationStatus::Deleted
        {
            continue;
        }

        // Extract text from this run
        let mut text = String::new();
        let mut indices = Vec::new();

        for &(idx, atom) in group {
            if let ContentElement::Text(c) = atom.content_element {
                text.push(c);
            }
            indices.push(idx);
        }

        if text.is_empty() {
            continue;
        }

        // Normalize text for comparison
        let normalized = normalize_text_for_move_detection(&text, settings);

        if normalized.len() >= MIN_TEXT_LENGTH {
            match status {
                ComparisonCorrelationStatus::Inserted => inserted_runs.push((indices, normalized)),
                ComparisonCorrelationStatus::Deleted => deleted_runs.push((indices, normalized)),
                _ => {}
            }
        }
    }

    // Build set of deleted indices to filter (moved content at old location)
    // We DON'T blindly change all inserted atoms - that loses fine-grained comparison
    let mut deleted_indices_to_filter: std::collections::HashSet<usize> =
        std::collections::HashSet::new();

    for (del_indices, del_text) in &deleted_runs {
        for (_ins_indices, ins_text) in &inserted_runs {
            // Check if the deleted text is a substring of the inserted text (or vice versa)
            let is_move = ins_text.contains(del_text) || del_text.contains(ins_text);

            if is_move {
                // Only filter out the deleted atoms - the "moved" content at its old location
                // DON'T change the inserted atoms - preserve original LCS comparison results
                deleted_indices_to_filter.extend(del_indices.iter().copied());
                break;
            }
        }
    }

    // If nothing to filter, return original atoms unchanged
    if deleted_indices_to_filter.is_empty() {
        return atoms;
    }

    // Filter out DELETED atoms that represent moved text
    // The inserted atoms keep their original status from LCS comparison
    atoms
        .into_iter()
        .enumerate()
        .filter(|(idx, _)| !deleted_indices_to_filter.contains(idx))
        .map(|(_, atom)| atom)
        .collect()
}

/// Normalize text for move detection comparison
fn normalize_text_for_move_detection(text: &str, settings: &WmlComparerSettings) -> String {
    let mut result = String::with_capacity(text.len());

    for c in text.chars() {
        // Normalize case if case-insensitive
        let c = if settings.case_insensitive {
            c.to_lowercase().next().unwrap_or(c)
        } else {
            c
        };

        // Normalize whitespace: collapse multiple spaces, treat NBSP as space
        let c = if settings.conflate_breaking_and_nonbreaking_spaces && c == '\u{00A0}' {
            ' '
        } else {
            c
        };

        // Skip word separators for more robust matching
        // (punctuation differences shouldn't prevent move detection)
        if c.is_whitespace() {
            // Collapse whitespace
            if !result.ends_with(' ') && !result.is_empty() {
                result.push(' ');
            }
        } else if !settings.is_word_separator(c) || c.is_alphabetic() || c.is_numeric() {
            result.push(c);
        }
    }

    result.trim().to_string()
}

/// Port of C# GetRevisions grouping logic (WmlComparer.cs:3909-3926)
///
/// Uses GroupAdjacent to group atoms by a key that combines:
/// 1. CorrelationStatus (as string: "Equal", "Inserted", "Deleted", etc.)
/// 2. For non-Equal status, also includes RevTrackElement info (type, author, date - excluding id)
///
/// This means adjacent atoms with the same (status, author, date) count as ONE revision.
/// The revision count = number of groups where key != "Equal".
///
/// Returns (insertions, deletions)
/// Note: Format changes are counted separately from the final XML document (not from atoms)
fn count_revisions_from_atoms(atoms: &[ComparisonUnitAtom]) -> (usize, usize) {
    use crate::util::group_adjacent;

    // C# GetRevisions key function (lines 3910-3921):
    // - For Equal: key = "Equal"
    // - For non-Equal: key = status.ToString() + serialized RevTrackElement (minus id/Unid)
    //
    // The RevTrackElement serialization in C# creates an XElement with:
    // - The element name (w:ins or w:del)
    // - Attributes except w:id and PtOpenXml.Unid
    // - The xmlns:w namespace declaration
    //
    // Since our atoms don't store the full RevTrackElement, we use the rev_track_element
    // field (which is "ins" or "del") combined with the correlation status.
    // Since all revisions are generated with the same settings (author/date),
    // adjacent atoms with the same status will naturally have the same author/date.

    // Key function for grouping - takes &&ComparisonUnitAtom since we're iterating over references
    let key_fn = |atom: &&ComparisonUnitAtom| -> String {
        match atom.correlation_status {
            ComparisonCorrelationStatus::Equal => "Equal".to_string(),
            ComparisonCorrelationStatus::Inserted => {
                // C#: "Inserted<w:ins ... />" (serialized element)
                // We use "Inserted|ins" since author/date are uniform
                format!(
                    "Inserted|{}",
                    atom.rev_track_element.as_deref().unwrap_or("ins")
                )
            }
            ComparisonCorrelationStatus::Deleted => {
                format!(
                    "Deleted|{}",
                    atom.rev_track_element.as_deref().unwrap_or("del")
                )
            }
            ComparisonCorrelationStatus::FormatChanged => "FormatChanged".to_string(),
            status => status.to_string(),
        }
    };

    // Use group_adjacent to group atoms by key
    let groups = group_adjacent(atoms.iter(), key_fn);

    // Count revisions: number of groups where key starts with "Inserted" or "Deleted"
    // Note: FormatChanged is intentionally NOT counted here - format changes are
    // detected from the final XML document by looking for w:rPrChange and w:pPrChange elements
    let mut insertions = 0;
    let mut deletions = 0;

    for group in &groups {
        if let Some(first) = group.first() {
            match first.correlation_status {
                ComparisonCorrelationStatus::Inserted => insertions += 1,
                ComparisonCorrelationStatus::Deleted => deletions += 1,
                _ => {}
            }
        }
    }

    (insertions, deletions)
}

/// Port of C# NormalizeTxbxContentAncestorUnids (WmlComparer.cs:7571-7805)
///
/// This function normalizes ancestor UNIDs for atoms inside textboxes.
/// It groups atoms by txbxContent depth, subdivides into paragraph sub-groups,
/// and normalizes UNIDs at appropriate levels.
fn normalize_txbx_content_ancestor_unids(atoms: &mut [ComparisonUnitAtom]) {
    // Step 1: Find contiguous groups of atoms where any ancestor is txbxContent
    // Group by txbxContent depth (which ancestor index is txbxContent)
    let mut groups: Vec<Vec<usize>> = Vec::new();
    let mut current_group: Option<Vec<usize>> = None;
    let mut current_txbx_depth: Option<usize> = None;

    for (atom_idx, atom) in atoms.iter().enumerate() {
        // Find txbxContent depth for this atom
        let txbx_depth = atom
            .ancestor_elements
            .iter()
            .position(|a| a.local_name == "txbxContent");

        if let Some(depth) = txbx_depth {
            // This atom is inside txbxContent
            if current_group.is_none() || current_txbx_depth != Some(depth) {
                // Start a new group
                if let Some(group) = current_group.take() {
                    groups.push(group);
                }
                current_group = Some(vec![atom_idx]);
                current_txbx_depth = Some(depth);
            } else {
                // Add to current group
                current_group.as_mut().unwrap().push(atom_idx);
            }
        } else {
            // Not in txbxContent, end current group
            if let Some(group) = current_group.take() {
                groups.push(group);
            }
            current_group = None;
            current_txbx_depth = None;
        }
    }

    // Don't forget the last group
    if let Some(group) = current_group.take() {
        groups.push(group);
    }

    // Step 2: For each group, normalize all atoms to use consistent unids
    for group_indices in groups {
        if group_indices.is_empty() {
            continue;
        }

        // Find the txbxContent index from the first atom in the group
        let txbx_content_index = {
            let first_atom = &atoms[group_indices[0]];
            first_atom
                .ancestor_elements
                .iter()
                .position(|a| a.local_name == "txbxContent")
        };

        let txbx_content_index = match txbx_content_index {
            Some(idx) => idx,
            None => continue,
        };

        // Find a reference atom for OUTER levels (up to and including txbxContent)
        // Prefer an Equal atom which has source1's normalized unids
        let outer_ref_atom_idx = group_indices
            .iter()
            .find(|&&idx| {
                let atom = &atoms[idx];
                atom.correlation_status == ComparisonCorrelationStatus::Equal
            })
            .or_else(|| {
                group_indices.iter().find(|&&idx| {
                    let atom = &atoms[idx];
                    atom.correlation_status == ComparisonCorrelationStatus::Deleted
                })
            })
            .or_else(|| group_indices.iter().next());

        let outer_ref_atom_idx = match outer_ref_atom_idx {
            Some(&idx) => idx,
            None => continue,
        };

        // Step 3: Subdivide the group into paragraph sub-groups
        // Each pPr atom marks the start of a new paragraph
        let mut paragraph_sub_groups: Vec<Vec<usize>> = Vec::new();
        let mut current_paragraph: Option<Vec<usize>> = None;

        for &atom_idx in &group_indices {
            let atom = &atoms[atom_idx];

            // Check if this atom is a pPr (paragraph properties) - marks start of new paragraph
            if matches!(
                atom.content_element,
                ContentElement::ParagraphProperties { .. }
            ) {
                // Start new paragraph
                if let Some(para) = current_paragraph.take() {
                    paragraph_sub_groups.push(para);
                }
                current_paragraph = Some(vec![atom_idx]);
            } else {
                if current_paragraph.is_none() {
                    // Atom before first pPr - create a paragraph for it
                    current_paragraph = Some(vec![atom_idx]);
                } else {
                    current_paragraph.as_mut().unwrap().push(atom_idx);
                }
            }
        }

        // Don't forget the last paragraph
        if let Some(para) = current_paragraph.take() {
            paragraph_sub_groups.push(para);
        }

        // Step 4: For each paragraph sub-group, normalize unids
        for para_group_indices in paragraph_sub_groups {
            if para_group_indices.is_empty() {
                continue;
            }

            // Check if this paragraph has mixed correlation statuses (both Equal and Inserted/Deleted)
            let has_equal = para_group_indices
                .iter()
                .any(|&idx| atoms[idx].correlation_status == ComparisonCorrelationStatus::Equal);
            let has_inserted_or_deleted = para_group_indices.iter().any(|&idx| {
                atoms[idx].correlation_status == ComparisonCorrelationStatus::Inserted
                    || atoms[idx].correlation_status == ComparisonCorrelationStatus::Deleted
            });
            let is_mixed_paragraph = has_equal && has_inserted_or_deleted;

            // Find a reference atom for paragraph level
            let para_ref_atom_idx = para_group_indices
                .iter()
                .find(|&&idx| {
                    let atom = &atoms[idx];
                    atom.correlation_status == ComparisonCorrelationStatus::Equal
                })
                .or_else(|| {
                    para_group_indices.iter().find(|&&idx| {
                        let atom = &atoms[idx];
                        atom.correlation_status == ComparisonCorrelationStatus::Deleted
                    })
                })
                .or_else(|| para_group_indices.iter().next());

            // Find a reference atom for run level (needs to have run-level ancestors)
            let run_level_idx = txbx_content_index + 2;
            let run_ref_atom_idx = para_group_indices
                .iter()
                .find(|&&idx| {
                    let atom = &atoms[idx];
                    atom.correlation_status == ComparisonCorrelationStatus::Equal
                        && atom.ancestor_unids.len() > run_level_idx
                })
                .or_else(|| {
                    para_group_indices.iter().find(|&&idx| {
                        let atom = &atoms[idx];
                        atom.correlation_status == ComparisonCorrelationStatus::Deleted
                            && atom.ancestor_unids.len() > run_level_idx
                    })
                })
                .or_else(|| {
                    para_group_indices
                        .iter()
                        .find(|&&idx| atoms[idx].ancestor_unids.len() > run_level_idx)
                });

            // Clone the reference UNIDs we need before borrowing atoms mutably
            let outer_ref_unids = atoms[outer_ref_atom_idx].ancestor_unids.clone();
            let para_ref_unids = para_ref_atom_idx.map(|&idx| atoms[idx].ancestor_unids.clone());
            let run_ref_unids = run_ref_atom_idx.map(|&idx| atoms[idx].ancestor_unids.clone());

            // Step 5: Normalize UNIDs for each atom in the paragraph
            for &atom_idx in &para_group_indices {
                let atom = &mut atoms[atom_idx];

                // Find txbxContent index for this atom
                let this_atom_txbx_index = atom
                    .ancestor_elements
                    .iter()
                    .position(|a| a.local_name == "txbxContent");

                if this_atom_txbx_index != Some(txbx_content_index) {
                    continue;
                }

                // Normalize outer levels and paragraph level:
                // - Outer levels (0 to txbxContentIndex) use the group's outer reference atom
                // - Paragraph level (txbxContentIndex + 1) uses this paragraph's inner reference atom
                // - Run level (txbxContentIndex + 2) is ONLY normalized for mixed paragraphs
                let paragraph_level_index = txbx_content_index + 1;
                let run_level_index = txbx_content_index + 2;
                let max_level_to_normalize = if is_mixed_paragraph {
                    // Mixed paragraph - also normalize run level
                    (run_level_index + 1).min(atom.ancestor_unids.len())
                } else {
                    // Pure paragraph - keep runs separate
                    (paragraph_level_index + 1).min(atom.ancestor_unids.len())
                };

                for i in 0..max_level_to_normalize {
                    let ref_unid: Option<String> = if i <= txbx_content_index {
                        // Outer level - use the group's outer reference atom
                        if i < outer_ref_unids.len() {
                            Some(outer_ref_unids[i].clone())
                        } else {
                            None
                        }
                    } else if i == paragraph_level_index {
                        // Paragraph level - use the paragraph reference atom
                        para_ref_unids
                            .as_ref()
                            .and_then(|unids: &Vec<String>| unids.get(i).cloned())
                    } else if i == run_level_index && is_mixed_paragraph {
                        // Run level - only for mixed paragraphs
                        run_ref_unids
                            .as_ref()
                            .and_then(|unids: &Vec<String>| unids.get(i).cloned())
                    } else {
                        None
                    };

                    if let Some(ref ref_unid_val) = ref_unid {
                        // Update both the ancestor element's unid attribute and the ancestor_unids array
                        if i < atom.ancestor_elements.len() {
                            atom.ancestor_elements[i].unid = ref_unid_val.clone();
                        }
                        if i < atom.ancestor_unids.len() {
                            atom.ancestor_unids[i] = ref_unid_val.clone();
                        }
                    }
                }
            }
        }
    }
}

fn assemble_ancestor_unids(atoms: &mut [ComparisonUnitAtom]) {
    // Phase 1: Initial UNID normalization (C# lines 3441-3494)
    // For atoms inside textboxes (txbxContent), copy UNIDs from "before" document to "after" document.
    // This applies to ALL atoms inside textboxes with Equal status, not just pPr atoms.
    for atom in atoms.iter_mut() {
        let is_in_textbox = atom
            .ancestor_elements
            .iter()
            .any(|a| a.local_name == "txbxContent");

        let do_set = if matches!(
            atom.content_element,
            ContentElement::ParagraphProperties { .. }
        ) {
            // pPr: normalize if in textbox OR if status is Equal
            is_in_textbox || atom.correlation_status == ComparisonCorrelationStatus::Equal
        } else {
            // Other atoms: normalize only if in textbox AND status is Equal
            is_in_textbox && atom.correlation_status == ComparisonCorrelationStatus::Equal
        };

        if do_set {
            if let Some(ref before) = atom.ancestor_elements_before {
                if atom.ancestor_elements.len() == before.len() {
                    for i in 0..atom.ancestor_elements.len() {
                        atom.ancestor_elements[i].unid = before[i].unid.clone();
                    }
                }
            }
        }
    }

    let deepest_ancestor_unid = atoms
        .iter()
        .rev()
        .next()
        .and_then(|atom| atom.ancestor_elements.first())
        .and_then(|ancestor| {
            if ancestor.local_name == "footnote" || ancestor.local_name == "endnote" {
                Some(ancestor.unid.clone())
            } else {
                None
            }
        });

    // Phase 2a: First pass - process non-textbox paragraphs (C# lines 3531-3578)
    // Key insight: C# iterates in reverse and propagates currentAncestorUnids to ALL
    // subsequent atoms until the next pPr is hit. No paragraph boundary checking needed
    // because pPr is always last in a paragraph (document order), so first in reverse.
    let mut current_ancestor_unids: Vec<String> = Vec::new();

    for atom in atoms.iter_mut().rev() {
        if matches!(
            atom.content_element,
            ContentElement::ParagraphProperties { .. }
        ) {
            let ppr_in_textbox = atom
                .ancestor_elements
                .iter()
                .any(|ae| ae.local_name == "txbxContent");

            if !ppr_in_textbox {
                // C# lines 3544-3554: Collect ancestor unids for the paragraph
                current_ancestor_unids = atom
                    .ancestor_elements
                    .iter()
                    .map(|ae| ae.unid.clone())
                    .collect();
                atom.ancestor_unids = current_ancestor_unids.clone();
                // C# lines 3555-3556: Override deepest ancestor with footnote/endnote UNID
                if let Some(ref unid) = deepest_ancestor_unid {
                    if !atom.ancestor_unids.is_empty() {
                        atom.ancestor_unids[0] = unid.clone();
                    }
                }
                continue;
            }
        }

        // C# lines 3561-3577: For non-pPr atoms, propagate ancestor unids from current paragraph
        // Note: C# does NOT check if atom belongs to same paragraph - it just propagates
        // currentAncestorUnids to all atoms until the next pPr is encountered.
        let additional_unids: Vec<String> = atom
            .ancestor_elements
            .iter()
            .skip(current_ancestor_unids.len())
            .map(|ae| ae.unid.clone())
            .collect();

        let mut this_ancestor_unids = current_ancestor_unids.clone();
        this_ancestor_unids.extend(additional_unids);
        atom.ancestor_unids = this_ancestor_unids;

        // C# lines 3576-3577: Override deepest ancestor with footnote/endnote UNID
        if let Some(ref unid) = deepest_ancestor_unid {
            if !atom.ancestor_unids.is_empty() {
                atom.ancestor_unids[0] = unid.clone();
            }
        }
    }

    // Phase 2b: Second pass - process textbox content specifically (C# lines 3589-3658)
    current_ancestor_unids = Vec::new();
    let mut skip_until_next_ppr = false;

    for atom in atoms.iter_mut().rev() {
        if !current_ancestor_unids.is_empty()
            && atom.ancestor_elements.len() < current_ancestor_unids.len()
        {
            skip_until_next_ppr = true;
            current_ancestor_unids = Vec::new();
            continue;
        }

        if matches!(
            atom.content_element,
            ContentElement::ParagraphProperties { .. }
        ) {
            let ppr_in_textbox = atom
                .ancestor_elements
                .iter()
                .any(|ae| ae.local_name == "txbxContent");

            if !ppr_in_textbox {
                skip_until_next_ppr = true;
                current_ancestor_unids = Vec::new();
                continue;
            } else {
                skip_until_next_ppr = false;
                current_ancestor_unids = atom
                    .ancestor_elements
                    .iter()
                    .map(|ae| ae.unid.clone())
                    .collect();
                atom.ancestor_unids = current_ancestor_unids.clone();
                continue;
            }
        }

        if skip_until_next_ppr {
            continue;
        }

        // For atoms inside textbox paragraphs
        let additional_unids: Vec<String> = atom
            .ancestor_elements
            .iter()
            .skip(current_ancestor_unids.len())
            .map(|ae| ae.unid.clone())
            .collect();

        let mut this_ancestor_unids = current_ancestor_unids.clone();
        this_ancestor_unids.extend(additional_unids);
        atom.ancestor_unids = this_ancestor_unids;
    }
}

fn compare_atoms_internal(
    atoms1: Vec<ComparisonUnitAtom>,
    atoms2: Vec<ComparisonUnitAtom>,
    root_name: XName,
    root_attrs: Vec<XAttribute>,
    settings: &WmlComparerSettings,
) -> Result<(
    usize,
    usize,
    usize,
    super::coalesce::CoalesceResult,
    Vec<ComparisonUnitAtom>,
    Option<Vec<super::settings::LcsTraceOutput>>,
)> {
    // Timing is only available on native platforms, not WASM
    #[cfg(not(target_arch = "wasm32"))]
    let timing_enabled = std::env::var("WML_TIMING").is_ok();
    #[cfg(not(target_arch = "wasm32"))]
    let t0 = std::time::Instant::now();

    let mut word_settings = WordSeparatorSettings::default();
    if settings.conflate_breaking_and_nonbreaking_spaces {
        word_settings.word_separators.push('\u{00a0}');
    }

    let units1 = get_comparison_unit_list(atoms1, &word_settings);
    let units2 = get_comparison_unit_list(atoms2, &word_settings);
    #[cfg(not(target_arch = "wasm32"))]
    if timing_enabled {
        eprintln!("    unit_list: {:?}", t0.elapsed());
    }

    // Trace generation - only compiled when "trace" feature is enabled
    #[cfg(feature = "trace")]
    let lcs_traces = {
        let matched1 = units_match_filter(&units1, settings);
        let matched2 = units_match_filter(&units2, settings);

        if let Some(ref m1) = matched1 {
            let trace = generate_focused_trace(&units1, &units2, m1, matched2.as_ref(), settings);
            Some(vec![trace])
        } else if let Some(ref m2) = matched2 {
            let trace = generate_focused_trace(&units2, &units1, m2, matched1.as_ref(), settings);
            Some(vec![trace])
        } else {
            None
        }
    };

    // No trace when feature is disabled - zero overhead
    #[cfg(not(feature = "trace"))]
    let lcs_traces: Option<Vec<super::settings::LcsTraceOutput>> = None;

    let correlated = lcs(units1.clone(), units2.clone(), settings);
    #[cfg(not(target_arch = "wasm32"))]
    if timing_enabled {
        eprintln!("    lcs: {:?}", t0.elapsed());
    }

    let mut flattened_atoms = flatten_to_atoms(&correlated);
    #[cfg(not(target_arch = "wasm32"))]
    if timing_enabled {
        eprintln!("    flatten: {:?}", t0.elapsed());
    }

    assemble_ancestor_unids(&mut flattened_atoms);
    #[cfg(not(target_arch = "wasm32"))]
    if timing_enabled {
        eprintln!("    ancestor_unids: {:?}", t0.elapsed());
    }

    // Phase 3: Additional normalization for textbox content (C# lines 7571-7805)
    normalize_txbx_content_ancestor_unids(&mut flattened_atoms);

    if settings.track_formatting_changes {
        reconcile_formatting_changes(&mut flattened_atoms, settings);
    }

    // Suppress deletions that represent "moved" content (appears elsewhere as EQUAL)
    // This handles the case where text moved from one paragraph to another
    flattened_atoms = suppress_moved_deletions(flattened_atoms, settings);
    #[cfg(not(target_arch = "wasm32"))]
    if timing_enabled {
        eprintln!("    pre_coalesce: {:?}", t0.elapsed());
    }

    // Count revisions from atom list (C# GetRevisions algorithm)
    // This groups adjacent atoms by correlation status, which is how C# counts
    let mut coalesce_result = coalesce(&flattened_atoms, settings, root_name, root_attrs);
    #[cfg(not(target_arch = "wasm32"))]
    if timing_enabled {
        eprintln!("    coalesce: {:?}", t0.elapsed());
    }

    // Wrap content in revision marks (C# line 2173)
    mark_content_as_deleted_or_inserted(
        &mut coalesce_result.document,
        coalesce_result.root,
        settings,
    );

    // Consolidate adjacent revisions (C# line 2174)
    coalesce_adjacent_runs(
        &mut coalesce_result.document,
        coalesce_result.root,
        &settings,
    );

    // Clean up empty w:rPr elements after all processing
    // This MUST happen after mark_content_as_deleted_or_inserted and coalesce_adjacent_runs
    // because those functions can create new empty rPr elements
    super::coalesce::remove_empty_rpr_elements(&mut coalesce_result.document, coalesce_result.root);
    #[cfg(not(target_arch = "wasm32"))]
    if timing_enabled {
        eprintln!("    post_process: {:?}", t0.elapsed());
    }

    // CRITICAL: Count revisions AFTER coalescing and merging, not before
    // This ensures we count the actual w:ins/w:del elements in the final XML,
    // not the pre-merged atoms. The C# code's GetRevisions function operates
    // on the final XML tree after CoalesceAdjacentRunsWithIdenticalFormatting.
    strip_pt_attributes(&mut coalesce_result.document, coalesce_result.root);
    let revision_counts = count_revisions(&coalesce_result.document, coalesce_result.root);
    let insertions = revision_counts.insertions;
    let deletions = revision_counts.deletions;
    let format_changes = revision_counts.format_changes;

    // Return flattened atoms for note reference collection
    Ok((
        insertions,
        deletions,
        format_changes,
        coalesce_result,
        flattened_atoms,
        lcs_traces,
    ))
}

struct NoteProcessingResult {
    insertions: usize,
    deletions: usize,
    format_changes: usize,
    document: XmlDocument,
}

struct NoteReference {
    before_id: Option<String>,
    after_id: String,
    status: ComparisonCorrelationStatus,
}

/// Collect note reference IDs from comparison unit atoms
/// Matches C# WmlComparer.cs:2910-2914
fn collect_note_references(atoms: &[ComparisonUnitAtom], note_type: &str) -> Vec<NoteReference> {
    let mut references = Vec::new();

    for atom in atoms {
        let is_match = match &atom.content_element {
            ContentElement::FootnoteReference { id } if note_type == "footnoteReference" => {
                Some(id.clone())
            }
            ContentElement::EndnoteReference { id } if note_type == "endnoteReference" => {
                Some(id.clone())
            }
            _ => None,
        };

        if let Some(after_id) = is_match {
            let before_id = if atom.correlation_status == ComparisonCorrelationStatus::Equal {
                atom.content_element_before
                    .as_ref()
                    .and_then(|before| match before {
                        ContentElement::FootnoteReference { id }
                            if note_type == "footnoteReference" =>
                        {
                            Some(id.clone())
                        }
                        ContentElement::EndnoteReference { id }
                            if note_type == "endnoteReference" =>
                        {
                            Some(id.clone())
                        }
                        _ => None,
                    })
            } else {
                None
            };

            references.push(NoteReference {
                before_id,
                after_id,
                status: atom.correlation_status,
            });
        }
    }

    references
}

/// Process notes (footnotes/endnotes) using per-reference comparison
/// Matches C# WmlComparer.cs:2885-3248 ProcessFootnoteEndnote
fn process_notes(
    source1_opt: Option<XmlDocument>,
    source2_opt: Option<XmlDocument>,
    part_type: &str,
    reference_ids: &[NoteReference],
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

    let result_root = if part_type == "footnotes" {
        find_footnotes_root(&result_doc)
    } else {
        find_endnotes_root(&result_doc)
    };

    let root1_opt = source1_opt.as_ref().and_then(|doc| {
        if part_type == "footnotes" {
            find_footnotes_root(doc)
        } else {
            find_endnotes_root(doc)
        }
    });
    let root2_opt = source2_opt.as_ref().and_then(|doc| {
        if part_type == "footnotes" {
            find_footnotes_root(doc)
        } else {
            find_endnotes_root(doc)
        }
    });

    let mut note_statuses: HashMap<String, NoteReference> = HashMap::new();
    for note_ref in reference_ids {
        let key = format!(
            "{}|{}",
            note_ref.before_id.clone().unwrap_or_default(),
            note_ref.after_id
        );
        note_statuses
            .entry(key)
            .and_modify(|existing| {
                if note_status_priority(note_ref.status) > note_status_priority(existing.status) {
                    *existing = NoteReference {
                        before_id: note_ref.before_id.clone(),
                        after_id: note_ref.after_id.clone(),
                        status: note_ref.status,
                    };
                }
            })
            .or_insert(NoteReference {
                before_id: note_ref.before_id.clone(),
                after_id: note_ref.after_id.clone(),
                status: note_ref.status,
            });
    }

    // Process each reference individually
    for note_ref in note_statuses.into_values() {
        if note_ref.after_id == "0" || note_ref.after_id == "-1" {
            continue;
        }

        let after_id = note_ref.after_id.as_str();
        let before_id = note_ref.before_id.as_deref().unwrap_or(after_id);

        match note_ref.status {
            ComparisonCorrelationStatus::Equal => {
                // Both documents have this note - compare them
                if let (Some(ref doc1), Some(ref doc2), Some(root1), Some(root2)) =
                    (&source1_opt, &source2_opt, root1_opt, root2_opt)
                {
                    if let (Some(_note1_id), Some(note2_id)) = (
                        find_note_by_id(doc1, root1, before_id),
                        find_note_by_id(doc2, root2, after_id),
                    ) {
                        let mut doc1_clone = clone_xml_doc(doc1)?;
                        let mut doc2_clone = clone_xml_doc(doc2)?;

                        let root1_clone = if part_type == "footnotes" {
                            find_footnotes_root(&doc1_clone)
                        } else {
                            find_endnotes_root(&doc1_clone)
                        };
                        let root2_clone = if part_type == "footnotes" {
                            find_footnotes_root(&doc2_clone)
                        } else {
                            find_endnotes_root(&doc2_clone)
                        };

                        let (note1_clone, note2_clone) = match (root1_clone, root2_clone) {
                            (Some(r1), Some(r2)) => (
                                find_note_by_id(&doc1_clone, r1, before_id),
                                find_note_by_id(&doc2_clone, r2, after_id),
                            ),
                            _ => (None, None),
                        };

                        let (Some(note1_clone), Some(note2_clone)) = (note1_clone, note2_clone)
                        else {
                            continue;
                        };

                        let (ins, del, fmt, coalesce_res) = compare_part_content(
                            &mut doc1_clone,
                            note1_clone,
                            &mut doc2_clone,
                            note2_clone,
                            part_type,
                            settings,
                        )?;

                        total_ins += ins;
                        total_del += del;
                        total_fmt += fmt;

                        if let Some(result_root) = result_root {
                            update_note_in_result(
                                &mut result_doc,
                                result_root,
                                after_id,
                                Some((doc2, note2_id)),
                                &coalesce_res.document,
                                coalesce_res.root,
                                part_type,
                            );
                        }
                    }
                }
            }
            ComparisonCorrelationStatus::Inserted => {
                // Note exists only in doc2 - all content is inserted
                if let (Some(ref doc2), Some(root2)) = (&source2_opt, root2_opt) {
                    if let Some(note_id) = find_note_by_id(doc2, root2, after_id) {
                        let (ins, del, fmt, coalesce_res) = build_note_doc_with_status(
                            doc2,
                            after_id,
                            part_type,
                            ComparisonCorrelationStatus::Inserted,
                            settings,
                        )?;
                        total_ins += ins;
                        total_del += del;
                        total_fmt += fmt;

                        if let Some(result_root) = result_root {
                            update_note_in_result(
                                &mut result_doc,
                                result_root,
                                after_id,
                                Some((doc2, note_id)),
                                &coalesce_res.document,
                                coalesce_res.root,
                                part_type,
                            );
                        }
                    }
                }
            }
            ComparisonCorrelationStatus::Deleted => {
                // Note exists only in doc1 - all content is deleted
                if let (Some(ref doc1), Some(root1)) = (&source1_opt, root1_opt) {
                    if let Some(note_id) = find_note_by_id(doc1, root1, before_id) {
                        let (ins, del, fmt, coalesce_res) = build_note_doc_with_status(
                            doc1,
                            before_id,
                            part_type,
                            ComparisonCorrelationStatus::Deleted,
                            settings,
                        )?;
                        total_ins += ins;
                        total_del += del;
                        total_fmt += fmt;

                        if let Some(result_root) = result_root {
                            update_note_in_result(
                                &mut result_doc,
                                result_root,
                                before_id,
                                Some((doc1, note_id)),
                                &coalesce_res.document,
                                coalesce_res.root,
                                part_type,
                            );
                        }
                    }
                }
            }
            _ => {
                // Ignore other correlation statuses
            }
        }
    }

    Ok(NoteProcessingResult {
        insertions: total_ins,
        deletions: total_del,
        format_changes: total_fmt,
        document: result_doc,
    })
}

fn note_status_priority(status: ComparisonCorrelationStatus) -> usize {
    match status {
        ComparisonCorrelationStatus::Equal => 3,
        ComparisonCorrelationStatus::Inserted => 2,
        ComparisonCorrelationStatus::Deleted => 1,
        _ => 0,
    }
}

fn clone_xml_doc(doc: &XmlDocument) -> Result<XmlDocument> {
    let xml = crate::xml::builder::serialize(doc)?;
    crate::xml::parser::parse(&xml)
}

fn build_note_doc_with_status(
    source_doc: &XmlDocument,
    ref_id: &str,
    part_type: &str,
    status: ComparisonCorrelationStatus,
    settings: &WmlComparerSettings,
) -> Result<(usize, usize, usize, super::coalesce::CoalesceResult)> {
    let mut doc = clone_xml_doc(source_doc)?;

    let root = if part_type == "footnotes" {
        find_footnotes_root(&doc)
    } else {
        find_endnotes_root(&doc)
    };

    let Some(root) = root else {
        let fallback_root = doc.root().ok_or_else(|| {
            RedlineError::ComparisonFailed("No root in note document".to_string())
        })?;
        return Ok((
            0,
            0,
            0,
            super::coalesce::CoalesceResult {
                document: doc,
                root: fallback_root,
            },
        ));
    };

    let Some(note_id) = find_note_by_id(&doc, root, ref_id) else {
        return Ok((
            0,
            0,
            0,
            super::coalesce::CoalesceResult {
                document: doc,
                root,
            },
        ));
    };

    // Preprocess the footnote/endnote content to add UNIDs
    // This is critical for the coalesce grouping to work correctly
    let preprocess_settings = PreProcessSettings::for_comparison();
    preprocess_markup(&mut doc, note_id, &preprocess_settings)
        .map_err(RedlineError::ComparisonFailed)?;

    let mut atoms = create_comparison_unit_atom_list(&mut doc, note_id, part_type, settings);
    for atom in atoms.iter_mut() {
        atom.correlation_status = status;
    }

    assemble_ancestor_unids(&mut atoms);
    normalize_txbx_content_ancestor_unids(&mut atoms);
    if settings.track_formatting_changes {
        reconcile_formatting_changes(&mut atoms, settings);
    }

    let root_data = doc.get(note_id).ok_or_else(|| {
        RedlineError::ComparisonFailed("Failed to get note node data".to_string())
    })?;
    let root_name = root_data
        .name()
        .ok_or_else(|| RedlineError::ComparisonFailed("Note node has no name".to_string()))?
        .clone();
    let root_attrs = root_data.attributes().unwrap_or(&[]).to_vec();

    let (ins, del) = count_revisions_from_atoms(&atoms);
    let mut coalesce_result = coalesce(&atoms, settings, root_name, root_attrs);
    mark_content_as_deleted_or_inserted(
        &mut coalesce_result.document,
        coalesce_result.root,
        settings,
    );
    coalesce_adjacent_runs(
        &mut coalesce_result.document,
        coalesce_result.root,
        settings,
    );

    // Clean up empty w:rPr elements after all processing
    super::coalesce::remove_empty_rpr_elements(&mut coalesce_result.document, coalesce_result.root);

    strip_pt_attributes(&mut coalesce_result.document, coalesce_result.root);
    let fmt = count_revisions(&coalesce_result.document, coalesce_result.root).format_changes;

    Ok((ins, del, fmt, coalesce_result))
}

fn compare_part_content(
    doc1: &mut XmlDocument,
    root1: NodeId,
    doc2: &mut XmlDocument,
    root2: NodeId,
    part_name: &str,
    settings: &WmlComparerSettings,
) -> Result<(usize, usize, usize, super::coalesce::CoalesceResult)> {
    // Preprocess both documents to add UNIDs
    // This is critical for the coalesce grouping to work correctly
    let preprocess_settings = PreProcessSettings::for_comparison();
    preprocess_markup(doc1, root1, &preprocess_settings).map_err(RedlineError::ComparisonFailed)?;
    preprocess_markup(doc2, root2, &preprocess_settings).map_err(RedlineError::ComparisonFailed)?;

    let atoms1 = create_comparison_unit_atom_list(doc1, root1, part_name, settings);
    let atoms2 = create_comparison_unit_atom_list(doc2, root2, part_name, settings);

    let root_data = doc2
        .get(root2)
        .ok_or_else(|| RedlineError::ComparisonFailed("Failed to get root2 data".to_string()))?;
    let root_name = root_data
        .name()
        .ok_or_else(|| RedlineError::ComparisonFailed("Root2 has no name".to_string()))?
        .clone();
    let root_attrs = root_data.attributes().unwrap_or(&[]).to_vec();

    // Call compare_atoms_internal and discard the flattened atoms and traces (not needed for notes)
    let (ins, del, fmt, coalesce_result, _flattened_atoms, _traces) =
        compare_atoms_internal(atoms1, atoms2, root_name, root_attrs, settings)?;

    Ok((ins, del, fmt, coalesce_result))
}

fn update_note_in_result(
    result_doc: &mut XmlDocument,
    result_root: NodeId,
    ref_id: &str,
    source_note: Option<(&XmlDocument, NodeId)>,
    updated_doc: &XmlDocument,
    updated_root: NodeId,
    part_type: &str,
) {
    let mut note_node = find_note_by_id(result_doc, result_root, ref_id);
    if note_node.is_none() {
        if let Some((source_doc, source_node)) = source_note {
            note_node = Some(append_cloned_element(
                result_doc,
                result_root,
                source_doc,
                source_node,
            ));
        } else {
            note_node = Some(append_cloned_element(
                result_doc,
                result_root,
                updated_doc,
                updated_root,
            ));
        }
    }

    let Some(note_node) = note_node else {
        return;
    };
    replace_children_with(result_doc, note_node, updated_doc, updated_root);
    ensure_note_reference_run(result_doc, note_node, part_type);
}

fn replace_children_with(
    target_doc: &mut XmlDocument,
    target_parent: NodeId,
    source_doc: &XmlDocument,
    source_parent: NodeId,
) {
    // Remove all existing children from target
    let children: Vec<_> = target_doc.children(target_parent).collect();
    for child in children {
        target_doc.remove(child);
    }

    // Clone children from source
    let source_children: Vec<_> = source_doc.children(source_parent).collect();
    for child in source_children {
        clone_subtree(source_doc, child, target_doc, target_parent);
    }
}

fn append_cloned_element(
    target_doc: &mut XmlDocument,
    target_parent: NodeId,
    source_doc: &XmlDocument,
    source_node: NodeId,
) -> NodeId {
    let data = source_doc.get(source_node).unwrap().clone();
    let new_node = target_doc.add_child(target_parent, data);
    let source_children: Vec<_> = source_doc.children(source_node).collect();
    for child in source_children {
        clone_subtree(source_doc, child, target_doc, new_node);
    }
    new_node
}

fn clone_subtree(
    source_doc: &XmlDocument,
    source_node: NodeId,
    target_doc: &mut XmlDocument,
    target_parent: NodeId,
) -> NodeId {
    let data = source_doc.get(source_node).unwrap().clone();
    let new_node = target_doc.add_child(target_parent, data);
    let children: Vec<_> = source_doc.children(source_node).collect();
    for child in children {
        clone_subtree(source_doc, child, target_doc, new_node);
    }
    new_node
}

fn ensure_note_reference_run(doc: &mut XmlDocument, note_node: NodeId, part_type: &str) {
    if note_has_reference_run(doc, note_node, part_type) {
        return;
    }

    let Some(first_para) = find_first_descendant(doc, note_node, &W::p()) else {
        return;
    };
    insert_reference_run(doc, first_para, part_type);
}

fn note_has_reference_run(doc: &XmlDocument, note_node: NodeId, part_type: &str) -> bool {
    let (style_val, ref_name) = note_reference_marker(part_type);

    for node in doc.descendants(note_node) {
        let Some(name) = doc.get(node).and_then(|d| d.name()) else {
            continue;
        };
        if name == &ref_name {
            return true;
        }
        if name == &W::r() && run_has_rstyle(doc, node, style_val) {
            return true;
        }
    }

    false
}

fn run_has_rstyle(doc: &XmlDocument, run: NodeId, style_val: &str) -> bool {
    for child in doc.children(run) {
        let Some(name) = doc.get(child).and_then(|d| d.name()) else {
            continue;
        };
        if name != &W::rPr() {
            continue;
        }
        for rpr_child in doc.children(child) {
            let Some(rpr_name) = doc.get(rpr_child).and_then(|d| d.name()) else {
                continue;
            };
            if rpr_name != &W::r_style() {
                continue;
            }
            if let Some(attrs) = doc.get(rpr_child).and_then(|d| d.attributes()) {
                if attrs
                    .iter()
                    .any(|a| a.name == W::val() && a.value == style_val)
                {
                    return true;
                }
            }
        }
    }
    false
}

fn insert_reference_run(doc: &mut XmlDocument, para: NodeId, part_type: &str) {
    let (style_val, ref_name) = note_reference_marker(part_type);

    let run = doc.new_node(XmlNodeData::element(W::r()));
    let rpr = doc.add_child(run, XmlNodeData::element(W::rPr()));
    doc.add_child(
        rpr,
        XmlNodeData::element_with_attrs(W::r_style(), vec![XAttribute::new(W::val(), style_val)]),
    );
    doc.add_child(run, XmlNodeData::element(ref_name));

    let first_child = doc.children(para).next();
    if let Some(first_child) = first_child {
        doc.insert_before(first_child, run);
    } else {
        doc.reparent(para, run);
    }
}

fn note_reference_marker(part_type: &str) -> (&'static str, XName) {
    if part_type == "footnotes" {
        ("FootnoteReference", W::footnoteRef())
    } else {
        ("EndnoteReference", W::endnoteRef())
    }
}

fn find_first_descendant(doc: &XmlDocument, root: NodeId, name: &XName) -> Option<NodeId> {
    for node in doc.descendants(root) {
        if doc.get(node).and_then(|d| d.name()) == Some(name) {
            return Some(node);
        }
    }
    None
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
