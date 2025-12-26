//! Coalesce - Reconstruct XML tree from comparison atoms
//!
//! This is a port of the C# CoalesceRecurse from WmlComparer.cs (lines 5161-5738).
//!
//! The algorithm:
//! 1. Group atoms by their ancestor Unids at each tree level
//! 2. Recursively reconstruct elements (paragraphs, runs, text, tables, drawings)
//! 3. Add pt:Status markers ("Deleted"/"Inserted") based on CorrelationStatus
//! 4. Later, MarkContentAsDeletedOrInserted converts markers to w:ins/w:del
//!
//! Key features ported from C#:
//! - txbxContent grouping: uses "TXBX" marker and forces Equal status
//! - VML content: special handling to preserve structure
//! - Text coalescing: combines adjacent text atoms into single w:t elements
//! - xml:space="preserve": added when text has leading/trailing whitespace
//! - Formatting signatures: prevents merging atoms with different formatting

use super::comparison_unit::{ComparisonCorrelationStatus, ComparisonUnitAtom, ContentElement};
use super::settings::WmlComparerSettings;
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{M, W};
use crate::xml::node::XmlNodeData;
use crate::xml::xname::{XAttribute, XName};
use indextree::NodeId;
use std::collections::HashMap;

/// PowerTools status attribute namespace and name
pub const PT_STATUS_NS: &str = "http://powertools.codeplex.com/2011";

/// Create the pt:Status attribute name
pub fn pt_status() -> XName {
    XName::new(PT_STATUS_NS, "Status")
}

/// Result of coalescing - a new document with status markers
pub struct CoalesceResult {
    /// The reconstructed document
    pub document: XmlDocument,
    /// Root node ID
    pub root: NodeId,
}

/// Produce new WML markup from a correlated sequence of atoms
///
/// This is the main entry point, equivalent to C# ProduceNewWmlMarkupFromCorrelatedSequence
pub fn produce_markup_from_atoms(
    atoms: &[ComparisonUnitAtom],
    settings: &WmlComparerSettings,
) -> CoalesceResult {
    let mut doc = XmlDocument::new();
    
    // Create document root (w:document)
    let doc_root = doc.add_root(XmlNodeData::element(W::document()));
    
    // Create body
    let body = doc.add_child(doc_root, XmlNodeData::element(W::body()));
    
    // Coalesce atoms into body children
    coalesce_recurse(&mut doc, body, atoms, 0, settings);
    
    CoalesceResult {
        document: doc,
        root: doc_root,
    }
}

/// Recursively coalesce atoms into XML tree structure
///
/// This groups atoms by their ancestor Unid at the current level,
/// then recursively builds child elements.
fn coalesce_recurse(
    doc: &mut XmlDocument,
    parent: NodeId,
    atoms: &[ComparisonUnitAtom],
    level: usize,
    settings: &WmlComparerSettings,
) {
    if atoms.is_empty() {
        return;
    }
    
    // Group atoms by their ancestor Unid at this level
    let groups = group_by_ancestor_unid(atoms, level);
    
    if groups.is_empty() {
        return;
    }
    
    for (unid, group_atoms) in groups {
        if unid.is_empty() || group_atoms.is_empty() {
            continue;
        }
        
        let first_atom = &group_atoms[0];
        if level >= first_atom.ancestor_elements.len() {
            let is_inside_vml = is_inside_vml_content(first_atom, level.saturating_sub(1));
            coalesce_text_content(
                doc,
                parent,
                &group_atoms,
                first_atom.correlation_status,
                is_inside_vml,
            );
            continue;
        }
        
        let ancestor_info = &first_atom.ancestor_elements[level];
        let local_name = &ancestor_info.local_name;
        
        // Create the appropriate element based on ancestor type
        match local_name.as_str() {
            "p" => {
                let para = create_paragraph(doc, parent, &unid);
                coalesce_paragraph_children(doc, para, &group_atoms, level, settings);
            }
            "r" => {
                let run = create_run(doc, parent, &unid);
                coalesce_run_children(doc, run, &group_atoms, level, settings);
            }
            "tbl" => {
                let table = create_table(doc, parent, &unid);
                coalesce_recurse(doc, table, &group_atoms, level + 1, settings);
            }
            "tr" => {
                let row = create_table_row(doc, parent, &unid);
                coalesce_recurse(doc, row, &group_atoms, level + 1, settings);
            }
            "tc" => {
                let cell = create_table_cell(doc, parent, &unid);
                coalesce_recurse(doc, cell, &group_atoms, level + 1, settings);
            }
            "txbxContent" => {
                let txbx = create_textbox_content(doc, parent);
                coalesce_recurse(doc, txbx, &group_atoms, level + 1, settings);
            }
            _ => {
                // Generic element - recurse
                let elem = create_generic_element(doc, parent, local_name, &unid);
                coalesce_recurse(doc, elem, &group_atoms, level + 1, settings);
            }
        }
    }
}

/// Group atoms by their ancestor Unid at the given level
/// Preserves first-seen order (matching C# LINQ GroupBy behavior)
fn group_by_ancestor_unid(atoms: &[ComparisonUnitAtom], level: usize) -> Vec<(String, Vec<ComparisonUnitAtom>)> {
    let mut groups: HashMap<String, Vec<ComparisonUnitAtom>> = HashMap::new();
    let mut order: Vec<String> = Vec::new();
    
    for atom in atoms {
        let unid = if level < atom.ancestor_elements.len() {
            atom.ancestor_elements[level].unid.clone()
        } else {
            String::new()
        };
        
        if !groups.contains_key(&unid) {
            order.push(unid.clone());
        }
        groups.entry(unid).or_default().push(atom.clone());
    }
    
    order.into_iter()
        .filter_map(|unid| {
            groups.remove(&unid).map(|atoms| (unid, atoms))
        })
        .collect()
}

/// Key for grouping adjacent atoms (ports C# GroupAdjacent key)
#[derive(Clone, PartialEq, Eq, Debug)]
struct AdjacentGroupKey {
    child_unid: String,
    status: ComparisonCorrelationStatus,
    formatting_sig: Option<String>,
    before_sig: Option<String>,
}

/// Group adjacent atoms by key (ports C# GroupAdjacent)
fn group_adjacent_by_key(
    atoms: &[ComparisonUnitAtom],
    level: usize,
    is_inside_txbx: bool,
    settings: &WmlComparerSettings,
) -> Vec<(AdjacentGroupKey, Vec<ComparisonUnitAtom>)> {
    let mut groups: Vec<(AdjacentGroupKey, Vec<ComparisonUnitAtom>)> = Vec::new();
    
    for atom in atoms {
        let key = compute_adjacent_key(atom, level, is_inside_txbx, settings);
        
        match groups.last_mut() {
            Some((last_key, group_atoms)) if *last_key == key => {
                group_atoms.push(atom.clone());
            }
            _ => {
                groups.push((key, vec![atom.clone()]));
            }
        }
    }
    
    groups
}

fn compute_adjacent_key(
    atom: &ComparisonUnitAtom,
    level: usize,
    is_inside_txbx: bool,
    settings: &WmlComparerSettings,
) -> AdjacentGroupKey {
    let child_unid = if level + 1 < atom.ancestor_elements.len() {
        atom.ancestor_elements[level + 1].unid.clone()
    } else {
        String::new()
    };

    let (effective_unid, effective_status) = if is_inside_txbx && !child_unid.is_empty() {
        ("TXBX".to_string(), ComparisonCorrelationStatus::Equal)
    } else {
        (child_unid, atom.correlation_status)
    };

    let (formatting_sig, before_sig) = if settings.track_formatting_changes && !is_inside_txbx {
        match atom.correlation_status {
            ComparisonCorrelationStatus::FormatChanged => {
                (atom.formatting_signature.clone(), atom.normalized_rpr.clone())
            }
            ComparisonCorrelationStatus::Equal => {
                (atom.formatting_signature.clone(), None)
            }
            _ => (None, None),
        }
    } else {
        (None, None)
    };

    AdjacentGroupKey {
        child_unid: effective_unid,
        status: effective_status,
        formatting_sig,
        before_sig,
    }
}

fn is_inside_txbx_content(atom: &ComparisonUnitAtom, level: usize) -> bool {
    for i in 0..level {
        if let Some(ancestor) = atom.ancestor_elements.get(i) {
            if ancestor.local_name == "txbxContent" {
                return true;
            }
        }
    }
    false
}

fn is_inside_vml_content(atom: &ComparisonUnitAtom, level: usize) -> bool {
    for i in 0..=level {
        if let Some(ancestor) = atom.ancestor_elements.get(i) {
            if is_vml_related_element(&ancestor.local_name) {
                return true;
            }
        }
    }
    false
}

fn is_vml_related_element(name: &str) -> bool {
    matches!(
        name,
        "pict" | "shape" | "rect" | "group" | "shapetype" |
        "oval" | "line" | "arc" | "curve" | "polyline" | "roundrect"
    )
}

fn coalesce_paragraph_children(
    doc: &mut XmlDocument,
    para: NodeId,
    atoms: &[ComparisonUnitAtom],
    level: usize,
    settings: &WmlComparerSettings,
) {
    let first_atom = match atoms.first() {
        Some(a) => a,
        None => return,
    };
    let is_inside_txbx = is_inside_txbx_content(first_atom, level);
    let is_inside_vml = is_inside_vml_content(first_atom, level);
    
    let grouped = group_adjacent_by_key(atoms, level, is_inside_txbx, settings);
    
    for (key, group_atoms) in grouped {
        if key.child_unid.is_empty() {
            for atom in &group_atoms {
                if is_inside_vml
                    && matches!(&atom.content_element, ContentElement::ParagraphProperties)
                    && key.status == ComparisonCorrelationStatus::Inserted
                {
                    continue;
                }
                
                if let ContentElement::ParagraphProperties = &atom.content_element {
                    let ppr = doc.add_child(para, XmlNodeData::element(W::pPr()));
                    add_status_attribute(doc, ppr, key.status);
                }
            }
        } else {
            coalesce_recurse(doc, para, &group_atoms, level + 1, settings);
        }
    }
}

fn coalesce_run_children(
    doc: &mut XmlDocument,
    run: NodeId,
    atoms: &[ComparisonUnitAtom],
    level: usize,
    settings: &WmlComparerSettings,
) {
    let first_atom = match atoms.first() {
        Some(a) => a,
        None => return,
    };
    let is_inside_txbx = is_inside_txbx_content(first_atom, level);
    let is_inside_vml = is_inside_vml_content(first_atom, level);
    
    let grouped = group_adjacent_by_key(atoms, level, is_inside_txbx, settings);
    
    for (key, group_atoms) in grouped {
        if key.child_unid.is_empty() {
            coalesce_text_content(doc, run, &group_atoms, key.status, is_inside_vml);
        } else {
            coalesce_recurse(doc, run, &group_atoms, level + 1, settings);
        }
    }
}

fn coalesce_text_content(
    doc: &mut XmlDocument,
    parent: NodeId,
    atoms: &[ComparisonUnitAtom],
    status: ComparisonCorrelationStatus,
    is_inside_vml: bool,
) {
    let mut text_chars: Vec<char> = Vec::new();
    
    for atom in atoms {
        if is_inside_vml
            && matches!(&atom.content_element, ContentElement::ParagraphProperties)
            && atom.correlation_status == ComparisonCorrelationStatus::Inserted
        {
            continue;
        }
        
        match &atom.content_element {
            ContentElement::Text(ch) => {
                text_chars.push(*ch);
            }
            ContentElement::Break => {
                flush_text(doc, parent, &mut text_chars, status);
                let br = doc.add_child(parent, XmlNodeData::element(W::br()));
                add_status_attribute(doc, br, status);
            }
            ContentElement::Tab => {
                flush_text(doc, parent, &mut text_chars, status);
                let tab = doc.add_child(parent, XmlNodeData::element(W::tab()));
                add_status_attribute(doc, tab, status);
            }
            ContentElement::Drawing { .. } => {
                flush_text(doc, parent, &mut text_chars, status);
                let drawing = doc.add_child(parent, XmlNodeData::element(W::drawing()));
                add_status_attribute(doc, drawing, status);
            }
            ContentElement::ParagraphProperties => {
                flush_text(doc, parent, &mut text_chars, status);
                let ppr = doc.add_child(parent, XmlNodeData::element(W::pPr()));
                add_status_attribute(doc, ppr, status);
            }
            ContentElement::RunProperties => {
                flush_text(doc, parent, &mut text_chars, status);
                let rpr = doc.add_child(parent, XmlNodeData::element(W::rPr()));
                add_status_attribute(doc, rpr, status);
            }
            ContentElement::FootnoteReference { id } => {
                flush_text(doc, parent, &mut text_chars, status);
                let fnref = doc.add_child(parent, XmlNodeData::element_with_attrs(
                    W::footnoteReference(),
                    vec![XAttribute::new(W::id(), id)],
                ));
                add_status_attribute(doc, fnref, status);
            }
            ContentElement::EndnoteReference { id } => {
                flush_text(doc, parent, &mut text_chars, status);
                let enref = doc.add_child(parent, XmlNodeData::element_with_attrs(
                    W::endnoteReference(),
                    vec![XAttribute::new(W::id(), id)],
                ));
                add_status_attribute(doc, enref, status);
            }
            ContentElement::Math { .. } => {
                flush_text(doc, parent, &mut text_chars, status);
                let math = doc.add_child(parent, XmlNodeData::element(M::oMath()));
                add_status_attribute(doc, math, status);
            }
            _ => {}
        }
    }
    
    flush_text(doc, parent, &mut text_chars, status);
}

fn flush_text(
    doc: &mut XmlDocument,
    parent: NodeId,
    chars: &mut Vec<char>,
    status: ComparisonCorrelationStatus,
) {
    if chars.is_empty() {
        return;
    }
    
    let text: String = chars.drain(..).collect();
    let is_deleted = status == ComparisonCorrelationStatus::Deleted;
    let text_elem_name = if is_deleted { W::delText() } else { W::t() };
    
    let mut attrs = Vec::new();
    if needs_xml_space_preserve(&text) {
        attrs.push(XAttribute::new(
            XName::new("http://www.w3.org/XML/1998/namespace", "space"),
            "preserve",
        ));
    }
    
    let text_elem = if attrs.is_empty() {
        doc.add_child(parent, XmlNodeData::element(text_elem_name))
    } else {
        doc.add_child(parent, XmlNodeData::element_with_attrs(text_elem_name, attrs))
    };
    
    doc.add_child(text_elem, XmlNodeData::Text(text));
    add_status_attribute(doc, text_elem, status);
}

fn needs_xml_space_preserve(text: &str) -> bool {
    !text.is_empty()
        && (text.chars().next().map(|c| c.is_whitespace()).unwrap_or(false)
            || text.chars().last().map(|c| c.is_whitespace()).unwrap_or(false))
}



/// Add status attribute to an element if not Equal
fn add_status_attribute(doc: &mut XmlDocument, node: NodeId, status: ComparisonCorrelationStatus) {
    match status {
        ComparisonCorrelationStatus::Deleted => {
            doc.set_attribute(node, &pt_status(), "Deleted");
        }
        ComparisonCorrelationStatus::Inserted => {
            doc.set_attribute(node, &pt_status(), "Inserted");
        }
        _ => {}
    }
}

/// Create a paragraph element with Unid
fn create_paragraph(doc: &mut XmlDocument, parent: NodeId, unid: &str) -> NodeId {
    let attrs = vec![
        XAttribute::new(XName::new(PT_STATUS_NS, "Unid"), unid),
    ];
    doc.add_child(parent, XmlNodeData::element_with_attrs(W::p(), attrs))
}

/// Create a run element with Unid
fn create_run(doc: &mut XmlDocument, parent: NodeId, unid: &str) -> NodeId {
    let attrs = vec![
        XAttribute::new(XName::new(PT_STATUS_NS, "Unid"), unid),
    ];
    doc.add_child(parent, XmlNodeData::element_with_attrs(W::r(), attrs))
}

/// Create a table element
fn create_table(doc: &mut XmlDocument, parent: NodeId, unid: &str) -> NodeId {
    let attrs = vec![
        XAttribute::new(XName::new(PT_STATUS_NS, "Unid"), unid),
    ];
    doc.add_child(parent, XmlNodeData::element_with_attrs(W::tbl(), attrs))
}

/// Create a table row element
fn create_table_row(doc: &mut XmlDocument, parent: NodeId, unid: &str) -> NodeId {
    let attrs = vec![
        XAttribute::new(XName::new(PT_STATUS_NS, "Unid"), unid),
    ];
    doc.add_child(parent, XmlNodeData::element_with_attrs(W::tr(), attrs))
}

/// Create a table cell element
fn create_table_cell(doc: &mut XmlDocument, parent: NodeId, unid: &str) -> NodeId {
    let attrs = vec![
        XAttribute::new(XName::new(PT_STATUS_NS, "Unid"), unid),
    ];
    doc.add_child(parent, XmlNodeData::element_with_attrs(W::tc(), attrs))
}

/// Create a textbox content element
fn create_textbox_content(doc: &mut XmlDocument, parent: NodeId) -> NodeId {
    doc.add_child(parent, XmlNodeData::element(W::txbxContent()))
}

fn create_generic_element(doc: &mut XmlDocument, parent: NodeId, local_name: &str, unid: &str) -> NodeId {
    let name = XName::new(W::NS, local_name);
    let attrs = vec![
        XAttribute::new(XName::new(PT_STATUS_NS, "Unid"), unid),
    ];
    doc.add_child(parent, XmlNodeData::element_with_attrs(name, attrs))
}

// ============================================================================
// MarkContentAsDeletedOrInserted - Convert status markers to w:ins/w:del
// Port of C# WmlComparer.cs lines 2646-2740
// ============================================================================

use std::sync::atomic::{AtomicI32, Ordering};

static REVISION_ID: AtomicI32 = AtomicI32::new(0);

fn get_next_revision_id() -> i32 {
    REVISION_ID.fetch_add(1, Ordering::SeqCst)
}

pub fn reset_coalesce_revision_id(value: i32) {
    REVISION_ID.store(value, Ordering::SeqCst);
}

pub fn mark_content_as_deleted_or_inserted(
    doc: &mut XmlDocument,
    root: NodeId,
    settings: &WmlComparerSettings,
) {
    let nodes_to_transform: Vec<NodeId> = doc.descendants(root).collect();
    
    for node_id in nodes_to_transform {
        transform_node_for_revisions(doc, node_id, settings);
    }
    
    cleanup_pt_attributes(doc, root);
}

fn transform_node_for_revisions(
    doc: &mut XmlDocument,
    node_id: NodeId,
    settings: &WmlComparerSettings,
) {
    let node_name = {
        let data = match doc.get(node_id) {
            Some(d) => d,
            None => return,
        };
        match data.name() {
            Some(n) => n.clone(),
            None => return,
        }
    };
    
    if node_name == W::r() {
        transform_run_for_revisions(doc, node_id, settings);
    } else if node_name == W::pPr() {
        transform_ppr_for_revisions(doc, node_id, settings);
    }
}

fn transform_run_for_revisions(
    doc: &mut XmlDocument,
    run_id: NodeId,
    settings: &WmlComparerSettings,
) {
    let status = get_run_content_status(doc, run_id);
    
    match status.as_deref() {
        Some("Deleted") => {
            wrap_run_in_revision(doc, run_id, true, settings);
        }
        Some("Inserted") => {
            wrap_run_in_revision(doc, run_id, false, settings);
        }
        _ => {}
    }
}

fn get_run_content_status(doc: &XmlDocument, run_id: NodeId) -> Option<String> {
    for child_id in doc.children(run_id) {
        if let Some(data) = doc.get(child_id) {
            if let Some(attrs) = data.attributes() {
                for attr in attrs {
                    if attr.name == pt_status() {
                        return Some(attr.value.clone());
                    }
                }
            }
        }
    }
    None
}

fn wrap_run_in_revision(
    doc: &mut XmlDocument,
    run_id: NodeId,
    is_deletion: bool,
    settings: &WmlComparerSettings,
) {
    if doc.parent(run_id).is_none() {
        return;
    }
    
    let rev_name = if is_deletion { W::del() } else { W::ins() };
    let rev_id = get_next_revision_id();
    
    let author = settings.author_for_revisions.as_deref().unwrap_or("Unknown");
    let rev_attrs = vec![
        XAttribute::new(W::author(), author),
        XAttribute::new(W::id(), &rev_id.to_string()),
        XAttribute::new(W::date(), &settings.date_time_for_revisions),
    ];
    
    let rev_elem = doc.add_before(run_id, XmlNodeData::element_with_attrs(rev_name, rev_attrs));
    
    doc.reparent(rev_elem, run_id);
}

fn transform_ppr_for_revisions(
    doc: &mut XmlDocument,
    ppr_id: NodeId,
    settings: &WmlComparerSettings,
) {
    let status = {
        if let Some(data) = doc.get(ppr_id) {
            if let Some(attrs) = data.attributes() {
                attrs.iter()
                    .find(|a| a.name == pt_status())
                    .map(|a| a.value.clone())
            } else {
                None
            }
        } else {
            None
        }
    };
    
    if let Some(status_val) = status {
        let rpr_id = get_or_create_rpr(doc, ppr_id);
        let rev_name = if status_val == "Deleted" { W::del() } else { W::ins() };
        let rev_id = get_next_revision_id();
        let author = settings.author_for_revisions.as_deref().unwrap_or("Unknown");
        
        let rev_attrs = vec![
            XAttribute::new(W::author(), author),
            XAttribute::new(W::id(), &rev_id.to_string()),
            XAttribute::new(W::date(), &settings.date_time_for_revisions),
        ];
        
        doc.add_child(rpr_id, XmlNodeData::element_with_attrs(rev_name, rev_attrs));
    }
}

fn get_or_create_rpr(doc: &mut XmlDocument, ppr_id: NodeId) -> NodeId {
    for child_id in doc.children(ppr_id) {
        if let Some(data) = doc.get(child_id) {
            if data.name() == Some(&W::rPr()) {
                return child_id;
            }
        }
    }
    doc.add_child(ppr_id, XmlNodeData::element(W::rPr()))
}

fn cleanup_pt_attributes(doc: &mut XmlDocument, root: NodeId) {
    let pt_status_name = pt_status();
    let pt_unid_name = XName::new(PT_STATUS_NS, "Unid");
    
    let nodes: Vec<NodeId> = doc.descendants(root).collect();
    for node_id in nodes {
        doc.remove_attribute(node_id, &pt_status_name);
        doc.remove_attribute(node_id, &pt_unid_name);
    }
}

// ============================================================================
// CoalesceAdjacentRunsWithIdenticalFormatting
// Port of C# PtOpenXmlUtil.cs lines 799-991
// ============================================================================

use crate::util::group::group_adjacent;
use crate::util::descendants::descendants_trimmed;

/// Coalesce adjacent runs in an entire document
///
/// This is the top-level function that processes all paragraphs in a document,
/// matching the C# CoalesceAdjacentRunsWithIdenticalFormatting(XDocument) signature.
///
/// CRITICAL: This must be called AFTER mark_content_as_deleted_or_inserted to match C# behavior.
pub fn coalesce_document(doc: &mut XmlDocument, doc_root: NodeId) {
    // Process main document paragraphs (excluding those in txbxContent)
    // This matches C# line 2338: var paras = xDoc.Root.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.p);
    let paras: Vec<NodeId> = descendants_trimmed(doc, doc_root, |d| {
        d.name().map(|n| n == &W::txbxContent()).unwrap_or(false)
    })
    .filter(|&node| {
        doc.get(node)
            .and_then(|d| d.name())
            .map(|n| n == &W::p())
            .unwrap_or(false)
    })
    .collect();
    
    for para in paras {
        coalesce_adjacent_runs_with_identical_formatting(doc, para);
    }
    
    // Process txbxContent paragraphs recursively (C# lines 2344-2351)
    let txbx_containers: Vec<NodeId> = doc
        .descendants(doc_root)
        .filter(|&node| {
            doc.get(node)
                .and_then(|d| d.name())
                .map(|n| n == &W::txbxContent())
                .unwrap_or(false)
        })
        .collect();
    
    for txbx in txbx_containers {
        let txbx_paras: Vec<NodeId> = descendants_trimmed(doc, txbx, |d| {
            d.name().map(|n| n == &W::txbxContent()).unwrap_or(false)
        })
        .filter(|&node| {
            doc.get(node)
                .and_then(|d| d.name())
                .map(|n| n == &W::p())
                .unwrap_or(false)
        })
        .collect();
        
        for para in txbx_paras {
            coalesce_adjacent_runs_with_identical_formatting(doc, para);
        }
    }
}

/// Coalesce adjacent runs with identical formatting in a single container
///
/// This ports the C# CoalesceAdjacentRunsWithIdenticalFormatting function which
/// consolidates adjacent w:r, w:ins, and w:del elements that have identical formatting.
///
/// CRITICAL: This must be called AFTER mark_content_as_deleted_or_inserted to match C# behavior.
pub fn coalesce_adjacent_runs_with_identical_formatting(
    doc: &mut XmlDocument,
    run_container: NodeId,
) {
    const DONT_CONSOLIDATE: &str = "DontConsolidate";
    
    // Get all direct children
    let children: Vec<NodeId> = doc.children(run_container).collect();
    
    // Group adjacent runs by their grouping key
    let groups = group_adjacent(children.into_iter(), |&child| {
        compute_run_grouping_key(doc, child)
    });
    
    // Process each group
    for group in groups {
        if group.is_empty() {
            continue;
        }
        
        // Get the grouping key for this group
        let key = compute_run_grouping_key(doc, group[0]);
        
        if key == DONT_CONSOLIDATE || group.len() == 1 {
            // Don't consolidate this group - keep as is
            continue;
        }
        
        // Consolidate this group
        consolidate_run_group(doc, run_container, &group, &key);
    }
    
    // Recursively process w:txbxContent//w:p
    process_textbox_content(doc, run_container);
    
    // Process additional run containers
    process_additional_run_containers(doc, run_container);
}

/// Compute the grouping key for a run element
fn compute_run_grouping_key(doc: &XmlDocument, node: NodeId) -> String {
    const DONT_CONSOLIDATE: &str = "DontConsolidate";
    
    let node_data = match doc.get(node) {
        Some(d) => d,
        None => return DONT_CONSOLIDATE.to_string(),
    };
    
    let node_name = match node_data.name() {
        Some(n) => n,
        None => return DONT_CONSOLIDATE.to_string(),
    };
    
    // w:r handling
    if node_name == &W::r() {
        return compute_run_key(doc, node);
    }
    
    // w:ins handling
    if node_name == &W::ins() {
        return compute_ins_key(doc, node);
    }
    
    // w:del handling
    if node_name == &W::del() {
        return compute_del_key(doc, node);
    }
    
    DONT_CONSOLIDATE.to_string()
}

/// Compute grouping key for w:r element
fn compute_run_key(doc: &XmlDocument, run: NodeId) -> String {
    const DONT_CONSOLIDATE: &str = "DontConsolidate";
    
    // Count non-rPr children
    let non_rpr_count = doc.children(run)
        .filter(|&child| {
            doc.get(child)
                .and_then(|d| d.name())
                .map(|n| n != &W::rPr())
                .unwrap_or(false)
        })
        .count();
    
    if non_rpr_count != 1 {
        return DONT_CONSOLIDATE.to_string();
    }
    
    // Check for AbstractNumId attribute (C# line 813)
    if has_abstract_num_id_attribute(doc, run) {
        return DONT_CONSOLIDATE.to_string();
    }
    
    // Get rPr string
    let rpr_string = get_rpr_string(doc, run);
    
    // Check for w:t child
    if has_child_element(doc, run, &W::t()) {
        return format!("Wt{}", rpr_string);
    }
    
    // Check for w:instrText child
    if has_child_element(doc, run, &W::instrText()) {
        return format!("WinstrText{}", rpr_string);
    }
    
    DONT_CONSOLIDATE.to_string()
}

/// Compute grouping key for w:ins element
fn compute_ins_key(doc: &XmlDocument, ins: NodeId) -> String {
    const DONT_CONSOLIDATE: &str = "DontConsolidate";
    
    // Check for nested w:del (C# line 830-832)
    if has_child_element(doc, ins, &W::del()) {
        return DONT_CONSOLIDATE.to_string();
    }
    
    // Must be w:ins/w:r/w:t pattern (C# line 866-868)
    let has_valid_structure = doc.children(ins).all(|run_id| {
        let run_children: Vec<_> = doc.children(run_id)
            .filter(|&child| {
                doc.get(child)
                    .and_then(|d| d.name())
                    .map(|n| n != &W::rPr())
                    .unwrap_or(false)
            })
            .collect();
        
        run_children.len() == 1 && has_child_element(doc, run_id, &W::t())
    });
    
    if !has_valid_structure {
        return DONT_CONSOLIDATE.to_string();
    }
    
    // Get author, date, and id (C# lines 870-882)
    let author = get_attribute_value(doc, ins, &W::author()).unwrap_or_default();
    let date_str = get_date_attribute_iso(doc, ins);
    let id_str = get_attribute_value(doc, ins, &W::id()).unwrap_or_default();
    
    // Get rPr strings from all w:r children
    let rpr_strings: String = doc.children(ins)
        .filter_map(|run_id| {
            doc.children(run_id)
                .find(|&child| {
                    doc.get(child)
                        .and_then(|d| d.name())
                        .map(|n| n == &W::rPr())
                        .unwrap_or(false)
                })
                .map(|rpr_id| serialize_element(doc, rpr_id))
        })
        .collect();
    
    format!("Wins2{}{}{}{}", author, date_str, id_str, rpr_strings)
}

/// Compute grouping key for w:del element  
fn compute_del_key(doc: &XmlDocument, del: NodeId) -> String {
    const DONT_CONSOLIDATE: &str = "DontConsolidate";
    
    // Must be w:del/w:r/w:delText pattern (C# line 891-893)
    let has_valid_structure = doc.children(del).all(|run_id| {
        let run_children: Vec<_> = doc.children(run_id)
            .filter(|&child| {
                doc.get(child)
                    .and_then(|d| d.name())
                    .map(|n| n != &W::rPr())
                    .unwrap_or(false)
            })
            .collect();
        
        run_children.len() == 1 && has_child_element(doc, run_id, &W::delText())
    });
    
    if !has_valid_structure {
        return DONT_CONSOLIDATE.to_string();
    }
    
    // Get author and date (C# lines 895-898) - NOTE: NO ID for deletions!
    let author = get_attribute_value(doc, del, &W::author()).unwrap_or_default();
    let date_str = get_date_attribute_iso(doc, del);
    
    // Get rPr strings from all w:r children
    let rpr_strings: String = doc.children(del)
        .filter_map(|run_id| {
            doc.children(run_id)
                .find(|&child| {
                    doc.get(child)
                        .and_then(|d| d.name())
                        .map(|n| n == &W::rPr())
                        .unwrap_or(false)
                })
                .map(|rpr_id| serialize_element(doc, rpr_id))
        })
        .collect();
    
    format!("Wdel{}{}{}", author, date_str, rpr_strings)
}

/// Consolidate a group of runs with identical formatting
fn consolidate_run_group(
    doc: &mut XmlDocument,
    parent: NodeId,
    group: &[NodeId],
    key: &str,
) {
    if group.is_empty() {
        return;
    }
    
    let first = group[0];
    let first_name = doc.get(first).and_then(|d| d.name()).cloned();
    
    // Collect all text values from the group
    let text_value = collect_text_from_group(doc, group);
    
    // Create consolidated element based on type
    match first_name.as_ref() {
        Some(name) if name == &W::r() => {
            consolidate_run_group_r(doc, parent, group, &text_value, key);
        }
        Some(name) if name == &W::ins() => {
            consolidate_run_group_ins(doc, parent, group, &text_value);
        }
        Some(name) if name == &W::del() => {
            consolidate_run_group_del(doc, parent, group, &text_value);
        }
        _ => {}
    }
}

/// Consolidate w:r group
fn consolidate_run_group_r(
    doc: &mut XmlDocument,
    parent: NodeId,
    group: &[NodeId],
    text_value: &str,
    key: &str,
) {
    let first = group[0];
    
    // Determine element type from key
    let is_instr_text = key.starts_with("WinstrText");
    
    // Clone first run's attributes
    let attrs = doc.get(first)
        .and_then(|d| d.attributes())
        .map(|a| a.to_vec())
        .unwrap_or_default();
    
    // Create new consolidated run
    let new_run = doc.add_child(parent, XmlNodeData::element_with_attrs(W::r(), attrs));
    
    // Copy rPr from first run
    if let Some(rpr_id) = find_child_element(doc, first, &W::rPr()) {
        clone_element_into(doc, rpr_id, new_run);
    }
    
    // Create text element with xml:space if needed
    let text_elem_name = if is_instr_text { W::instrText() } else { W::t() };
    let mut text_attrs = Vec::new();
    
    if needs_xml_space_preserve(text_value) {
        text_attrs.push(XAttribute::new(
            XName::new("http://www.w3.org/XML/1998/namespace", "space"),
            "preserve",
        ));
    }
    
    // Collect pt:Status attributes from w:t elements in the group (C# line 933)
    let status_attrs = collect_status_attributes_from_group(doc, group);
    text_attrs.extend(status_attrs);
    
    let text_elem = if text_attrs.is_empty() {
        doc.add_child(new_run, XmlNodeData::element(text_elem_name))
    } else {
        doc.add_child(new_run, XmlNodeData::element_with_attrs(text_elem_name, text_attrs))
    };
    
    doc.add_child(text_elem, XmlNodeData::Text(text_value.to_string()));
    
    for &old_run in group {
        doc.detach(old_run);
    }
}

/// Consolidate w:ins group
fn consolidate_run_group_ins(
    doc: &mut XmlDocument,
    parent: NodeId,
    group: &[NodeId],
    text_value: &str,
) {
    let first = group[0];
    
    // Clone first ins's attributes
    let attrs = doc.get(first)
        .and_then(|d| d.attributes())
        .map(|a| a.to_vec())
        .unwrap_or_default();
    
    // Create new consolidated ins
    let new_ins = doc.add_child(parent, XmlNodeData::element_with_attrs(W::ins(), attrs));
    
    // Get first run's attributes
    let first_run = doc.children(first).next();
    let run_attrs = first_run
        .and_then(|r| doc.get(r).and_then(|d| d.attributes()).map(|a| a.to_vec()))
        .unwrap_or_default();
    
    let new_run = doc.add_child(new_ins, XmlNodeData::element_with_attrs(W::r(), run_attrs));
    
    // Copy rPr from first run
    if let Some(first_r) = first_run {
        if let Some(rpr_id) = find_child_element(doc, first_r, &W::rPr()) {
            clone_element_into(doc, rpr_id, new_run);
        }
    }
    
    // Create text element with xml:space if needed
    let mut text_attrs = Vec::new();
    if needs_xml_space_preserve(text_value) {
        text_attrs.push(XAttribute::new(
            XName::new("http://www.w3.org/XML/1998/namespace", "space"),
            "preserve",
        ));
    }
    
    let text_elem = if text_attrs.is_empty() {
        doc.add_child(new_run, XmlNodeData::element(W::t()))
    } else {
        doc.add_child(new_run, XmlNodeData::element_with_attrs(W::t(), text_attrs))
    };
    
    doc.add_child(text_elem, XmlNodeData::Text(text_value.to_string()));
    
    for &old_ins in group {
        doc.detach(old_ins);
    }
}

/// Consolidate w:del group
fn consolidate_run_group_del(
    doc: &mut XmlDocument,
    parent: NodeId,
    group: &[NodeId],
    text_value: &str,
) {
    let first = group[0];
    
    // Clone first del's attributes
    let attrs = doc.get(first)
        .and_then(|d| d.attributes())
        .map(|a| a.to_vec())
        .unwrap_or_default();
    
    // Create new consolidated del
    let new_del = doc.add_child(parent, XmlNodeData::element_with_attrs(W::del(), attrs));
    
    // Get first run's attributes
    let first_run = doc.children(first).next();
    let run_attrs = first_run
        .and_then(|r| doc.get(r).and_then(|d| d.attributes()).map(|a| a.to_vec()))
        .unwrap_or_default();
    
    let new_run = doc.add_child(new_del, XmlNodeData::element_with_attrs(W::r(), run_attrs));
    
    // Copy rPr from first run
    if let Some(first_r) = first_run {
        if let Some(rpr_id) = find_child_element(doc, first_r, &W::rPr()) {
            clone_element_into(doc, rpr_id, new_run);
        }
    }
    
    // Create delText element with xml:space if needed
    let mut text_attrs = Vec::new();
    if needs_xml_space_preserve(text_value) {
        text_attrs.push(XAttribute::new(
            XName::new("http://www.w3.org/XML/1998/namespace", "space"),
            "preserve",
        ));
    }
    
    let text_elem = if text_attrs.is_empty() {
        doc.add_child(new_run, XmlNodeData::element(W::delText()))
    } else {
        doc.add_child(new_run, XmlNodeData::element_with_attrs(W::delText(), text_attrs))
    };
    
    doc.add_child(text_elem, XmlNodeData::Text(text_value.to_string()));
    
    for &old_del in group {
        doc.detach(old_del);
    }
}

/// Collect text from all elements in group (C# lines 919-925)
fn collect_text_from_group(doc: &XmlDocument, group: &[NodeId]) -> String {
    group
        .iter()
        .flat_map(|&node| {
            doc.descendants(node)
                .filter_map(|desc| {
                    doc.get(desc).and_then(|d| d.name()).and_then(|n| {
                        if n == &W::t() || n == &W::delText() || n == &W::instrText() {
                            doc.children(desc)
                                .filter_map(|text_node| {
                                    doc.get(text_node).and_then(|d| d.text_content()).map(|s| s.to_string())
                                })
                                .collect::<Vec<_>>()
                                .join("")
                                .into()
                        } else {
                            None
                        }
                    })
                })
                .collect::<Vec<_>>()
        })
        .collect::<Vec<_>>()
        .join("")
}

/// Collect pt:Status attributes from w:t descendants (C# line 933)
fn collect_status_attributes_from_group(doc: &XmlDocument, group: &[NodeId]) -> Vec<XAttribute> {
    let pt_status_name = pt_status();
    
    group
        .iter()
        .flat_map(|&node| {
            doc.descendants(node)
                .filter(|&desc| {
                    doc.get(desc)
                        .and_then(|d| d.name())
                        .map(|n| n == &W::t())
                        .unwrap_or(false)
                })
                .take(1)  // Take first w:t only
                .filter_map(|t_elem| {
                    doc.get(t_elem)
                        .and_then(|d| d.attributes())
                        .and_then(|attrs| {
                            attrs
                                .iter()
                                .find(|a| a.name == pt_status_name)
                                .cloned()
                        })
                })
                .collect::<Vec<_>>()
        })
        .collect()
}

/// Process w:txbxContent//w:p recursively (C# lines 971-977)
fn process_textbox_content(doc: &mut XmlDocument, container: NodeId) {
    let txbx_paras: Vec<NodeId> = descendants_trimmed(doc, container, |d| {
        d.name().map(|n| n == &W::txbxContent()).unwrap_or(false)
    })
    .filter(|&node| {
        doc.get(node)
            .and_then(|d| d.name())
            .map(|n| n == &W::p())
            .unwrap_or(false)
    })
    .collect();
    
    for para in txbx_paras {
        coalesce_adjacent_runs_with_identical_formatting(doc, para);
    }
}

/// Process additional run containers (C# lines 979-988)
fn process_additional_run_containers(doc: &mut XmlDocument, container: NodeId) {
    // Additional run container names from C# line 787-797
    let additional_names = [
        W::bdo(),
        W::customXml(),
        W::dir(),
        W::fldSimple(),
        W::hyperlink(),
        W::moveFrom(),
        W::moveTo(),
        W::sdtContent(),
    ];
    
    let containers: Vec<NodeId> = doc
        .descendants(container)
        .filter(|&node| {
            doc.get(node)
                .and_then(|d| d.name())
                .map(|n| additional_names.contains(n))
                .unwrap_or(false)
        })
        .collect();
    
    for cont in containers {
        coalesce_adjacent_runs_with_identical_formatting(doc, cont);
    }
}

// Helper functions

fn has_abstract_num_id_attribute(doc: &XmlDocument, node: NodeId) -> bool {
    doc.get(node)
        .and_then(|d| d.attributes())
        .map(|attrs| {
            attrs.iter().any(|a| a.name.local_name == "AbstractNumId")
        })
        .unwrap_or(false)
}

fn get_rpr_string(doc: &XmlDocument, run: NodeId) -> String {
    find_child_element(doc, run, &W::rPr())
        .map(|rpr| serialize_element(doc, rpr))
        .unwrap_or_default()
}

fn has_child_element(doc: &XmlDocument, parent: NodeId, name: &XName) -> bool {
    doc.children(parent).any(|child| {
        doc.get(child)
            .and_then(|d| d.name())
            .map(|n| n == name)
            .unwrap_or(false)
    })
}

fn find_child_element(doc: &XmlDocument, parent: NodeId, name: &XName) -> Option<NodeId> {
    doc.children(parent).find(|&child| {
        doc.get(child)
            .and_then(|d| d.name())
            .map(|n| n == name)
            .unwrap_or(false)
    })
}

fn get_attribute_value(doc: &XmlDocument, node: NodeId, attr_name: &XName) -> Option<String> {
    doc.get(node)
        .and_then(|d| d.attributes())
        .and_then(|attrs| {
            attrs
                .iter()
                .find(|a| &a.name == attr_name)
                .map(|a| a.value.clone())
        })
}

fn get_date_attribute_iso(doc: &XmlDocument, node: NodeId) -> String {
    get_attribute_value(doc, node, &W::date())
        .map(|date_str| {
            // C# converts DateTime to ISO format (line 874, 898)
            // For now, we'll pass through the existing value
            // TODO: Parse and format if needed for exact C# compatibility
            date_str
        })
        .unwrap_or_default()
}

fn serialize_element(doc: &XmlDocument, node: NodeId) -> String {
    // Simplified serialization - in production, this should match C# ToString(SaveOptions.None)
    // For now, we'll use element name as a proxy
    doc.get(node)
        .and_then(|d| d.name())
        .map(|n| format!("<{}/>", n.local_name))
        .unwrap_or_default()
}

fn clone_element_into(doc: &mut XmlDocument, source: NodeId, dest_parent: NodeId) {
    // Clone source element and all its descendants into dest_parent
    let source_data = match doc.get(source) {
        Some(d) => d.clone(),
        None => return,
    };
    
    let new_node = doc.add_child(dest_parent, source_data);
    
    // Recursively clone children
    let children: Vec<NodeId> = doc.children(source).collect();
    for child in children {
        clone_element_into(doc, child, new_node);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::wml::comparison_unit::AncestorInfo;
    
    #[test]
    fn test_group_by_ancestor_unid() {
        let atoms = vec![
            create_test_atom("p1", "Text1"),
            create_test_atom("p1", "Text2"),
            create_test_atom("p2", "Text3"),
        ];
        
        let groups = group_by_ancestor_unid(&atoms, 0);
        
        assert_eq!(groups.len(), 2);
        assert_eq!(groups[0].0, "p1");
        assert_eq!(groups[0].1.len(), 2);
        assert_eq!(groups[1].0, "p2");
        assert_eq!(groups[1].1.len(), 1);
    }
    
    fn create_test_atom(para_unid: &str, _text: &str) -> ComparisonUnitAtom {
        use crate::xml::arena::XmlDocument;
        let mut doc = XmlDocument::new();
        let dummy = doc.add_root(XmlNodeData::Text("dummy".to_string()));
        
        ComparisonUnitAtom {
            content_element: ContentElement::Text('a'),
            sha1_hash: "test".to_string(),
            ancestor_elements: vec![
                AncestorInfo {
                    node_id: dummy,
                    local_name: "p".to_string(),
                    unid: para_unid.to_string(),
                },
            ],
            correlation_status: ComparisonCorrelationStatus::Equal,
            formatting_signature: None,
            normalized_rpr: None,
            part_name: "main".to_string(),
        }
    }
    
    #[test]
    fn test_needs_xml_space_preserve() {
        assert!(needs_xml_space_preserve(" hello"));
        assert!(needs_xml_space_preserve("hello "));
        assert!(needs_xml_space_preserve(" hello "));
        assert!(!needs_xml_space_preserve("hello"));
        assert!(!needs_xml_space_preserve("hello world"));
    }
    
    #[test]
    fn test_is_vml_related() {
        assert!(is_vml_related_element("pict"));
        assert!(is_vml_related_element("shape"));
        assert!(!is_vml_related_element("p"));
    }
    
    #[test]
    fn test_compute_run_key_simple() {
        let mut doc = XmlDocument::new();
        let run = doc.add_root(XmlNodeData::element(W::r()));
        let _t = doc.add_child(run, XmlNodeData::element(W::t()));
        
        let key = compute_run_key(&doc, run);
        assert!(key.starts_with("Wt"));
    }
}
