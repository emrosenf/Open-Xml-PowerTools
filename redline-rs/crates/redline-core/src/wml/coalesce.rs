//! Coalesce - Reconstruct XML tree from comparison atoms
//!
//! This is a faithful line-by-line port of CoalesceRecurse from C# WmlComparer.cs (lines 5161-5738).

use super::comparison_unit::{ComparisonCorrelationStatus, ComparisonUnitAtom, ContentElement};
use super::settings::WmlComparerSettings;
use crate::wml::revision::{create_run_property_change, RevisionSettings};
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{W, W16DU};
use crate::xml::node::XmlNodeData;
use crate::xml::parser::parse;
use crate::xml::xname::{XAttribute, XName};
use indextree::NodeId;
use std::collections::HashMap;

/// PowerTools namespace for internal tracking attributes
pub const PT_STATUS_NS: &str = "http://powertools.codeplex.com/2011";

/// Create the pt:Status attribute name
pub fn pt_status() -> XName {
    XName::new(PT_STATUS_NS, "Status")
}

/// Create the pt:Unid attribute name
pub fn pt_unid() -> XName {
    XName::new(PT_STATUS_NS, "Unid")
}

pub fn strip_pt_attributes(doc: &mut XmlDocument, root: NodeId) {
    let nodes: Vec<_> = std::iter::once(root).chain(doc.descendants(root)).collect();
    for node in nodes {
        let Some(node_data) = doc.get_mut(node) else { continue; };
        let Some(attrs) = node_data.attributes_mut() else { continue; };
        attrs.retain(|attr| attr.name.namespace.as_deref() != Some(PT_STATUS_NS));
    }
}

/// Helper struct to represent an ancestor element's information
#[derive(Clone, Debug)]
struct AncestorElementInfo {
    namespace: Option<String>,
    local_name: String,
    attributes: Vec<XAttribute>,
    /// Serialized run properties (w:rPr) for run elements (w:r)
    /// This is needed to preserve formatting when reconstructing runs
    rpr_xml: Option<String>,
}

/// VML-related element names (from C# VmlRelatedElements set)
static VML_RELATED_ELEMENTS: &[&str] = &[
    "pict", "shape", "rect", "group", "shapetype", "oval", "line", "arc", "curve", "polyline", "roundrect",
];

/// Allowable run children that can have pt:Status
static ALLOWABLE_RUN_CHILDREN: &[&str] = &[
    "br", "tab", "sym", "ptab", "cr", "dayShort", "dayLong", "monthShort", "monthLong", "yearShort", "yearLong",
    "footnoteReference", "endnoteReference", "drawing", "object",
];

pub struct CoalesceResult {
    pub document: XmlDocument,
    pub root: NodeId,
}

pub fn coalesce(atoms: &[ComparisonUnitAtom], settings: &WmlComparerSettings, root_name: XName, root_attrs: Vec<XAttribute>) -> CoalesceResult {
    let mut doc = XmlDocument::new();
    
    let mut attrs = root_attrs;
    // Ensure standard namespaces are present
    let standard_namespaces = vec![
        ("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        ("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
        ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        ("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
        ("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
        ("w14", "http://schemas.microsoft.com/office/word/2010/wordml"),
        // Word 2023 Date UTC namespace for revision timestamps
        ("w16du", W16DU::NS),
    ];
    for (prefix, uri) in standard_namespaces {
        if !attrs.iter().any(|a| a.name.local_name == prefix && a.name.namespace.as_deref() == Some("http://www.w3.org/2000/xmlns/")) {
            attrs.push(XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", prefix), uri));
        }
    }
    // Note: Do NOT add pt14 namespace - it's for internal tracking only, not output
    
    let doc_root = doc.add_root(XmlNodeData::element_with_attrs(root_name.clone(), attrs));
    
    if root_name.local_name == "document" {
        let body = doc.add_child(doc_root, XmlNodeData::element(W::body()));
        coalesce_recurse(&mut doc, body, atoms, 0, None, settings);
        move_last_sect_pr_to_child_of_body(&mut doc, doc_root);
    } else {
        coalesce_recurse(&mut doc, doc_root, atoms, 0, None, settings);
    }
    
    // NOTE: Do NOT remove empty rPr here - other functions will create more after this
    // The cleanup happens at the very end of the processing pipeline in comparer.rs
    
    CoalesceResult { document: doc, root: doc_root }
}

/// Remove empty w:rPr elements from the document tree.
/// Empty w:rPr elements (those with no children) are non-standard OOXML
/// and may cause MS Word to reject the file.
pub fn remove_empty_rpr_elements(doc: &mut XmlDocument, root: NodeId) {
    // Collect all empty rPr elements first to avoid borrowing issues
    let empty_rprs: Vec<NodeId> = std::iter::once(root)
        .chain(doc.descendants(root))
        .filter(|&node| {
            // Check if this is a w:rPr element
            let is_rpr = doc.get(node)
                .and_then(|d| d.name())
                .map(|n| n.namespace.as_deref() == Some(W::NS) && n.local_name == "rPr")
                .unwrap_or(false);
            
            // Check if it has no children
            is_rpr && doc.children(node).next().is_none()
        })
        .collect();
    
    // Remove all empty rPr elements
    for node in empty_rprs {
        doc.remove(node);
    }
}

/// Mark content as deleted or inserted - faithful port of C# MarkContentAsDeletedOrInserted (line 2646-2740)
/// 
/// This uses a recursive transformation pattern matching the C# algorithm exactly.
pub fn mark_content_as_deleted_or_inserted(
    doc: &mut XmlDocument,
    root: NodeId,
    settings: &WmlComparerSettings,
) {
    let revision_settings = RevisionSettings {
        author: settings.author_for_revisions.clone().unwrap_or_else(|| "Unknown".to_string()),
        date_time: settings.date_time_for_revisions.clone().unwrap_or_else(|| chrono::Utc::now().format("%Y-%m-%dT%H:%M:%SZ").to_string()),
    };

    // Transform in-place following C# pattern
    mark_content_transform(doc, root, &revision_settings);
}

/// Static counter for revision IDs (matching C# s_MaxId)
/// NOTE: Starts at 0 to match gold standard (w:id="0", "1", "2"...)
static REVISION_ID: std::sync::atomic::AtomicU32 = std::sync::atomic::AtomicU32::new(0);

fn next_revision_id() -> u32 {
    REVISION_ID.fetch_add(1, std::sync::atomic::Ordering::SeqCst)
}

/// Reset revision ID counter (for testing)
#[allow(dead_code)]
pub fn reset_revision_id() {
    REVISION_ID.store(0, std::sync::atomic::Ordering::SeqCst);
}

/// Recursive transform matching C# MarkContentAsDeletedOrInsertedTransform (line 2652-2740)
fn mark_content_transform(doc: &mut XmlDocument, node: NodeId, settings: &RevisionSettings) {
    let Some(data) = doc.get(node) else { return };
    let Some(name) = data.name().cloned() else { return };
    
    // Check if this is a w:r element (C# line 2657)
    if name.namespace.as_deref() == Some(W::NS) && name.local_name == "r" {
        handle_run_element(doc, node, settings);
        return;
    }
    
    // Check if this is a w:pPr element (C# line 2694)
    if name.namespace.as_deref() == Some(W::NS) && name.local_name == "pPr" {
        handle_ppr_element(doc, node, settings);
        return;
    }

    // Check if this is a math element (oMath or oMathPara)
    if name.namespace.as_deref() == Some("http://schemas.openxmlformats.org/officeDocument/2006/math") 
       && (name.local_name == "oMath" || name.local_name == "oMathPara") {
        handle_omath_element(doc, node, settings);
        return;
    }
    
    // Otherwise, recurse into children (C# line 2735-2737)
    let children: Vec<_> = doc.children(node).collect();
    for child in children {
        mark_content_transform(doc, child, settings);
    }
}

/// Handle w:r elements - C# lines 2657-2691
fn handle_run_element(doc: &mut XmlDocument, run: NodeId, settings: &RevisionSettings) {
    // Get status from descendants (w:t, w:delText, or AllowableRunChildren)
    // C# lines 2659-2665: DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.t || d.Name == W.delText || AllowableRunChildren.Contains(d.Name))
    let status_list = get_run_descendant_statuses(doc, run);
    
    if status_list.len() > 1 {
        // C# line 2667: throw new OpenXmlPowerToolsException("Internal error - have both deleted and inserted text elements in the same run.");
        // In Rust, we'll just log a warning and proceed with the first status
        eprintln!("Warning: have both deleted and inserted text elements in the same run");
    }
    
    if status_list.is_empty() {
        // C# lines 2668-2671: No status, just recurse into children
        let children: Vec<_> = doc.children(run).collect();
        for child in children {
            mark_content_transform(doc, child, settings);
        }
        return;
    }
    
    let status = &status_list[0];
    
    // Remove pt:Status from all descendants before wrapping
    let descendants: Vec<_> = doc.descendants(run).collect();
    for desc in descendants {
        doc.remove_attribute(desc, &pt_status());
    }
    
    // Wrap the run in w:del or w:ins (C# lines 2672-2691)
    if status == "Deleted" {
        let id_str = next_revision_id().to_string();
        let del_attrs = vec![
            XAttribute::new(W::id(), &id_str),  // w:id MUST come first per ECMA-376
            XAttribute::new(W::author(), &settings.author),
            XAttribute::new(W::date(), &settings.date_time),
            // Add w16du:dateUtc for modern Word timezone handling
            XAttribute::new(W16DU::dateUtc(), &settings.date_time),
        ];
        let del_elem = doc.new_node(XmlNodeData::element_with_attrs(W::del(), del_attrs));
        
        doc.insert_before(run, del_elem);
        doc.detach(run);
        doc.reparent(del_elem, run);
    } else if status == "Inserted" {
        let id_str = next_revision_id().to_string();
        let ins_attrs = vec![
            XAttribute::new(W::id(), &id_str),  // w:id MUST come first per ECMA-376
            XAttribute::new(W::author(), &settings.author),
            XAttribute::new(W::date(), &settings.date_time),
            // Add w16du:dateUtc for modern Word timezone handling
            XAttribute::new(W16DU::dateUtc(), &settings.date_time),
        ];
        let ins_elem = doc.new_node(XmlNodeData::element_with_attrs(W::ins(), ins_attrs));
        
        doc.insert_before(run, ins_elem);
        doc.detach(run);
        doc.reparent(ins_elem, run);
    }
}

/// Get distinct status values from run descendants - C# lines 2659-2665
fn get_run_descendant_statuses(doc: &XmlDocument, run: NodeId) -> Vec<String> {
    let mut statuses = Vec::new();
    collect_run_descendant_statuses(doc, run, &mut statuses, false);
    statuses.sort();
    statuses.dedup();
    statuses
}

/// Collect status values from descendants, trimming at w:txbxContent
fn collect_run_descendant_statuses(doc: &XmlDocument, node: NodeId, statuses: &mut Vec<String>, skip_txbx: bool) {
    for child in doc.children(node) {
        let Some(data) = doc.get(child) else { continue };
        let Some(name) = data.name() else { continue };
        
        // Skip w:txbxContent and its descendants (C# DescendantsTrimmed(W.txbxContent))
        if name.namespace.as_deref() == Some(W::NS) && name.local_name == "txbxContent" {
            continue;
        }
        
        // Check if this is w:t, w:delText, or AllowableRunChildren
        let is_target = name.namespace.as_deref() == Some(W::NS) && 
            (name.local_name == "t" || name.local_name == "delText" || 
             ALLOWABLE_RUN_CHILDREN.contains(&name.local_name.as_str()));
        
        if is_target {
            if let Some(attrs) = data.attributes() {
                if let Some(attr) = attrs.iter().find(|a| a.name == pt_status()) {
                    statuses.push(attr.value.clone());
                }
            }
        }
        
        // Recurse
        collect_run_descendant_statuses(doc, child, statuses, skip_txbx);
    }
}

/// Handle w:pPr elements - C# lines 2694-2732
fn handle_ppr_element(doc: &mut XmlDocument, ppr: NodeId, settings: &RevisionSettings) {
    // Get status attribute directly on pPr
    let status = doc.get(ppr)
        .and_then(|d| d.attributes())
        .and_then(|attrs| attrs.iter().find(|a| a.name == pt_status()))
        .map(|a| a.value.clone());
    
    let Some(status) = status else {
        // No status, just recurse (C# lines 2697-2700)
        let children: Vec<_> = doc.children(ppr).collect();
        for child in children {
            mark_content_transform(doc, child, settings);
        }
        return;
    };
    
    // Remove pt:Status from pPr
    doc.remove_attribute(ppr, &pt_status());
    
    // Find or create w:rPr child (C# lines 2704-2706, 2718-2720)
    let rpr = doc.children(ppr)
        .find(|&c| {
            doc.get(c)
                .and_then(|d| d.name())
                .map(|n| n.namespace.as_deref() == Some(W::NS) && n.local_name == "rPr")
                .unwrap_or(false)
        });
    
    let rpr = match rpr {
        Some(r) => r,
        None => {
            // Create new rPr and add as first child
            let new_rpr = doc.new_node(XmlNodeData::element(W::rPr()));
            // Insert at beginning of pPr children
            let first_child = doc.children(ppr).next();
            if let Some(first) = first_child {
                doc.insert_before(first, new_rpr);
            } else {
                doc.reparent(ppr, new_rpr);
            }
            new_rpr
        }
    };
    
    // Add w:del or w:ins inside rPr (C# lines 2707-2710, 2721-2724)
    if status == "Deleted" {
        let id_str = next_revision_id().to_string();
        let del_attrs = vec![
            XAttribute::new(W::id(), &id_str),  // w:id MUST come first per ECMA-376
            XAttribute::new(W::author(), &settings.author),
            XAttribute::new(W::date(), &settings.date_time),
            // Add w16du:dateUtc for modern Word timezone handling
            XAttribute::new(W16DU::dateUtc(), &settings.date_time),
        ];
        doc.add_child(rpr, XmlNodeData::element_with_attrs(W::del(), del_attrs));
    } else if status == "Inserted" {
        let id_str = next_revision_id().to_string();
        let ins_attrs = vec![
            XAttribute::new(W::id(), &id_str),  // w:id MUST come first per ECMA-376
            XAttribute::new(W::author(), &settings.author),
            XAttribute::new(W::date(), &settings.date_time),
            // Add w16du:dateUtc for modern Word timezone handling
            XAttribute::new(W16DU::dateUtc(), &settings.date_time),
        ];
        doc.add_child(rpr, XmlNodeData::element_with_attrs(W::ins(), ins_attrs));
    }
}

fn handle_omath_element(doc: &mut XmlDocument, node: NodeId, settings: &RevisionSettings) {
    // Get status attribute directly on oMath/oMathPara
    let status = doc.get(node)
        .and_then(|d| d.attributes())
        .and_then(|attrs| attrs.iter().find(|a| a.name == pt_status()))
        .map(|a| a.value.clone());
    
    let Some(status) = status else {
        // No status, just recurse
        let children: Vec<_> = doc.children(node).collect();
        for child in children {
            mark_content_transform(doc, child, settings);
        }
        return;
    };
    
    // Remove pt:Status
    doc.remove_attribute(node, &pt_status());
    
    // Wrap in w:del or w:ins
    if status == "Deleted" {
        let id_str = next_revision_id().to_string();
        let del_attrs = vec![
            XAttribute::new(W::id(), &id_str),
            XAttribute::new(W::author(), &settings.author),
            XAttribute::new(W::date(), &settings.date_time),
            XAttribute::new(W16DU::dateUtc(), &settings.date_time),
        ];
        let del_elem = doc.new_node(XmlNodeData::element_with_attrs(W::del(), del_attrs));
        
        doc.insert_before(node, del_elem);
        doc.detach(node);
        doc.reparent(del_elem, node);
    } else if status == "Inserted" {
        let id_str = next_revision_id().to_string();
        let ins_attrs = vec![
            XAttribute::new(W::id(), &id_str),
            XAttribute::new(W::author(), &settings.author),
            XAttribute::new(W::date(), &settings.date_time),
            XAttribute::new(W16DU::dateUtc(), &settings.date_time),
        ];
        let ins_elem = doc.new_node(XmlNodeData::element_with_attrs(W::ins(), ins_attrs));
        
        doc.insert_before(node, ins_elem);
        doc.detach(node);
        doc.reparent(ins_elem, node);
    }
}

fn is_props_element(doc: &XmlDocument, node: NodeId) -> bool {
    doc.get(node).and_then(|d| d.name()).map(|n| {
        n.namespace.as_deref() == Some(W::NS) && n.local_name.ends_with("Pr")
    }).unwrap_or(false)
}

/// Additional run container element names that should be processed recursively.
/// These are elements that can contain runs like paragraphs.
/// From C# AdditionalRunContainerNames set.
static ADDITIONAL_RUN_CONTAINER_NAMES: &[&str] = &[
    "bdo", "customXml", "dir", "fld", "hyperlink", "sdtContent", "smartTag",
];

/// Coalesce adjacent runs with identical formatting - faithful port of C# CoalesceAdjacentRunsWithIdenticalFormatting.
/// 
/// This function merges adjacent w:r, w:ins, and w:del elements that have identical formatting
/// into single elements. This is a post-processing step after document comparison to clean up
/// the output and reduce redundant markup.
/// 
/// # Algorithm (from C# PtOpenXmlUtil.cs lines 799-991)
/// 1. For each element in the run container:
///    - w:r with single w:t child: Key = "Wt" + rPr.ToString()
///    - w:r with single w:instrText child: Key = "WinstrText" + rPr.ToString()
///    - w:ins with valid structure: Key = "Wins2" + author + date + id + rPr strings
///    - w:del with valid structure: Key = "Wdel" + author + date + rPr strings
///    - Anything else: Key = "DontConsolidate"
/// 2. GroupAdjacent by key
/// 3. For each group with same key, merge text content
/// 4. Process w:txbxContent and additional run containers recursively
pub fn coalesce_adjacent_runs_with_identical_formatting(doc: &mut XmlDocument, run_container: NodeId) {
    // Step 1: Group adjacent elements by consolidation key
    let children: Vec<_> = doc.children(run_container).collect();
    if children.is_empty() {
        return;
    }

    // Compute keys for all children once (cache to avoid recomputation)
    let keys: Vec<String> = children.iter().map(|&c| get_consolidation_key(doc, c)).collect();

    // Step 2: Group adjacent elements with same key
    let groups = group_adjacent_by_key(&children, &keys);

    // Step 3: Process each group - merge if consolidatable, keep original otherwise
    let mut new_children: Vec<NodeId> = Vec::new();
    for (key, group) in groups {
        if key == DONT_CONSOLIDATE {
            // Keep original elements unchanged
            new_children.extend(group);
        } else if group.len() == 1 {
            // Single element, no merging needed
            new_children.push(group[0]);
        } else {
            // Merge group into single element
            let merged = merge_consolidated_group(doc, &key, &group);
            new_children.push(merged);
        }
    }

    // Step 4: Replace children with new consolidated children
    for &child in &children {
        doc.detach(child);
    }
    for &child in &new_children {
        doc.reparent(run_container, child);
    }

    // Step 5: Process w:txbxContent recursively (C# lines 971-977)
    process_textbox_content(doc, run_container);

    // Step 6: Process additional run containers recursively (C# lines 979-988)
    process_additional_run_containers(doc, run_container);
}

/// Wrapper that processes all paragraphs in a document tree.
/// This is the entry point for coalescing after document comparison.
pub fn coalesce_adjacent_runs(doc: &mut XmlDocument, root: NodeId, _settings: &WmlComparerSettings) {
    let mut paragraphs = Vec::new();
    collect_run_containers(doc, root, &mut paragraphs);
    for para in paragraphs {
        coalesce_adjacent_runs_with_identical_formatting(doc, para);
    }
}

/// Collect all paragraph elements (w:p) for processing
fn collect_run_containers(doc: &XmlDocument, node: NodeId, result: &mut Vec<NodeId>) {
    if let Some(data) = doc.get(node) {
        if let Some(name) = data.name() {
            if name.namespace.as_deref() == Some(W::NS) && name.local_name == "p" {
                result.push(node);
                // Don't recurse into paragraph children for this collection
                return;
            }
        }
    }
    let children: Vec<_> = doc.children(node).collect();
    for child in children {
        collect_run_containers(doc, child, result);
    }
}

/// Process w:txbxContent elements recursively (C# lines 971-977)
fn process_textbox_content(doc: &mut XmlDocument, node: NodeId) {
    let txbx_elements: Vec<_> = collect_descendants_by_name(doc, node, W::NS, "txbxContent");
    for txbx in txbx_elements {
        // Find all paragraphs within this textbox (trimmed - don't recurse into nested txbxContent)
        let paras = collect_descendants_trimmed(doc, txbx, W::NS, "p", W::NS, "txbxContent");
        for para in paras {
            coalesce_adjacent_runs_with_identical_formatting(doc, para);
        }
    }
}

/// Process additional run containers recursively (C# lines 979-988)
fn process_additional_run_containers(doc: &mut XmlDocument, node: NodeId) {
    let containers: Vec<_> = collect_additional_run_containers(doc, node);
    for container in containers {
        coalesce_adjacent_runs_with_identical_formatting(doc, container);
    }
}

/// Collect descendants that match a specific name
fn collect_descendants_by_name(doc: &XmlDocument, node: NodeId, ns: &str, local: &str) -> Vec<NodeId> {
    let mut result = Vec::new();
    for desc in doc.descendants(node) {
        if let Some(data) = doc.get(desc) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(ns) && name.local_name == local {
                    result.push(desc);
                }
            }
        }
    }
    result
}

/// Collect descendants that match target name, but stop at trim elements (like DescendantsTrimmed in C#)
fn collect_descendants_trimmed(
    doc: &XmlDocument,
    start: NodeId,
    target_ns: &str,
    target_local: &str,
    trim_ns: &str,
    trim_local: &str,
) -> Vec<NodeId> {
    let mut result = Vec::new();
    let mut stack = vec![start];
    
    while let Some(current) = stack.pop() {
        for child in doc.children(current) {
            if let Some(data) = doc.get(child) {
                if let Some(name) = data.name() {
                    // Skip if this is a trim element
                    if name.namespace.as_deref() == Some(trim_ns) && name.local_name == trim_local {
                        continue;
                    }
                    // Check if this matches our target
                    if name.namespace.as_deref() == Some(target_ns) && name.local_name == target_local {
                        result.push(child);
                    }
                }
            }
            // Continue searching children
            stack.push(child);
        }
    }
    result
}

/// Collect additional run container elements for recursive processing
fn collect_additional_run_containers(doc: &XmlDocument, node: NodeId) -> Vec<NodeId> {
    let mut result = Vec::new();
    for desc in doc.descendants(node) {
        if let Some(data) = doc.get(desc) {
            if let Some(name) = data.name() {
                if ADDITIONAL_RUN_CONTAINER_NAMES.contains(&name.local_name.as_str()) {
                    result.push(desc);
                }
            }
        }
    }
    result
}

/// Sentinel value for elements that should not be consolidated
const DONT_CONSOLIDATE: &str = "DontConsolidate";

/// Generate consolidation key for an element - faithful port of C# logic (lines 806-910)
/// 
/// Key generation rules:
/// - w:r with single w:t: "Wt" + rPr.ToString(SaveOptions.None)
/// - w:r with single w:instrText: "WinstrText" + rPr.ToString(SaveOptions.None)
/// - w:ins with valid structure: "Wins2" + author + date + id + rPr strings
/// - w:del with valid structure: "Wdel" + author + date + rPr strings
/// - Everything else: "DontConsolidate"
fn get_consolidation_key(doc: &XmlDocument, node: NodeId) -> String {
    let Some(data) = doc.get(node) else {
        return DONT_CONSOLIDATE.to_string();
    };
    let Some(name) = data.name() else {
        return DONT_CONSOLIDATE.to_string();
    };
    
    // Only process WordprocessingML elements
    if name.namespace.as_deref() != Some(W::NS) {
        return DONT_CONSOLIDATE.to_string();
    }

    match name.local_name.as_str() {
        "r" => get_run_consolidation_key(doc, node),
        "ins" => get_ins_consolidation_key(doc, node),
        "del" => get_del_consolidation_key(doc, node),
        _ => DONT_CONSOLIDATE.to_string(),
    }
}

/// Get consolidation key for w:r element (C# lines 808-826)
fn get_run_consolidation_key(doc: &XmlDocument, run: NodeId) -> String {
    // Count non-rPr children - must be exactly 1
    let non_rpr_children: Vec<_> = doc.children(run)
        .filter(|&c| !is_element_named(doc, c, W::NS, "rPr"))
        .collect();
    
    if non_rpr_children.len() != 1 {
        return DONT_CONSOLIDATE.to_string();
    }

    // Check for pt:AbstractNumId attribute (C# line 813-814)
    if has_attribute(doc, run, PT_STATUS_NS, "AbstractNumId") {
        return DONT_CONSOLIDATE.to_string();
    }

    // Get rPr serialization
    let rpr_string = get_rpr_string(doc, run);

    // Check what kind of content child we have
    let child = non_rpr_children[0];
    if is_element_named(doc, child, W::NS, "t") {
        format!("Wt{}", rpr_string)
    } else if is_element_named(doc, child, W::NS, "instrText") {
        format!("WinstrText{}", rpr_string)
    } else {
        DONT_CONSOLIDATE.to_string()
    }
}

/// Get consolidation key for w:ins element (C# lines 828-887)
fn get_ins_consolidation_key(doc: &XmlDocument, ins: NodeId) -> String {
    // If contains w:del, don't consolidate (C# line 830-832)
    if has_child_element(doc, ins, W::NS, "del") {
        return DONT_CONSOLIDATE.to_string();
    }

    // Check grandchildren: ce.Elements().Elements().Count(e => e.Name != W.rPr) != 1
    // And require at least one w:t grandchild
    let grandchildren: Vec<_> = doc.children(ins)
        .flat_map(|c| doc.children(c))
        .collect();
    
    let non_rpr_grandchildren_count = grandchildren.iter()
        .filter(|&&gc| !is_element_named(doc, gc, W::NS, "rPr"))
        .count();
    
    if non_rpr_grandchildren_count != 1 {
        return DONT_CONSOLIDATE.to_string();
    }

    let has_t_grandchild = grandchildren.iter()
        .any(|&gc| is_element_named(doc, gc, W::NS, "t"));
    
    if !has_t_grandchild {
        return DONT_CONSOLIDATE.to_string();
    }

    // Build key: "Wins2" + author + date + rPr strings
    // DEVIATION FROM C#: We intentionally EXCLUDE w:id from the key.
    // C# includes w:id (PtOpenXmlUtil.cs lines 877-886), but since we generate
    // unique IDs for each w:ins, including it prevents merging of adjacent
    // insertions with identical author/date/formatting. Excluding it produces
    // cleaner XML with fewer, larger revision blocks.
    // To restore C# parity, see commit 39a834a.
    let author = get_attr(doc, ins, W::NS, "author").unwrap_or_default();
    let date = format_date_for_key(&get_attr(doc, ins, W::NS, "date").unwrap_or_default());

    // Concatenate rPr strings from all child runs (C# lines 883-886)
    let rpr_strings: String = doc.children(ins)
        .filter_map(|c| {
            if is_element_named(doc, c, W::NS, "r") {
                Some(get_rpr_string(doc, c))
            } else {
                None
            }
        })
        .collect::<Vec<_>>()
        .join("");

    format!("Wins2{}{}{}", author, date, rpr_strings)
}

/// Get consolidation key for w:del element (C# lines 889-907)
fn get_del_consolidation_key(doc: &XmlDocument, del: NodeId) -> String {
    // Check ce.Elements(W.r).Elements().Count(e => e.Name != W.rPr) != 1
    let r_children: Vec<_> = doc.children(del)
        .filter(|&c| is_element_named(doc, c, W::NS, "r"))
        .collect();
    
    let grandchildren: Vec<_> = r_children.iter()
        .flat_map(|&r| doc.children(r))
        .collect();
    
    let non_rpr_grandchildren_count = grandchildren.iter()
        .filter(|&&gc| !is_element_named(doc, gc, W::NS, "rPr"))
        .count();
    
    if non_rpr_grandchildren_count != 1 {
        return DONT_CONSOLIDATE.to_string();
    }

    // Require at least one w:delText grandchild
    let has_del_text = grandchildren.iter()
        .any(|&gc| is_element_named(doc, gc, W::NS, "delText"));
    
    if !has_del_text {
        return DONT_CONSOLIDATE.to_string();
    }

    // Build key: "Wdel" + author + date + rPr strings (note: no id, unlike w:ins)
    let author = get_attr(doc, del, W::NS, "author").unwrap_or_default();
    let date = format_date_for_key(&get_attr(doc, del, W::NS, "date").unwrap_or_default());

    // Concatenate rPr strings from all w:r children (C# lines 903-906)
    let rpr_strings: String = r_children.iter()
        .map(|&r| get_rpr_string(doc, r))
        .collect::<Vec<_>>()
        .join("");

    format!("Wdel{}{}{}", author, date, rpr_strings)
}

/// Format date value for key generation.
/// C# uses ((DateTime)date).ToString("s") which produces "yyyy-MM-ddTHH:mm:ss"
fn format_date_for_key(date_str: &str) -> String {
    // If the date is already in a compatible format, try to parse and reformat
    // to match C#'s "s" (sortable) format: yyyy-MM-ddTHH:mm:ss
    if date_str.is_empty() {
        return String::new();
    }
    
    // Try to parse ISO 8601 format and reformat without timezone
    // Most Word dates are like "2023-01-15T10:30:00Z" or "2023-01-15T10:30:00+00:00"
    if date_str.len() >= 19 {
        // Take first 19 chars: "yyyy-MM-ddTHH:mm:ss"
        return date_str.chars().take(19).collect();
    }
    
    // Return as-is if not in expected format
    date_str.to_string()
}

/// Get serialized rPr string for key generation.
/// This should match C#'s rPr.ToString(SaveOptions.None) behavior.
fn get_rpr_string(doc: &XmlDocument, run: NodeId) -> String {
    let rpr = doc.children(run)
        .find(|&c| is_element_named(doc, c, W::NS, "rPr"));
    
    match rpr {
        Some(rpr_node) => serialize_element_for_key(doc, rpr_node),
        None => String::new(),
    }
}

/// Serialize an element to string for use in consolidation key.
/// This aims to match C#'s XElement.ToString(SaveOptions.None) behavior:
/// - No indentation or newlines
/// - Attributes in document order
/// - Namespace prefixes as stored
fn serialize_element_for_key(doc: &XmlDocument, node: NodeId) -> String {
    let mut result = String::new();
    serialize_element_recursive(doc, node, &mut result);
    result
}

/// Recursive helper for element serialization
fn serialize_element_recursive(doc: &XmlDocument, node: NodeId, result: &mut String) {
    let Some(data) = doc.get(node) else { return };
    
    match data {
        XmlNodeData::Element { name, attributes } => {
            result.push('<');
            
            // Add prefix if namespace is present
            if let Some(ns) = &name.namespace {
                if let Some(prefix) = get_prefix_for_ns(ns) {
                    result.push_str(prefix);
                    result.push(':');
                }
            }
            result.push_str(&name.local_name);
            
            // Add attributes in document order
            for attr in attributes {
                result.push(' ');
                if let Some(ns) = &attr.name.namespace {
                    if let Some(prefix) = get_prefix_for_ns(ns) {
                        result.push_str(prefix);
                        result.push(':');
                    }
                }
                result.push_str(&attr.name.local_name);
                result.push_str("=\"");
                result.push_str(&escape_xml_attr(&attr.value));
                result.push('"');
            }
            
            let children: Vec<_> = doc.children(node).collect();
            if children.is_empty() {
                result.push_str(" />");
            } else {
                result.push('>');
                for child in children {
                    serialize_element_recursive(doc, child, result);
                }
                result.push_str("</");
                if let Some(ns) = &name.namespace {
                    if let Some(prefix) = get_prefix_for_ns(ns) {
                        result.push_str(prefix);
                        result.push(':');
                    }
                }
                result.push_str(&name.local_name);
                result.push('>');
            }
        }
        XmlNodeData::Text(text) => {
            result.push_str(&escape_xml_text(text));
        }
        _ => {}
    }
}

/// Get namespace prefix for common namespaces
fn get_prefix_for_ns(ns: &str) -> Option<&'static str> {
    match ns {
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main" => Some("w"),
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships" => Some("r"),
        "http://schemas.openxmlformats.org/markup-compatibility/2006" => Some("mc"),
        "http://powertools.codeplex.com/2011" => Some("pt14"),
        "http://www.w3.org/XML/1998/namespace" => Some("xml"),
        _ => None,
    }
}

/// Escape XML attribute value
fn escape_xml_attr(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// Escape XML text content
fn escape_xml_text(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

/// Group adjacent elements by key
fn group_adjacent_by_key(children: &[NodeId], keys: &[String]) -> Vec<(String, Vec<NodeId>)> {
    if children.is_empty() {
        return Vec::new();
    }

    let mut groups: Vec<(String, Vec<NodeId>)> = Vec::new();
    let mut current_key = keys[0].clone();
    let mut current_group = vec![children[0]];

    for i in 1..children.len() {
        if keys[i] == current_key {
            current_group.push(children[i]);
        } else {
            groups.push((current_key, current_group));
            current_key = keys[i].clone();
            current_group = vec![children[i]];
        }
    }
    groups.push((current_key, current_group));

    groups
}

/// Merge a group of elements with identical formatting into a single element.
/// This creates a new element with merged text content.
fn merge_consolidated_group(doc: &mut XmlDocument, key: &str, group: &[NodeId]) -> NodeId {
    // Determine element type from key prefix
    let first = group[0];
    
    // Collect all text content from the group
    let merged_text = collect_text_from_group(doc, group, key);
    
    // Get xml:space attribute based on merged text
    let needs_preserve = needs_xml_space_preserve(&merged_text);
    
    if key.starts_with("Wt") || key.starts_with("WinstrText") {
        // w:r merge
        merge_run_group(doc, group, &merged_text, needs_preserve, key.starts_with("WinstrText"))
    } else if key.starts_with("Wins2") {
        // w:ins merge
        merge_ins_group(doc, group, &merged_text, needs_preserve)
    } else if key.starts_with("Wdel") {
        // w:del merge
        merge_del_group(doc, group, &merged_text, needs_preserve)
    } else {
        // Should not reach here, but return first element as fallback
        first
    }
}

/// Collect text content from a group of elements
fn collect_text_from_group(doc: &XmlDocument, group: &[NodeId], _key: &str) -> String {
    let mut text = String::new();
    
    for &node in group {
        // Collect text from descendants (w:t, w:delText, or w:instrText)
        for desc in doc.descendants(node) {
            if let Some(data) = doc.get(desc) {
                if let Some(name) = data.name() {
                    let is_text_element = 
                        (name.namespace.as_deref() == Some(W::NS)) &&
                        (name.local_name == "t" || name.local_name == "delText" || name.local_name == "instrText");
                    
                    if is_text_element {
                        // Get text content
                        for child in doc.children(desc) {
                            if let Some(XmlNodeData::Text(t)) = doc.get(child) {
                                text.push_str(t);
                            }
                        }
                    }
                }
            }
        }
    }
    
    text
}

/// Check if xml:space="preserve" is needed
fn needs_xml_space_preserve(text: &str) -> bool {
    if text.is_empty() {
        return false;
    }
    let chars: Vec<char> = text.chars().collect();
    chars[0].is_whitespace() || chars[chars.len() - 1].is_whitespace()
}

/// Merge w:r elements (C# lines 928-944)
fn merge_run_group(doc: &mut XmlDocument, group: &[NodeId], text: &str, preserve_space: bool, is_instr_text: bool) -> NodeId {
    let first = group[0];
    
    // Get attributes and rPr from first element
    let first_attrs = get_element_attributes(doc, first);
    let first_rpr = find_child_by_name(doc, first, W::NS, "rPr");
    
    // Collect pt:Status attributes from first w:t of each run (C# lines 932-933)
    let status_attrs: Vec<XAttribute> = if !is_instr_text {
        group.iter()
            .filter_map(|&r| {
                let t = find_descendant_by_name(doc, r, W::NS, "t")?;
                get_attribute(doc, t, PT_STATUS_NS, "Status")
                    .map(|val| XAttribute::new(pt_status(), &val))
            })
            .collect()
    } else {
        Vec::new()
    };

    // Create new run element
    let new_run = doc.new_node(XmlNodeData::element_with_attrs(W::r(), first_attrs));
    
    // Add rPr if present and non-empty
    if let Some(rpr) = first_rpr {
        if doc.children(rpr).next().is_some() {
            let cloned_rpr = clone_element_deep(doc, rpr);
            doc.reparent(new_run, cloned_rpr);
        }
    }
    
    // Create text element (w:t or w:instrText)
    let text_elem_name = if is_instr_text { W::instrText() } else { W::t() };
    let mut text_attrs = status_attrs;
    if preserve_space {
        text_attrs.push(XAttribute::new(
            XName::new("http://www.w3.org/XML/1998/namespace", "space"),
            "preserve",
        ));
    }
    
    let text_elem = doc.add_child(new_run, XmlNodeData::element_with_attrs(text_elem_name, text_attrs));
    doc.add_child(text_elem, XmlNodeData::Text(text.to_string()));
    
    new_run
}

/// Merge w:ins elements (C# lines 947-956)
fn merge_ins_group(doc: &mut XmlDocument, group: &[NodeId], text: &str, preserve_space: bool) -> NodeId {
    let first = group[0];
    
    // Get attributes from w:ins
    let ins_attrs = get_element_attributes(doc, first);
    
    // Get first w:r and its attributes/rPr
    let first_r = find_child_by_name(doc, first, W::NS, "r");
    let r_attrs = first_r.map(|r| get_element_attributes(doc, r)).unwrap_or_default();
    let r_rpr = first_r.and_then(|r| find_child_by_name(doc, r, W::NS, "rPr"));
    
    // Create new w:ins
    let new_ins = doc.new_node(XmlNodeData::element_with_attrs(W::ins(), ins_attrs));
    
    // Create inner w:r
    let new_r = doc.add_child(new_ins, XmlNodeData::element_with_attrs(W::r(), r_attrs));
    
    // Add rPr if present and non-empty
    if let Some(rpr) = r_rpr {
        if doc.children(rpr).next().is_some() {
            let cloned_rpr = clone_element_deep(doc, rpr);
            doc.reparent(new_r, cloned_rpr);
        }
    }
    
    // Create w:t with text
    let mut t_attrs = Vec::new();
    if preserve_space {
        t_attrs.push(XAttribute::new(
            XName::new("http://www.w3.org/XML/1998/namespace", "space"),
            "preserve",
        ));
    }
    let t_elem = doc.add_child(new_r, XmlNodeData::element_with_attrs(W::t(), t_attrs));
    doc.add_child(t_elem, XmlNodeData::Text(text.to_string()));
    
    new_ins
}

/// Merge w:del elements (C# lines 958-967)
fn merge_del_group(doc: &mut XmlDocument, group: &[NodeId], text: &str, preserve_space: bool) -> NodeId {
    let first = group[0];
    
    // Get attributes from w:del
    let del_attrs = get_element_attributes(doc, first);
    
    // Get first w:r and its attributes/rPr
    let first_r = find_child_by_name(doc, first, W::NS, "r");
    let r_attrs = first_r.map(|r| get_element_attributes(doc, r)).unwrap_or_default();
    let r_rpr = first_r.and_then(|r| find_child_by_name(doc, r, W::NS, "rPr"));
    
    // Create new w:del
    let new_del = doc.new_node(XmlNodeData::element_with_attrs(W::del(), del_attrs));
    
    // Create inner w:r
    let new_r = doc.add_child(new_del, XmlNodeData::element_with_attrs(W::r(), r_attrs));
    
    // Add rPr if present and non-empty
    if let Some(rpr) = r_rpr {
        if doc.children(rpr).next().is_some() {
            let cloned_rpr = clone_element_deep(doc, rpr);
            doc.reparent(new_r, cloned_rpr);
        }
    }
    
    // Create w:delText with text
    let mut dt_attrs = Vec::new();
    if preserve_space {
        dt_attrs.push(XAttribute::new(
            XName::new("http://www.w3.org/XML/1998/namespace", "space"),
            "preserve",
        ));
    }
    let dt_elem = doc.add_child(new_r, XmlNodeData::element_with_attrs(W::delText(), dt_attrs));
    doc.add_child(dt_elem, XmlNodeData::Text(text.to_string()));
    
    new_del
}

/// Get attributes from an element
fn get_element_attributes(doc: &XmlDocument, node: NodeId) -> Vec<XAttribute> {
    doc.get(node)
        .and_then(|data| data.attributes())
        .map(|attrs| attrs.to_vec())
        .unwrap_or_default()
}

/// Get a specific attribute value
fn get_attribute(doc: &XmlDocument, node: NodeId, ns: &str, local: &str) -> Option<String> {
    doc.get(node)?
        .attributes()?
        .iter()
        .find(|a| a.name.namespace.as_deref() == Some(ns) && a.name.local_name == local)
        .map(|a| a.value.clone())
}

/// Find first descendant with given name
fn find_descendant_by_name(doc: &XmlDocument, node: NodeId, ns: &str, local: &str) -> Option<NodeId> {
    for desc in doc.descendants(node) {
        if is_element_named(doc, desc, ns, local) {
            return Some(desc);
        }
    }
    None
}

/// Clone an element and all its children
fn clone_element_deep(doc: &mut XmlDocument, source: NodeId) -> NodeId {
    let source_data = doc.get(source).expect("Source node must exist").clone();
    let cloned = doc.new_node(source_data);
    
    let children: Vec<_> = doc.children(source).collect();
    for child in children {
        let cloned_child = clone_element_deep(doc, child);
        doc.reparent(cloned, cloned_child);
    }
    
    cloned
}

/// Check if element has a specific name
fn is_element_named(doc: &XmlDocument, node: NodeId, ns: &str, local: &str) -> bool {
    doc.get(node)
        .and_then(|data| data.name())
        .map(|name| name.namespace.as_deref() == Some(ns) && name.local_name == local)
        .unwrap_or(false)
}

/// Check if element has a child with specific name
fn has_child_element(doc: &XmlDocument, parent: NodeId, ns: &str, local: &str) -> bool {
    doc.children(parent).any(|c| is_element_named(doc, c, ns, local))
}

/// Check if element has a specific attribute
fn has_attribute(doc: &XmlDocument, node: NodeId, ns: &str, local: &str) -> bool {
    doc.get(node)
        .and_then(|data| data.attributes())
        .map(|attrs| attrs.iter().any(|a| 
            a.name.namespace.as_deref() == Some(ns) && a.name.local_name == local
        ))
        .unwrap_or(false)
}

fn is_r_pr(doc: &XmlDocument, node: NodeId) -> bool {
    doc.get(node).and_then(|d| d.name()).map(|n| n == &W::rPr()).unwrap_or(false)
}

fn is_t(doc: &XmlDocument, node: NodeId) -> bool {
    doc.get(node).and_then(|d| d.name()).map(|n| n == &W::t()).unwrap_or(false)
}

fn is_del_text(doc: &XmlDocument, node: NodeId) -> bool {
    doc.get(node).and_then(|d| d.name()).map(|n| n == &W::delText()).unwrap_or(false)
}

fn get_r_pr_signature(doc: &mut XmlDocument, run_node: NodeId, settings: &WmlComparerSettings) -> String {
    let normalized = crate::wml::formatting::compute_normalized_rpr(doc, run_node, settings);
    crate::wml::formatting::compute_formatting_signature(doc, normalized).unwrap_or_default()
}

fn get_attr(doc: &XmlDocument, node: NodeId, ns: &str, local: &str) -> Option<String> {
    doc.get(node)?.attributes()?.iter().find(|a| a.name.local_name == local && a.name.namespace.as_deref() == Some(ns)).map(|a| a.value.clone())
}

fn merge_nodes(doc: &mut XmlDocument, nodes: &[NodeId]) -> NodeId {
    let first = nodes[0];
    let mut combined_text = String::new();
    for &node in nodes { combined_text.push_str(&get_text_from_run_container(doc, node)); }
    set_text_in_run_container(doc, first, &combined_text);
    first
}

fn get_text_from_run_container(doc: &XmlDocument, node: NodeId) -> String {
    let run = if doc.get(node).and_then(|d| d.name()).map(|n| n.local_name == "r").unwrap_or(false) { node } else { find_child_by_name(doc, node, W::NS, "r").unwrap_or(node) };
    let t_node = doc.children(run).find(|&c| is_t(doc, c) || is_del_text(doc, c));
    if let Some(t) = t_node {
        if let Some(XmlNodeData::Text(txt)) = doc.children(t).next().and_then(|c| doc.get(c)) { return txt.clone(); }
    }
    String::new()
}

fn set_text_in_run_container(doc: &mut XmlDocument, node: NodeId, text: &str) {
    let run = if doc.get(node).and_then(|d| d.name()).map(|n| n.local_name == "r").unwrap_or(false) { node } else { find_child_by_name(doc, node, W::NS, "r").unwrap_or(node) };
    let t_node = doc.children(run).find(|&c| is_t(doc, c) || is_del_text(doc, c));
    if let Some(t) = t_node {
        let text_node = doc.children(t).next().unwrap();
        if let Some(XmlNodeData::Text(ref mut txt)) = doc.get_mut(text_node) { *txt = text.to_string(); }
    }
}

fn find_child_by_name(doc: &XmlDocument, parent: NodeId, ns: &str, local: &str) -> Option<NodeId> {
    for child in doc.children(parent) {
        if let Some(data) = doc.get(child) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(ns) && name.local_name == local { return Some(child); }
            }
        }
    }
    None
}

fn coalesce_recurse(
    doc: &mut XmlDocument,
    parent: NodeId,
    list: &[ComparisonUnitAtom],
    level: usize,
    _part: Option<()>,
    settings: &WmlComparerSettings,
) {
    // Use range-based grouping to avoid cloning
    let grouped = group_by_key_ranges(list, |ca| {
        if level >= ca.ancestor_unids.len() { String::new() } else { ca.ancestor_unids[level].clone() }
    });
    
    // DEBUG: Track empty group_key atoms
    // static DEBUG_EMPTY_GROUPS: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(false);
    
    for (group_key, start, end) in grouped {
        if group_key.is_empty() {
            // Handle direct children (atoms that don't have further ancestors at this level)
            // e.g. w:ptab inside w:p, or w:br inside w:p
            let direct_children = &list[start..end];
            
            /*
            // DEBUG: Log first occurrence of empty group key
            if !DEBUG_EMPTY_GROUPS.swap(true, std::sync::atomic::Ordering::SeqCst) {
                let del_count = direct_children.iter().filter(|a| a.correlation_status == ComparisonCorrelationStatus::Deleted).count();
                let ins_count = direct_children.iter().filter(|a| a.correlation_status == ComparisonCorrelationStatus::Inserted).count();
                eprintln!("DEBUG empty group_key (first): level={}, direct_children={}, del={}, ins={}", level, direct_children.len(), del_count, ins_count);
                if let Some(first) = direct_children.first() {
                    eprintln!("  First atom: {:?}, ancestor_unids.len()={}", first.content_element.local_name(), first.ancestor_unids.len());
                }
            }
            */
            
            let is_inside_vml = direct_children.first().map(|a| is_inside_vml_content(a, level)).unwrap_or(false);
            let grouped_direct = group_adjacent_by_correlation_ranges(direct_children, level, is_inside_vml, settings);
            
            for (key, s, e) in grouped_direct {
                let atoms = &direct_children[s..e];
                let first = &atoms[0];
                let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
                let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
                
                let spl: Vec<&str> = key.split('|').collect();
                // If the key indicates a specific status override from grouping (e.g. from formatting change), use it
                // Otherwise default to the atom's status
                let status = if spl.len() > 1 && (spl[1] == "Deleted" || spl[1] == "Inserted") {
                    spl[1]
                } else if del { 
                    "Deleted" 
                } else if ins { 
                    "Inserted" 
                } else { 
                    "" 
                };
                
                for atom in atoms {
                    let content_node = create_content_element(doc, atom, status);
                    if let Some(node) = content_node {
                        doc.reparent(parent, node);
                    }
                }
            }
            continue;
        }
        
        let group_atoms = &list[start..end];
        let first_atom = &group_atoms[0];
        let ancestor_being_constructed = get_ancestor_element_for_level(first_atom, level, group_atoms);
        let ancestor_name = &ancestor_being_constructed.local_name;
        let is_inside_vml = is_inside_vml_content(first_atom, level);
        
        // Use range-based correlation grouping
        let grouped_children = group_adjacent_by_correlation_ranges(group_atoms, level, is_inside_vml, settings);
        
        match ancestor_name.as_str() {
            "p" => reconstruct_paragraph(doc, parent, &group_key, &ancestor_being_constructed, group_atoms, &grouped_children, level, is_inside_vml, _part, settings),
            "r" => reconstruct_run(doc, parent, &ancestor_being_constructed, group_atoms, &grouped_children, level, is_inside_vml, _part, settings),
            "t" => reconstruct_text_elements(doc, parent, group_atoms, &grouped_children),
            "drawing" => reconstruct_drawing_elements(doc, parent, group_atoms, &grouped_children, _part, settings),
            "oMath" | "oMathPara" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &[], group_atoms, level, _part, settings),
            elem if ALLOWABLE_RUN_CHILDREN.contains(&elem) => reconstruct_allowable_run_children(doc, parent, &ancestor_being_constructed, group_atoms, &grouped_children),
            "tbl" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["tblPr", "tblGrid"], group_atoms, level, _part, settings),
            "tr" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["trPr"], group_atoms, level, _part, settings),
            "tc" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["tcPr"], group_atoms, level, _part, settings),
            "sdt" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["sdtPr", "sdtEndPr"], group_atoms, level, _part, settings),
            "pict" | "shape" | "rect" | "group" | "shapetype" | "oval" | "line" | "arc" | "curve" | "polyline" | "roundrect" => reconstruct_vml_element(doc, parent, &group_key, &ancestor_being_constructed, group_atoms, level, _part, settings),
            "object" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["shapetype", "shape", "OLEObject"], group_atoms, level, _part, settings),
            "ruby" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["rubyPr"], group_atoms, level, _part, settings),
            _ => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &[], group_atoms, level, _part, settings),
        }
    }
}

// Helper to check if all atoms have uniform status (Inserted or Deleted)
fn get_uniform_status(atoms: &[ComparisonUnitAtom]) -> Option<ComparisonCorrelationStatus> {
    if atoms.is_empty() { return None; }
    let status = atoms[0].correlation_status;
    if matches!(status, ComparisonCorrelationStatus::Inserted | ComparisonCorrelationStatus::Deleted) {
        if atoms.iter().all(|a| a.correlation_status == status) {
            return Some(status);
        }
    }
    None
}

// Helper to create w:ins or w:del wrapper
fn create_revision_wrapper(
    doc: &mut XmlDocument, 
    parent: NodeId, 
    status: ComparisonCorrelationStatus, 
    settings: &WmlComparerSettings
) -> Option<NodeId> {
    let name = match status {
        ComparisonCorrelationStatus::Inserted => W::ins(),
        ComparisonCorrelationStatus::Deleted => W::del(),
        _ => return None,
    };
    
    let id_str = next_revision_id().to_string();
    let author = settings.author_for_revisions.as_deref().unwrap_or("Unknown");
    let date = settings.date_time_for_revisions.as_deref().unwrap_or("1970-01-01T00:00:00Z");
    
    let attrs = vec![
        XAttribute::new(W::id(), &id_str),
        XAttribute::new(W::author(), author),
        XAttribute::new(W::date(), date),
        XAttribute::new(W16DU::dateUtc(), date),
    ];
    
    Some(doc.add_child(parent, XmlNodeData::element_with_attrs(name, attrs)))
}

fn reconstruct_paragraph(doc: &mut XmlDocument, parent: NodeId, _group_key: &str, ancestor: &AncestorElementInfo, atoms: &[ComparisonUnitAtom], grouped_children: &[(String, usize, usize)], level: usize, is_inside_vml: bool, part: Option<()>, settings: &WmlComparerSettings) {
    let mut para_attrs = ancestor.attributes.clone();
    para_attrs.retain(|a| a.name.namespace.as_deref() != Some(PT_STATUS_NS));
    
    // CRITICAL FIX: Filter out empty paragraphs with w:rsidDel attribute
    // MS Word doesn't show deleted empty paragraphs in comparison output
    // MUST be checked BEFORE creating revision wrapper to avoid empty wrappers
    let has_rsid_del = ancestor.attributes.iter().any(|a| 
        a.name.namespace.as_deref() == Some(W::NS) && a.name.local_name == "rsidDel"
    );
    
    if has_rsid_del {
        // Check if paragraph is empty (no text content)
        // Use original atoms for check
        let has_content = grouped_children.iter().any(|(_, start, end)| {
            atoms[*start..*end].iter().any(|atom| {
                matches!(atom.content_element, ContentElement::Text(_))
            })
        });
        
        if !has_content {
            // Skip this empty paragraph with rsidDel - MS Word filters these out
            return;
        }
    }

    // Check for uniform status to wrap entire paragraph in w:ins or w:del
    let uniform_status = get_uniform_status(atoms);
    
    let (container, atoms_vec) = if let Some(status) = uniform_status {
        let wrapper = create_revision_wrapper(doc, parent, status, settings).unwrap();
        // Create modified atoms with Equal status to suppress inner revisions
        let mut new_atoms = atoms.to_vec();
        for atom in &mut new_atoms {
            atom.correlation_status = ComparisonCorrelationStatus::Equal;
        }
        (wrapper, Some(new_atoms))
    } else {
        (parent, None)
    };
    
    // Use either the modified atoms or the original slice
    let atoms_ref = atoms_vec.as_deref().unwrap_or(atoms);
    
    // Note: Do NOT add pt_unid to output - it's an internal tracking attribute
    let para = doc.add_child(container, XmlNodeData::element_with_attrs(W::p(), para_attrs));
    
    // OOXML requires w:pPr to be the FIRST child of w:p
    // First pass: add pPr elements
    for (key, start, end) in grouped_children {
        let group_atoms = &atoms_ref[*start..*end];
        let spl: Vec<&str> = key.split('|').collect();
        if spl.get(0) == Some(&"") {
            for gcc in group_atoms {
                if !matches!(&gcc.content_element, ContentElement::ParagraphProperties { .. }) { continue; }
                if is_inside_vml && spl.get(1) == Some(&"Inserted") { continue; }
                // Suppress status if uniform wrapper exists
                let status_arg = if uniform_status.is_some() { "" } else { spl.get(1).unwrap_or(&"") };
                let content_elem_node = create_content_element(doc, gcc, status_arg);
                if let Some(node) = content_elem_node { doc.reparent(para, node); }
            }
        }
    }
    
    // Second pass: add all other content (runs, etc.)
    for (key, start, end) in grouped_children {
        let group_atoms = &atoms_ref[*start..*end];
        let spl: Vec<&str> = key.split('|').collect();
        if spl.get(0) == Some(&"") {
            for gcc in group_atoms {
                // Skip pPr - already added in first pass
                if matches!(&gcc.content_element, ContentElement::ParagraphProperties { .. }) { continue; }
                // Suppress status if uniform wrapper exists
                let status_arg = if uniform_status.is_some() { "" } else { spl.get(1).unwrap_or(&"") };
                let content_elem_node = create_content_element(doc, gcc, status_arg);
                if let Some(node) = content_elem_node { doc.reparent(para, node); }
            }
        } else { coalesce_recurse(doc, para, group_atoms, level + 1, part, settings); }
    }
}


fn reconstruct_run(doc: &mut XmlDocument, parent: NodeId, ancestor: &AncestorElementInfo, atoms: &[ComparisonUnitAtom], grouped_children: &[(String, usize, usize)], level: usize, _is_inside_vml: bool, part: Option<()>, settings: &WmlComparerSettings) {
    let mut run_attrs = ancestor.attributes.clone();
    run_attrs.retain(|a| a.name.namespace.as_deref() != Some(PT_STATUS_NS));
    let run = doc.add_child(parent, XmlNodeData::element_with_attrs(W::r(), run_attrs));
    
    // CRITICAL FIX: Restore the w:rPr (run properties) from the ancestor
    // This matches C# WmlComparer.cs line 5362-5368:
    //   XElement rPr = ancestorBeingConstructed.Element(W.rPr);
    //   if (rPr != null) rPr = new XElement(rPr);
    //   var newRun = new XElement(W.r, ..., rPr, ...);
    if let Some(ref rpr_xml) = ancestor.rpr_xml {
        if let Some(rpr_node) = parse_rpr_xml(doc, rpr_xml) {
            doc.reparent(run, rpr_node);
        }
    }
    
    let format_changed = grouped_children.iter().any(|(_, start, end)| {
        atoms[*start..*end].iter().any(|atom| atom.correlation_status == ComparisonCorrelationStatus::FormatChanged)
    });
    for (key, start, end) in grouped_children {
        let group_atoms = &atoms[*start..*end];
        let spl: Vec<&str> = key.split('|').collect();
        if spl.get(0) == Some(&"") {
            for gcc in group_atoms {
                let content_elem_node = create_content_element(doc, gcc, spl.get(1).unwrap_or(&""));
                if let Some(node) = content_elem_node { doc.reparent(run, node); }
            }
        } else { coalesce_recurse(doc, run, group_atoms, level + 1, part, settings); }
    }
    if settings.track_formatting_changes && format_changed {
        let existing_rpr = doc.children(run).find(|&c| is_r_pr(doc, c));
        let rpr = match existing_rpr { Some(node) => node, None => doc.add_child(run, XmlNodeData::element(W::rPr())) };
        let revision_settings = RevisionSettings { author: settings.author_for_revisions.clone().unwrap_or_else(|| "Unknown".to_string()), date_time: settings.date_time_for_revisions.clone().unwrap_or_else(|| chrono::Utc::now().format("%Y-%m-%dT%H:%M:%SZ").to_string()) };
        let _rpr_change = create_run_property_change(doc, rpr, &revision_settings);
    }
}

/// Parse serialized rPr XML string and create nodes in the document.
/// Returns the root node of the parsed rPr element, or None if empty.
fn parse_rpr_xml(doc: &mut XmlDocument, rpr_xml: &str) -> Option<NodeId> {
    // Use the namespace wrapper parser to handle prefix-only serialized XML
    let node = parse_xml_fragment(doc, rpr_xml)?;
    
    // Don't return empty rPr elements - they're non-standard OOXML
    if doc.children(node).next().is_none() {
        return None;
    }
    
    Some(node)
}

/// Wrapper XML with all necessary namespace declarations for parsing XML fragments
/// that use prefixes like w:, a:, r:, m:, etc.
const NS_WRAPPER_PREFIX: &str = r#"<ns_wrapper xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:pt="http://powertools.codeplex.com/2011">"#;
const NS_WRAPPER_SUFFIX: &str = "</ns_wrapper>";

/// Parse an XML fragment that uses prefixes without namespace declarations.
/// Wraps the fragment in a temporary element with all necessary xmlns declarations,
/// parses it, then returns a deep copy of the first child (the actual element we want).
fn parse_xml_fragment(target_doc: &mut XmlDocument, fragment_xml: &str) -> Option<NodeId> {
    if fragment_xml.is_empty() {
        return None;
    }
    
    // Wrap the fragment with namespace declarations
    let wrapped_xml = format!("{}{}{}", NS_WRAPPER_PREFIX, fragment_xml, NS_WRAPPER_SUFFIX);
    
    // Parse the wrapped XML
    let source_doc = parse(&wrapped_xml).ok()?;
    let wrapper_root = source_doc.root()?;
    
    // Get the first child (our actual element) and deep copy it
    let first_child = source_doc.children(wrapper_root).next()?;
    
    // Create a temporary parent for detaching
    let temp_parent = target_doc.new_node(XmlNodeData::element(XName::local("temp")));
    let new_node = deep_copy_node_standalone(target_doc, temp_parent, &source_doc, first_child)?;
    target_doc.detach(new_node);
    
    Some(new_node)
}

/// Deep-copy a node from source_doc to target_doc as a child of parent (for standalone nodes)
fn deep_copy_node_standalone(target_doc: &mut XmlDocument, parent: NodeId, source_doc: &XmlDocument, source_node: NodeId) -> Option<NodeId> {
    let source_data = source_doc.get(source_node)?;

    // Clone the node data
    let new_data = source_data.clone();
    let new_node = target_doc.add_child(parent, new_data);

    // Recursively copy children
    for child in source_doc.children(source_node) {
        deep_copy_node_standalone(target_doc, new_node, source_doc, child);
    }

    Some(new_node)
}

/// Clone an entire subtree from one document into another.
fn clone_tree_into_doc(target_doc: &mut XmlDocument, source_doc: &XmlDocument, source_node: NodeId) -> Option<NodeId> {
    let source_data = source_doc.get(source_node)?;
    let cloned_data = source_data.clone();
    let new_node = target_doc.new_node(cloned_data);
    
    // Recursively clone children
    for child in source_doc.children(source_node) {
        if let Some(cloned_child) = clone_tree_into_doc(target_doc, source_doc, child) {
            target_doc.reparent(new_node, cloned_child);
        }
    }
    
    Some(new_node)
}

fn reconstruct_text_elements(doc: &mut XmlDocument, parent: NodeId, atoms: &[ComparisonUnitAtom], grouped_children: &[(String, usize, usize)]) {
    /*
    // DEBUG: Count incoming text atoms by status
    static DEBUG_CALLED: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(false);
    if !DEBUG_CALLED.swap(true, std::sync::atomic::Ordering::SeqCst) {
        let total = atoms.len();
        let text_atoms = atoms.iter().filter(|a| matches!(a.content_element, ContentElement::Text(_))).count();
        let del_text = atoms.iter().filter(|a| matches!(a.content_element, ContentElement::Text(_)) && a.correlation_status == ComparisonCorrelationStatus::Deleted).count();
        let ins_text = atoms.iter().filter(|a| matches!(a.content_element, ContentElement::Text(_)) && a.correlation_status == ComparisonCorrelationStatus::Inserted).count();
        eprintln!("DEBUG reconstruct_text_elements (first call): total atoms={}, text_atoms={}, del_text={}, ins_text={}", total, text_atoms, del_text, ins_text);
    }
    */
    
    for (_key, start, end) in grouped_children {
        let group_atoms = &atoms[*start..*end];
        let text_of_text_element: String = group_atoms.iter().filter_map(|gce| if let ContentElement::Text(ch) = gce.content_element { Some(ch) } else { None }).collect();
        if text_of_text_element.is_empty() { continue; }
        let first = &group_atoms[0];
        let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
        let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
        let elem_name = if del { W::delText() } else { W::t() };
        let mut attrs = Vec::new();
        if needs_xml_space(&text_of_text_element) { attrs.push(XAttribute::new(XName::new("http://www.w3.org/XML/1998/namespace", "space"), "preserve")); }
        if del { attrs.push(XAttribute::new(pt_status(), "Deleted")); } else if ins { attrs.push(XAttribute::new(pt_status(), "Inserted")); }
        let text_elem = doc.add_child(parent, XmlNodeData::element_with_attrs(elem_name, attrs));
        doc.add_child(text_elem, XmlNodeData::Text(text_of_text_element));
    }
}

fn reconstruct_drawing_elements(doc: &mut XmlDocument, parent: NodeId, atoms: &[ComparisonUnitAtom], grouped_children: &[(String, usize, usize)], _part: Option<()>, _settings: &WmlComparerSettings) {
    for (_key, start, end) in grouped_children {
        let group_atoms = &atoms[*start..*end];
        let first = &group_atoms[0];
        let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
        let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
        if del || ins {
            for gcc in group_atoms {
                if let ContentElement::Drawing { element_xml, .. } = &gcc.content_element {
                    // DEBUG: Check element_xml content
                    if element_xml.is_empty() || !element_xml.contains("inline") {
                        eprintln!("BUG in reconstruct_drawing_elements: element_xml empty or missing inline!");
                        eprintln!("  Length: {}", element_xml.len());
                        eprintln!("  First 200 chars: {}", &element_xml.chars().take(200).collect::<String>());
                    }
                    // Parse the stored element XML and deep-copy it
                    if let Some(drawing_node) = copy_element_from_xml(doc, parent, element_xml) {
                        let status = if del { "Deleted" } else { "Inserted" };
                        doc.set_attribute(drawing_node, &pt_status(), status);
                    }
                }
            }
        } else {
            for gcc in group_atoms {
                if let ContentElement::Drawing { element_xml, .. } = &gcc.content_element {
                    // DEBUG: Check element_xml content
                    if element_xml.is_empty() || !element_xml.contains("inline") {
                        eprintln!("BUG in reconstruct_drawing_elements (non-revision): element_xml empty or missing inline!");
                        eprintln!("  Length: {}", element_xml.len());
                        eprintln!("  First 200 chars: {}", &element_xml.chars().take(200).collect::<String>());
                    }
                    // Parse the stored element XML and deep-copy it
                    copy_element_from_xml(doc, parent, element_xml);
                }
            }
        }
    }
}

/// Parse serialized XML and deep-copy the root element into the target document as a child of parent
fn copy_element_from_xml(target_doc: &mut XmlDocument, parent: NodeId, element_xml: &str) -> Option<NodeId> {
    // DEBUG: Check what we're parsing
    if element_xml.len() < 100 {
        eprintln!("WARN copy_element_from_xml: element_xml very short: {}", element_xml);
    }
    // Use the namespace wrapper parser and attach to parent
    let node = parse_xml_fragment(target_doc, element_xml)?;
    // DEBUG: Count children
    let child_count = target_doc.descendants(node).count();
    eprintln!("DEBUG copy_element_from_xml: parsed node has {} descendants", child_count);
    if child_count == 0 {
        eprintln!("  WARNING: No descendants! element_xml first 200 chars: {}", &element_xml.chars().take(200).collect::<String>());
    }
    target_doc.reparent(parent, node);
    Some(node)
}

/// Deep-copy a node from source_doc to target_doc as a child of parent
fn deep_copy_node(target_doc: &mut XmlDocument, parent: NodeId, source_doc: &XmlDocument, source_node: NodeId) -> Option<NodeId> {
    let source_data = source_doc.get(source_node)?;

    // Clone the node data
    let new_data = source_data.clone();
    let new_node = target_doc.add_child(parent, new_data);

    // Recursively copy children
    for child in source_doc.children(source_node) {
        deep_copy_node(target_doc, new_node, source_doc, child);
    }

    Some(new_node)
}

fn reconstruct_math_elements(doc: &mut XmlDocument, parent: NodeId, _ancestor: &AncestorElementInfo, atoms: &[ComparisonUnitAtom], grouped_children: &[(String, usize, usize)], settings: &WmlComparerSettings) {
    for (_key, start, end) in grouped_children {
        let group_atoms = &atoms[*start..*end];
        let first = &group_atoms[0];
        let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
        let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
        let date_str = settings.date_time_for_revisions.as_deref().unwrap_or("");
        if del {
            for gcc in group_atoms {
                let del_elem = doc.add_child(parent, XmlNodeData::element_with_attrs(W::del(), vec![
                    XAttribute::new(W::id(), "0"),  // w:id MUST come first per ECMA-376
                    XAttribute::new(W::author(), settings.author_for_revisions.as_deref().unwrap_or("Unknown")),
                    XAttribute::new(W::date(), date_str),
                    // Add w16du:dateUtc for modern Word timezone handling
                    XAttribute::new(W16DU::dateUtc(), date_str),
                ]));
                if let Some(content) = create_content_element(doc, gcc, "") { doc.reparent(del_elem, content); }
            }
        } else if ins {
            for gcc in group_atoms {
                let ins_elem = doc.add_child(parent, XmlNodeData::element_with_attrs(W::ins(), vec![
                    XAttribute::new(W::id(), "0"),  // w:id MUST come first per ECMA-376
                    XAttribute::new(W::author(), settings.author_for_revisions.as_deref().unwrap_or("Unknown")),
                    XAttribute::new(W::date(), date_str),
                    // Add w16du:dateUtc for modern Word timezone handling
                    XAttribute::new(W16DU::dateUtc(), date_str),
                ]));
                if let Some(content) = create_content_element(doc, gcc, "") { doc.reparent(ins_elem, content); }
            }
        } else {
            for gcc in group_atoms {
                if let Some(content) = create_content_element(doc, gcc, "") { doc.reparent(parent, content); }
            }
        }
    }
}

fn reconstruct_allowable_run_children(doc: &mut XmlDocument, parent: NodeId, _ancestor: &AncestorElementInfo, atoms: &[ComparisonUnitAtom], grouped_children: &[(String, usize, usize)]) {
    for (_key, start, end) in grouped_children {
        let group_atoms = &atoms[*start..*end];
        let first = &group_atoms[0];
        let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
        let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
        // Use create_content_element for ALL cases to preserve element-specific attributes
        // (e.g., w:id for footnoteReference/endnoteReference). Pass status for del/ins cases.
        let status = if del { "Deleted" } else if ins { "Inserted" } else { "" };
        for gcc in group_atoms {
            if let Some(content) = create_content_element(doc, gcc, status) { 
                doc.reparent(parent, content); 
            }
        }
    }
}

fn reconstruct_element(doc: &mut XmlDocument, parent: NodeId, group_key: &str, ancestor: &AncestorElementInfo, _props_names: &[&str], group_atoms: &[ComparisonUnitAtom], level: usize, part: Option<()>, settings: &WmlComparerSettings) {
    // Check for uniform status
    let uniform_status = get_uniform_status(group_atoms);
    
    let (container, atoms_vec) = if let Some(status) = uniform_status {
        let wrapper = create_revision_wrapper(doc, parent, status, settings).unwrap();
        let mut new_atoms = group_atoms.to_vec();
        for atom in &mut new_atoms {
            atom.correlation_status = ComparisonCorrelationStatus::Equal;
        }
        (wrapper, Some(new_atoms))
    } else {
        (parent, None)
    };
    
    let atoms_ref = atoms_vec.as_deref().unwrap_or(group_atoms);
    
    let temp_container = doc.new_node(XmlNodeData::element(W::body()));
    coalesce_recurse(doc, temp_container, atoms_ref, level + 1, part, settings);
    let new_child_elements: Vec<NodeId> = doc.children(temp_container).collect();
    let mut attrs = ancestor.attributes.clone();
    attrs.retain(|a| a.name.namespace.as_deref() != Some(PT_STATUS_NS));
    attrs.push(XAttribute::new(pt_unid(), group_key));
    let elem_name = xname_from_ancestor(ancestor);
    let elem = doc.add_child(container, XmlNodeData::element_with_attrs(elem_name, attrs));
    for child in new_child_elements { doc.reparent(elem, child); }
}

fn reconstruct_vml_element(doc: &mut XmlDocument, parent: NodeId, group_key: &str, ancestor: &AncestorElementInfo, group_atoms: &[ComparisonUnitAtom], level: usize, part: Option<()>, settings: &WmlComparerSettings) {
    reconstruct_element(doc, parent, group_key, ancestor, &[], group_atoms, level, part, settings);
}

fn move_last_sect_pr_to_child_of_body(doc: &mut XmlDocument, doc_root: NodeId) {
    let body = doc.children(doc_root).find(|&child| doc.get(child).and_then(|d| d.name()).map(|n| n == &W::body()).unwrap_or(false));
    if body.is_none() { return; }
    let body = body.unwrap();
    let mut last_para_with_sect_pr: Option<NodeId> = None;
    let mut sect_pr_node: Option<NodeId> = None;
    for para in doc.children(body) {
        if doc.get(para).and_then(|d| d.name()).map(|n| n == &W::p()).unwrap_or(false) {
            for ppr in doc.children(para) {
                if doc.get(ppr).and_then(|d| d.name()).map(|n| n == &W::pPr()).unwrap_or(false) {
                    for sp in doc.children(ppr) {
                        if doc.get(sp).and_then(|d| d.name()).map(|n| n == &W::sectPr()).unwrap_or(false) {
                            last_para_with_sect_pr = Some(para);
                            sect_pr_node = Some(sp);
                        }
                    }
                }
            }
        }
    }
    if let (Some(_para), Some(sect_pr)) = (last_para_with_sect_pr, sect_pr_node) { doc.reparent(body, sect_pr); }
}

fn create_content_element(doc: &mut XmlDocument, atom: &ComparisonUnitAtom, status: &str) -> Option<NodeId> {
    match &atom.content_element {
        ContentElement::Text(_ch) => {
            // Text atoms should go through reconstruct_text_elements, not create_content_element.
            // Creating w:t directly here would place it under w:p without a w:r wrapper,
            // which is invalid OOXML and causes MS Word to report corruption.
            // Return None and let the atom be handled by the proper text reconstruction path.
            None
        }
        ContentElement::Break => {
            let br = doc.new_node(XmlNodeData::element(W::br()));
            if !status.is_empty() { doc.set_attribute(br, &pt_status(), status); }
            Some(br)
        }
        ContentElement::CarriageReturn => {
            let cr = doc.new_node(XmlNodeData::element(W::cr()));
            if !status.is_empty() { doc.set_attribute(cr, &pt_status(), status); }
            Some(cr)
        }
        ContentElement::Tab => {
            let tab = doc.new_node(XmlNodeData::element(W::tab()));
            if !status.is_empty() { doc.set_attribute(tab, &pt_status(), status); }
            Some(tab)
        }
        ContentElement::PositionalTab { alignment, relative_to, leader } => {
            let mut attrs = Vec::new();
            if !alignment.is_empty() { attrs.push(XAttribute::new(W::alignment(), alignment)); }
            if !relative_to.is_empty() { attrs.push(XAttribute::new(W::relativeTo(), relative_to)); }
            if !leader.is_empty() { attrs.push(XAttribute::new(W::leader(), leader)); }
            if !status.is_empty() { attrs.push(XAttribute::new(pt_status(), status)); }
            Some(doc.new_node(XmlNodeData::element_with_attrs(W::ptab(), attrs)))
        }
        ContentElement::ParagraphProperties { element_xml } => {
            // Parse and deep-copy the pPr element from stored XML
            // This preserves pStyle, jc, rPr, and other children
            if element_xml.is_empty() {
                // Fallback: create empty pPr if no stored XML
                let mut attrs = Vec::new();
                if !status.is_empty() { attrs.push(XAttribute::new(pt_status(), status)); }
                Some(doc.new_node(XmlNodeData::element_with_attrs(W::pPr(), attrs)))
            } else {
                // Parse and deep-copy the original pPr with all children using namespace wrapper
                if let Some(new_node) = parse_xml_fragment(doc, element_xml) {
                    if !status.is_empty() { doc.set_attribute(new_node, &pt_status(), status); }
                    Some(new_node)
                } else {
                    // Fallback: create empty pPr
                    let mut attrs = Vec::new();
                    if !status.is_empty() { attrs.push(XAttribute::new(pt_status(), status)); }
                    Some(doc.new_node(XmlNodeData::element_with_attrs(W::pPr(), attrs)))
                }
            }
        }
        ContentElement::Drawing { element_xml, .. } => {
            // Parse and deep-copy the drawing element from stored XML
            if element_xml.is_empty() {
                eprintln!("BUG: Drawing element_xml is EMPTY in create_content_element!");
                // Fallback: create empty drawing if no stored XML
                let drawing = doc.new_node(XmlNodeData::element(W::drawing()));
                if !status.is_empty() { doc.set_attribute(drawing, &pt_status(), status); }
                Some(drawing)
            } else {
                eprintln!("DEBUG: Drawing element_xml length={}, contains 'inline'={}", element_xml.len(), element_xml.contains("inline"));
                // Parse and deep-copy the original drawing using namespace wrapper
                if let Some(new_node) = parse_xml_fragment(doc, element_xml) {
                    let child_count = doc.descendants(new_node).count();
                    eprintln!("DEBUG: parse_xml_fragment returned node with {} descendants", child_count);
                    if child_count == 0 {
                        eprintln!("BUG: parse_xml_fragment returned node with NO children!");
                        eprintln!("  element_xml first 300 chars: {}", &element_xml.chars().take(300).collect::<String>());
                    }
                    if !status.is_empty() { doc.set_attribute(new_node, &pt_status(), status); }
                    Some(new_node)
                } else {
                    eprintln!("BUG: parse_xml_fragment returned None!");
                    eprintln!("  element_xml first 300 chars: {}", &element_xml.chars().take(300).collect::<String>());
                    // Fallback: create empty drawing
                    let drawing = doc.new_node(XmlNodeData::element(W::drawing()));
                    if !status.is_empty() { doc.set_attribute(drawing, &pt_status(), status); }
                    Some(drawing)
                }
            }
        }
        ContentElement::Picture { element_xml, .. } => {
            // Parse and deep-copy the pict element from stored XML
            if element_xml.is_empty() {
                let pict = doc.new_node(XmlNodeData::element(W::pict()));
                if !status.is_empty() { doc.set_attribute(pict, &pt_status(), status); }
                Some(pict)
            } else {
                // Parse and deep-copy the original pict using namespace wrapper
                if let Some(new_node) = parse_xml_fragment(doc, element_xml) {
                    if !status.is_empty() { doc.set_attribute(new_node, &pt_status(), status); }
                    return Some(new_node);
                }
                let pict = doc.new_node(XmlNodeData::element(W::pict()));
                if !status.is_empty() { doc.set_attribute(pict, &pt_status(), status); }
                Some(pict)
            }
        }
        ContentElement::Math { element_xml, .. } => {
            // Parse and deep-copy the math element from stored XML
            if element_xml.is_empty() {
                let math = doc.new_node(XmlNodeData::element(XName::new("http://schemas.openxmlformats.org/officeDocument/2006/math", "oMath")));
                if !status.is_empty() { doc.set_attribute(math, &pt_status(), status); }
                Some(math)
            } else {
                if let Ok(source_doc) = parse(element_xml) {
                    if let Some(source_root) = source_doc.root() {
                        let temp_parent = doc.new_node(XmlNodeData::element(XName::local("temp")));
                        if let Some(new_node) = deep_copy_node(doc, temp_parent, &source_doc, source_root) {
                            doc.detach(new_node);
                            if !status.is_empty() { doc.set_attribute(new_node, &pt_status(), status); }
                            return Some(new_node);
                        }
                    }
                }
                let math = doc.new_node(XmlNodeData::element(XName::new("http://schemas.openxmlformats.org/officeDocument/2006/math", "oMath")));
                if !status.is_empty() { doc.set_attribute(math, &pt_status(), status); }
                Some(math)
            }
        }
        ContentElement::CommentRangeStart { id } => {
            let mut attrs = vec![XAttribute::new(W::id(), id)];
            if !status.is_empty() { attrs.push(XAttribute::new(pt_status(), status)); }
            Some(doc.new_node(XmlNodeData::element_with_attrs(W::commentRangeStart(), attrs)))
        }
        ContentElement::CommentRangeEnd { id } => {
            let mut attrs = vec![XAttribute::new(W::id(), id)];
            if !status.is_empty() { attrs.push(XAttribute::new(pt_status(), status)); }
            Some(doc.new_node(XmlNodeData::element_with_attrs(W::commentRangeEnd(), attrs)))
        }
        ContentElement::FootnoteReference { id } => {
            let mut attrs = vec![XAttribute::new(W::id(), id)];
            if !status.is_empty() { attrs.push(XAttribute::new(pt_status(), status)); }
            Some(doc.new_node(XmlNodeData::element_with_attrs(W::footnoteReference(), attrs)))
        }
        ContentElement::EndnoteReference { id } => {
            let mut attrs = vec![XAttribute::new(W::id(), id)];
            if !status.is_empty() { attrs.push(XAttribute::new(pt_status(), status)); }
            Some(doc.new_node(XmlNodeData::element_with_attrs(W::endnoteReference(), attrs)))
        }
        ContentElement::Symbol { font, char_code } => {
            let mut attrs = Vec::new();
            if !font.is_empty() { attrs.push(XAttribute::new(XName::new(W::NS, "font"), font)); }
            if !char_code.is_empty() { attrs.push(XAttribute::new(XName::new(W::NS, "char"), char_code)); }
            if !status.is_empty() { attrs.push(XAttribute::new(pt_status(), status)); }
            Some(doc.new_node(XmlNodeData::element_with_attrs(W::sym(), attrs)))
        }
        _ => None,
    }
}

/// Group items by key, returning ranges into the original slice to avoid cloning
/// Returns (key, start_index, end_index) tuples
fn group_by_key_ranges<'a, T, F, K>(items: &'a [T], mut key_fn: F) -> Vec<(K, usize, usize)> 
where 
    F: FnMut(&T) -> K, 
    K: Eq + std::hash::Hash + Clone 
{
    if items.is_empty() {
        return Vec::new();
    }
    
    // Build index groups - maps key to list of indices
    let mut groups: HashMap<K, Vec<usize>> = HashMap::new();
    let mut order: Vec<K> = Vec::new();
    
    for (idx, item) in items.iter().enumerate() {
        let key = key_fn(item);
        if !groups.contains_key(&key) { 
            order.push(key.clone()); 
        }
        groups.entry(key).or_default().push(idx);
    }
    
    // Convert to contiguous ranges - this works because we process in order
    // and items with same key at same ancestor level are adjacent in practice
    let mut result = Vec::with_capacity(order.len());
    for key in order {
        if let Some(indices) = groups.remove(&key) {
            if !indices.is_empty() {
                // Find contiguous ranges
                let mut range_start = indices[0];
                let mut range_end = indices[0] + 1;
                
                for &idx in &indices[1..] {
                    if idx == range_end {
                        range_end = idx + 1;
                    } else {
                        // Non-contiguous - emit current range and start new one
                        result.push((key.clone(), range_start, range_end));
                        range_start = idx;
                        range_end = idx + 1;
                    }
                }
                result.push((key, range_start, range_end));
            }
        }
    }
    result
}

/// Group adjacent atoms by correlation status, returning ranges into the original slice
/// Returns (key, start_index, end_index) tuples - NO CLONING
fn group_adjacent_by_correlation_ranges(
    atoms: &[ComparisonUnitAtom],
    level: usize,
    _is_inside_vml: bool,
    settings: &WmlComparerSettings,
) -> Vec<(String, usize, usize)> {
    if atoms.is_empty() {
        return Vec::new();
    }
    
    let mut groups: Vec<(String, usize, usize)> = Vec::new();
    let mut current_key = String::new();
    let mut range_start = 0usize;
    
    for (idx, atom) in atoms.iter().enumerate() {
        let in_txbx_content = atom.ancestor_elements.iter().take(level).any(|a| a.local_name == "txbxContent");
        let ancestor_unid = if level < atom.ancestor_unids.len() - 1 { 
            if in_txbx_content { "TXBX" } else { &atom.ancestor_unids[level + 1] }
        } else { 
            "" 
        };
        
        let key = if in_txbx_content { 
            format!("{}|Equal", ancestor_unid) 
        } else {
            let status_str = format!("{:?}", atom.correlation_status);
            if settings.track_formatting_changes {
                if atom.correlation_status == ComparisonCorrelationStatus::FormatChanged {
                    format!("{}|{}|FMT:{}|TO:{}", ancestor_unid, status_str, 
                        atom.formatting_change_rpr_before_signature.as_deref().unwrap_or("<null>"), 
                        atom.formatting_signature.as_deref().unwrap_or("<null>"))
                } else if atom.correlation_status == ComparisonCorrelationStatus::Equal {
                    format!("{}|{}|SIG:{}", ancestor_unid, status_str, 
                        atom.formatting_signature.as_deref().unwrap_or("<null>"))
                } else if atom.correlation_status == ComparisonCorrelationStatus::Inserted 
                       || atom.correlation_status == ComparisonCorrelationStatus::Deleted {
                    // For Inserted/Deleted, don't include ancestor_unid to allow coalescing
                    // of contiguous revisions from different source runs
                    format!("{}|SIG:{}", status_str, 
                        atom.formatting_signature.as_deref().unwrap_or("<null>"))
                } else { 
                    // Keep ancestor_unid for Equal content to preserve source structure
                    format!("{}|{}", ancestor_unid, status_str) 
                }
            } else if atom.correlation_status == ComparisonCorrelationStatus::Inserted 
                   || atom.correlation_status == ComparisonCorrelationStatus::Deleted {
                // For Inserted/Deleted, don't include ancestor_unid to allow coalescing
                // of contiguous revisions from different source runs
                format!("{}|SIG:{}", status_str, 
                    atom.formatting_signature.as_deref().unwrap_or("<null>"))
            } else { 
                // Keep ancestor_unid for Equal content to preserve source structure
                format!("{}|{}", ancestor_unid, status_str) 
            }
        };
        
        if idx == 0 {
            current_key = key;
            range_start = 0;
        } else if key != current_key {
            // End previous group, start new one
            groups.push((std::mem::take(&mut current_key), range_start, idx));
            current_key = key;
            range_start = idx;
        }
    }
    
    // Push final group
    if !atoms.is_empty() {
        groups.push((current_key, range_start, atoms.len()));
    }
    
    groups
}

fn get_ancestor_element_for_level(
    first_atom: &ComparisonUnitAtom,
    level: usize,
    group_atoms: &[ComparisonUnitAtom],
) -> AncestorElementInfo {
    let mut is_inside_vml = false;
    for i in 0..=level {
        if i < first_atom.ancestor_elements.len() {
            if is_vml_related_element(&first_atom.ancestor_elements[i].local_name) {
                is_inside_vml = true;
                break;
            }
        }
    }
    if is_inside_vml {
        for atom in group_atoms {
            if let Some(ref before_ancestors) = atom.ancestor_elements_before {
                if level < before_ancestors.len() {
                    return AncestorElementInfo {
                        namespace: before_ancestors[level].namespace.clone(),
                        local_name: before_ancestors[level].local_name.clone(),
                        attributes: before_ancestors[level].attributes.as_ref().clone(),
                        rpr_xml: before_ancestors[level].rpr_xml.clone(),
                    };
                }
            }
        }
    }
    AncestorElementInfo {
        namespace: first_atom.ancestor_elements[level].namespace.clone(),
        local_name: first_atom.ancestor_elements[level].local_name.clone(),
        attributes: first_atom.ancestor_elements[level].attributes.as_ref().clone(),
        rpr_xml: first_atom.ancestor_elements[level].rpr_xml.clone(),
    }
}

fn xname_from_ancestor(ancestor: &AncestorElementInfo) -> XName {
    match ancestor.namespace.as_deref() {
        Some(namespace) if !namespace.is_empty() => XName::new(namespace, &ancestor.local_name),
        _ => XName::local(&ancestor.local_name),
    }
}

fn is_vml_related_element(name: &str) -> bool {
    VML_RELATED_ELEMENTS.contains(&name)
}

fn is_inside_vml_content(atom: &ComparisonUnitAtom, level: usize) -> bool {
    for i in 0..=level {
        if i < atom.ancestor_elements.len() {
            if is_vml_related_element(&atom.ancestor_elements[i].local_name) {
                return true;
            }
        }
    }
    false
}

fn needs_xml_space(text: &str) -> bool {
    if text.is_empty() { return false; }
    let chars: Vec<char> = text.chars().collect();
    chars[0].is_whitespace() || chars[chars.len() - 1].is_whitespace()
}

fn get_descendant_status(doc: &XmlDocument, node: NodeId) -> Option<String> {
    for desc in doc.descendants(node) {
        if let Some(data) = doc.get(desc) {
            if let Some(attrs) = data.attributes() {
                if let Some(attr) = attrs.iter().find(|a| a.name == pt_status()) {
                    return Some(attr.value.clone());
                }
            }
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::parser;

    #[test]
    fn test_is_vml_related_element() {
        assert!(is_vml_related_element("pict"));
        assert!(is_vml_related_element("shape"));
        assert!(!is_vml_related_element("p"));
    }

    #[test]
    fn test_needs_xml_space() {
        assert!(needs_xml_space(" hello"));
        assert!(needs_xml_space("hello "));
        assert!(needs_xml_space(" hello "));
        assert!(!needs_xml_space("hello"));
        assert!(!needs_xml_space("hello world"));
        assert!(!needs_xml_space(""));
    }

    #[test]
    fn test_needs_xml_space_preserve() {
        assert!(needs_xml_space_preserve(" hello"));
        assert!(needs_xml_space_preserve("hello "));
        assert!(needs_xml_space_preserve(" hello "));
        assert!(!needs_xml_space_preserve("hello"));
        assert!(!needs_xml_space_preserve("hello world"));
        assert!(!needs_xml_space_preserve(""));
    }

    #[test]
    fn test_format_date_for_key() {
        // Standard ISO date with timezone
        assert_eq!(format_date_for_key("2023-01-15T10:30:00Z"), "2023-01-15T10:30:00");
        // Date with offset
        assert_eq!(format_date_for_key("2023-01-15T10:30:00+05:00"), "2023-01-15T10:30:00");
        // Short date
        assert_eq!(format_date_for_key("2023-01-15"), "2023-01-15");
        // Empty
        assert_eq!(format_date_for_key(""), "");
    }

    #[test]
    fn test_group_adjacent_by_key() {
        let mut doc = XmlDocument::new();
        let r1 = doc.add_root(XmlNodeData::element(W::r()));
        let r2 = doc.new_node(XmlNodeData::element(W::r()));
        let r3 = doc.new_node(XmlNodeData::element(W::r()));
        
        let children = vec![r1, r2, r3];
        let keys = vec!["KeyA".to_string(), "KeyA".to_string(), "KeyB".to_string()];
        
        let groups = group_adjacent_by_key(&children, &keys);
        
        assert_eq!(groups.len(), 2);
        assert_eq!(groups[0].0, "KeyA");
        assert_eq!(groups[0].1.len(), 2);
        assert_eq!(groups[1].0, "KeyB");
        assert_eq!(groups[1].1.len(), 1);
    }

    #[test]
    fn test_get_consolidation_key_simple_run_with_t() {
        let mut doc = XmlDocument::new();
        let run = doc.add_root(XmlNodeData::element(W::r()));
        let t = doc.add_child(run, XmlNodeData::element(W::t()));
        doc.add_child(t, XmlNodeData::Text("Hello".to_string()));
        
        let key = get_consolidation_key(&doc, run);
        
        assert!(key.starts_with("Wt"));
        assert_ne!(key, DONT_CONSOLIDATE);
    }

    #[test]
    fn test_get_consolidation_key_run_with_rpr_and_t() {
        let mut doc = XmlDocument::new();
        let run = doc.add_root(XmlNodeData::element(W::r()));
        let rpr = doc.add_child(run, XmlNodeData::element(W::rPr()));
        let b = doc.add_child(rpr, XmlNodeData::element(XName::new(W::NS, "b")));
        let t = doc.add_child(run, XmlNodeData::element(W::t()));
        doc.add_child(t, XmlNodeData::Text("Hello".to_string()));
        
        let key = get_consolidation_key(&doc, run);
        
        assert!(key.starts_with("Wt"));
        assert!(key.contains("<w:b"));
        assert_ne!(key, DONT_CONSOLIDATE);
    }

    #[test]
    fn test_get_consolidation_key_run_with_instr_text() {
        let mut doc = XmlDocument::new();
        let run = doc.add_root(XmlNodeData::element(W::r()));
        let it = doc.add_child(run, XmlNodeData::element(W::instrText()));
        doc.add_child(it, XmlNodeData::Text("PAGEREF".to_string()));
        
        let key = get_consolidation_key(&doc, run);
        
        assert!(key.starts_with("WinstrText"));
        assert_ne!(key, DONT_CONSOLIDATE);
    }

    #[test]
    fn test_get_consolidation_key_run_with_multiple_children_returns_dont() {
        let mut doc = XmlDocument::new();
        let run = doc.add_root(XmlNodeData::element(W::r()));
        let t = doc.add_child(run, XmlNodeData::element(W::t()));
        doc.add_child(t, XmlNodeData::Text("Hello".to_string()));
        doc.add_child(run, XmlNodeData::element(W::br()));
        
        let key = get_consolidation_key(&doc, run);
        
        assert_eq!(key, DONT_CONSOLIDATE);
    }

    #[test]
    fn test_get_consolidation_key_ins_element() {
        let mut doc = XmlDocument::new();
        let ins = doc.add_root(XmlNodeData::element_with_attrs(W::ins(), vec![
            XAttribute::new(W::id(), "1"),  // w:id MUST come first per ECMA-376
            XAttribute::new(W::author(), "Test Author"),
            XAttribute::new(W::date(), "2023-01-15T10:30:00Z"),
        ]));
        let run = doc.add_child(ins, XmlNodeData::element(W::r()));
        let t = doc.add_child(run, XmlNodeData::element(W::t()));
        doc.add_child(t, XmlNodeData::Text("inserted".to_string()));
        
        let key = get_consolidation_key(&doc, ins);
        
        assert!(key.starts_with("Wins2"));
        assert!(key.contains("Test Author"));
        assert!(key.contains("2023-01-15T10:30:00"));
        assert!(key.contains("1"));
        assert_ne!(key, DONT_CONSOLIDATE);
    }

    #[test]
    fn test_get_consolidation_key_del_element() {
        let mut doc = XmlDocument::new();
        let del = doc.add_root(XmlNodeData::element_with_attrs(W::del(), vec![
            XAttribute::new(W::id(), "1"),  // w:id MUST come first per ECMA-376
            XAttribute::new(W::author(), "Test Author"),
            XAttribute::new(W::date(), "2023-01-15T10:30:00Z"),
        ]));
        let run = doc.add_child(del, XmlNodeData::element(W::r()));
        let dt = doc.add_child(run, XmlNodeData::element(W::delText()));
        doc.add_child(dt, XmlNodeData::Text("deleted".to_string()));
        
        let key = get_consolidation_key(&doc, del);
        
        assert!(key.starts_with("Wdel"));
        assert!(key.contains("Test Author"));
        assert!(key.contains("2023-01-15T10:30:00"));
        // Note: w:del does NOT include id in key (unlike w:ins)
        assert_ne!(key, DONT_CONSOLIDATE);
    }

    #[test]
    fn test_get_consolidation_key_ins_with_del_child_returns_dont() {
        let mut doc = XmlDocument::new();
        let ins = doc.add_root(XmlNodeData::element(W::ins()));
        doc.add_child(ins, XmlNodeData::element(W::del()));
        
        let key = get_consolidation_key(&doc, ins);
        
        assert_eq!(key, DONT_CONSOLIDATE);
    }

    #[test]
    fn test_coalesce_adjacent_runs_merges_identical_runs() {
        let mut doc = XmlDocument::new();
        let para = doc.add_root(XmlNodeData::element(W::p()));
        
        // Create two runs with same formatting
        let r1 = doc.add_child(para, XmlNodeData::element(W::r()));
        let t1 = doc.add_child(r1, XmlNodeData::element(W::t()));
        doc.add_child(t1, XmlNodeData::Text("Hello".to_string()));
        
        let r2 = doc.add_child(para, XmlNodeData::element(W::r()));
        let t2 = doc.add_child(r2, XmlNodeData::element(W::t()));
        doc.add_child(t2, XmlNodeData::Text(" World".to_string()));
        
        // Before coalescing: 2 runs
        assert_eq!(doc.children(para).count(), 2);
        
        // Coalesce
        coalesce_adjacent_runs_with_identical_formatting(&mut doc, para);
        
        // After coalescing: 1 run
        let children: Vec<_> = doc.children(para).collect();
        assert_eq!(children.len(), 1);
        
        // The merged run should contain "Hello World"
        let merged_run = children[0];
        let t_elem = doc.children(merged_run).find(|&c| is_element_named(&doc, c, W::NS, "t"));
        assert!(t_elem.is_some());
        
        let text_node = doc.children(t_elem.unwrap()).next();
        assert!(text_node.is_some());
        if let Some(XmlNodeData::Text(text)) = doc.get(text_node.unwrap()) {
            assert_eq!(text, "Hello World");
        } else {
            panic!("Expected text node");
        }
    }

    #[test]
    fn test_coalesce_does_not_merge_different_formatting() {
        let mut doc = XmlDocument::new();
        let para = doc.add_root(XmlNodeData::element(W::p()));
        
        // Create run with bold formatting
        let r1 = doc.add_child(para, XmlNodeData::element(W::r()));
        let rpr1 = doc.add_child(r1, XmlNodeData::element(W::rPr()));
        doc.add_child(rpr1, XmlNodeData::element(XName::new(W::NS, "b")));
        let t1 = doc.add_child(r1, XmlNodeData::element(W::t()));
        doc.add_child(t1, XmlNodeData::Text("Bold".to_string()));
        
        // Create run without formatting
        let r2 = doc.add_child(para, XmlNodeData::element(W::r()));
        let t2 = doc.add_child(r2, XmlNodeData::element(W::t()));
        doc.add_child(t2, XmlNodeData::Text("Normal".to_string()));
        
        // Before coalescing: 2 runs
        assert_eq!(doc.children(para).count(), 2);
        
        // Coalesce
        coalesce_adjacent_runs_with_identical_formatting(&mut doc, para);
        
        // After coalescing: still 2 runs (different formatting)
        assert_eq!(doc.children(para).count(), 2);
    }

    #[test]
    fn test_serialize_element_for_key() {
        let mut doc = XmlDocument::new();
        let rpr = doc.add_root(XmlNodeData::element(W::rPr()));
        let b = doc.add_child(rpr, XmlNodeData::element(XName::new(W::NS, "b")));
        
        let serialized = serialize_element_for_key(&doc, rpr);
        
        assert!(serialized.contains("w:rPr"));
        assert!(serialized.contains("w:b"));
    }

    #[test]
    fn test_escape_xml_attr() {
        assert_eq!(escape_xml_attr("hello"), "hello");
        assert_eq!(escape_xml_attr("a&b"), "a&amp;b");
        assert_eq!(escape_xml_attr("a<b"), "a&lt;b");
        assert_eq!(escape_xml_attr("a>b"), "a&gt;b");
        assert_eq!(escape_xml_attr("a\"b"), "a&quot;b");
    }

    #[test]
    fn test_collect_text_from_group() {
        let mut doc = XmlDocument::new();
        let r1 = doc.add_root(XmlNodeData::element(W::r()));
        let t1 = doc.add_child(r1, XmlNodeData::element(W::t()));
        doc.add_child(t1, XmlNodeData::Text("Hello".to_string()));
        
        let r2 = doc.new_node(XmlNodeData::element(W::r()));
        let t2 = doc.add_child(r2, XmlNodeData::element(W::t()));
        doc.add_child(t2, XmlNodeData::Text(" World".to_string()));
        
        let group = vec![r1, r2];
        let text = collect_text_from_group(&doc, &group, "Wt");
        
        assert_eq!(text, "Hello World");
    }

    #[test]
    fn test_parse_rpr_xml_with_bold() {
        let mut doc = XmlDocument::new();
        let rpr_xml = r#"<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:b/></w:rPr>"#;
        
        let rpr_node = parse_rpr_xml(&mut doc, rpr_xml);
        
        assert!(rpr_node.is_some());
        let rpr_node = rpr_node.unwrap();
        
        // Check that the node is an rPr element
        let data = doc.get(rpr_node).unwrap();
        let name = data.name().unwrap();
        assert_eq!(name.local_name, "rPr");
        assert_eq!(name.namespace.as_deref(), Some(W::NS));
        
        // Check that it has a w:b child
        let children: Vec<_> = doc.children(rpr_node).collect();
        assert_eq!(children.len(), 1);
        let b_data = doc.get(children[0]).unwrap();
        let b_name = b_data.name().unwrap();
        assert_eq!(b_name.local_name, "b");
    }

    #[test]
    fn test_parse_rpr_xml_with_multiple_properties() {
        let mut doc = XmlDocument::new();
        let rpr_xml = r#"<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:b/><w:szCs w:val="24"/></w:rPr>"#;
        
        let rpr_node = parse_rpr_xml(&mut doc, rpr_xml);
        
        assert!(rpr_node.is_some());
        let rpr_node = rpr_node.unwrap();
        
        // Check that it has two children (w:b and w:szCs)
        let children: Vec<_> = doc.children(rpr_node).collect();
        assert_eq!(children.len(), 2);
        
        // First child should be w:b
        let first_data = doc.get(children[0]).unwrap();
        let first_name = first_data.name().unwrap();
        assert_eq!(first_name.local_name, "b");
        
        // Second child should be w:szCs with val="24"
        let second_data = doc.get(children[1]).unwrap();
        let second_name = second_data.name().unwrap();
        assert_eq!(second_name.local_name, "szCs");
        let attrs = second_data.attributes().unwrap();
        let val_attr = attrs.iter().find(|a| a.name.local_name == "val");
        assert!(val_attr.is_some());
        assert_eq!(val_attr.unwrap().value, "24");
    }

    #[test]
    fn test_ancestor_element_info_rpr_xml_field() {
        // Test that AncestorElementInfo can hold rpr_xml
        let ancestor = AncestorElementInfo {
            namespace: Some(W::NS.to_string()),
            local_name: "r".to_string(),
            attributes: vec![],
            rpr_xml: Some("<w:rPr><w:b/></w:rPr>".to_string()),
        };
        
        assert!(ancestor.rpr_xml.is_some());
        assert!(ancestor.rpr_xml.as_ref().unwrap().contains("w:b"));
    }

    #[test]
    fn test_remove_empty_rpr_elements() {
        let mut doc = XmlDocument::new();
        
        // Create a structure with some empty and non-empty rPr elements
        let root = doc.add_root(XmlNodeData::element(W::body()));
        let para = doc.add_child(root, XmlNodeData::element(W::p()));
        
        // Run 1: has empty rPr (should be removed)
        let run1 = doc.add_child(para, XmlNodeData::element(W::r()));
        let _empty_rpr = doc.add_child(run1, XmlNodeData::element(W::rPr()));
        let t1 = doc.add_child(run1, XmlNodeData::element(W::t()));
        doc.add_child(t1, XmlNodeData::Text("Hello".to_string()));
        
        // Run 2: has non-empty rPr (should be kept)
        let run2 = doc.add_child(para, XmlNodeData::element(W::r()));
        let nonempty_rpr = doc.add_child(run2, XmlNodeData::element(W::rPr()));
        doc.add_child(nonempty_rpr, XmlNodeData::element(XName::new(W::NS, "b")));
        let t2 = doc.add_child(run2, XmlNodeData::element(W::t()));
        doc.add_child(t2, XmlNodeData::Text("World".to_string()));
        
        // Run 3: has another empty rPr (should be removed)
        let run3 = doc.add_child(para, XmlNodeData::element(W::r()));
        let _empty_rpr2 = doc.add_child(run3, XmlNodeData::element(W::rPr()));
        let t3 = doc.add_child(run3, XmlNodeData::element(W::t()));
        doc.add_child(t3, XmlNodeData::Text("!".to_string()));
        
        // Count rPr elements before
        let rpr_count_before: usize = std::iter::once(root)
            .chain(doc.descendants(root))
            .filter(|&n| {
                doc.get(n)
                    .and_then(|d| d.name())
                    .map(|name| name.namespace.as_deref() == Some(W::NS) && name.local_name == "rPr")
                    .unwrap_or(false)
            })
            .count();
        assert_eq!(rpr_count_before, 3);
        
        // Run the cleanup
        remove_empty_rpr_elements(&mut doc, root);
        
        // Count rPr elements after
        let rpr_count_after: usize = std::iter::once(root)
            .chain(doc.descendants(root))
            .filter(|&n| {
                doc.get(n)
                    .and_then(|d| d.name())
                    .map(|name| name.namespace.as_deref() == Some(W::NS) && name.local_name == "rPr")
                    .unwrap_or(false)
            })
            .count();
        
        // Should only have 1 rPr left (the non-empty one)
        assert_eq!(rpr_count_after, 1);
    }

    #[test]
    fn test_parse_rpr_xml_returns_none_for_empty() {
        let mut doc = XmlDocument::new();
        
        // Empty rPr element should return None
        let rpr_xml = r#"<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        let result = parse_rpr_xml(&mut doc, rpr_xml);
        assert!(result.is_none(), "Empty rPr should return None");
        
        // Empty string should return None
        let result2 = parse_rpr_xml(&mut doc, "");
        assert!(result2.is_none(), "Empty string should return None");
    }
}
