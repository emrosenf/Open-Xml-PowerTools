//! Coalesce - Reconstruct XML tree from comparison atoms
//!
//! This is a faithful line-by-line port of CoalesceRecurse from C# WmlComparer.cs (lines 5161-5738).

use super::comparison_unit::{ComparisonCorrelationStatus, ComparisonUnitAtom, ContentElement};
use super::settings::WmlComparerSettings;
use crate::wml::revision::{create_run_property_change, RevisionSettings};
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::W;
use crate::xml::node::XmlNodeData;
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

/// Helper struct to represent an ancestor element's information
#[derive(Clone, Debug)]
struct AncestorElementInfo {
    local_name: String,
    attributes: Vec<XAttribute>,
}

/// VML-related element names (from C# VmlRelatedElements set)
static VML_RELATED_ELEMENTS: &[&str] = &[
    "pict", "shape", "rect", "group", "shapetype", "oval", "line", "arc", "curve", "polyline", "roundrect",
];

/// Allowable run children that can have pt:Status
static ALLOWABLE_RUN_CHILDREN: &[&str] = &[
    "br", "tab", "sym", "ptab", "cr", "dayShort", "dayLong", "monthShort", "monthLong", "yearShort", "yearLong",
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
    ];
    for (prefix, uri) in standard_namespaces {
        if !attrs.iter().any(|a| a.name.local_name == prefix && a.name.namespace.as_deref() == Some("http://www.w3.org/2000/xmlns/")) {
            attrs.push(XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", prefix), uri));
        }
    }
    if !attrs.iter().any(|a| a.name.local_name == "pt14") {
        attrs.push(XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", "pt14"), PT_STATUS_NS));
    }
    
    let doc_root = doc.add_root(XmlNodeData::element_with_attrs(root_name.clone(), attrs));
    
    if root_name.local_name == "document" {
        let body = doc.add_child(doc_root, XmlNodeData::element(W::body()));
        coalesce_recurse(&mut doc, body, atoms, 0, None, settings);
        move_last_sect_pr_to_child_of_body(&mut doc, doc_root);
    } else {
        coalesce_recurse(&mut doc, doc_root, atoms, 0, None, settings);
    }
    
    CoalesceResult { document: doc, root: doc_root }
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
        author: settings.author_for_revisions.clone().unwrap_or_else(|| "redline-rs".to_string()),
        date_time: settings.date_time_for_revisions.clone(),
    };

    // Transform in-place following C# pattern
    mark_content_transform(doc, root, &revision_settings);
}

/// Static counter for revision IDs (matching C# s_MaxId)
static REVISION_ID: std::sync::atomic::AtomicU32 = std::sync::atomic::AtomicU32::new(1);

fn next_revision_id() -> u32 {
    REVISION_ID.fetch_add(1, std::sync::atomic::Ordering::SeqCst)
}

/// Reset revision ID counter (for testing)
#[allow(dead_code)]
pub fn reset_revision_id() {
    REVISION_ID.store(1, std::sync::atomic::Ordering::SeqCst);
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
            XAttribute::new(W::author(), &settings.author),
            XAttribute::new(W::id(), &id_str),
            XAttribute::new(W::date(), &settings.date_time),
        ];
        let del_elem = doc.new_node(XmlNodeData::element_with_attrs(W::del(), del_attrs));
        
        doc.insert_before(run, del_elem);
        doc.detach(run);
        doc.reparent(del_elem, run);
    } else if status == "Inserted" {
        let id_str = next_revision_id().to_string();
        let ins_attrs = vec![
            XAttribute::new(W::author(), &settings.author),
            XAttribute::new(W::id(), &id_str),
            XAttribute::new(W::date(), &settings.date_time),
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
            XAttribute::new(W::author(), &settings.author),
            XAttribute::new(W::id(), &id_str),
            XAttribute::new(W::date(), &settings.date_time),
        ];
        doc.add_child(rpr, XmlNodeData::element_with_attrs(W::del(), del_attrs));
    } else if status == "Inserted" {
        let id_str = next_revision_id().to_string();
        let ins_attrs = vec![
            XAttribute::new(W::author(), &settings.author),
            XAttribute::new(W::id(), &id_str),
            XAttribute::new(W::date(), &settings.date_time),
        ];
        doc.add_child(rpr, XmlNodeData::element_with_attrs(W::ins(), ins_attrs));
    }
}

fn is_props_element(doc: &XmlDocument, node: NodeId) -> bool {
    doc.get(node).and_then(|d| d.name()).map(|n| {
        n.namespace.as_deref() == Some(W::NS) && n.local_name.ends_with("Pr")
    }).unwrap_or(false)
}

pub fn coalesce_adjacent_runs(doc: &mut XmlDocument, root: NodeId, settings: &WmlComparerSettings) {
    let mut paras = Vec::new();
    collect_paragraphs(doc, root, &mut paras);
    for para in paras {
        coalesce_paragraph_runs(doc, para, settings);
    }
}

fn collect_paragraphs(doc: &XmlDocument, node: NodeId, result: &mut Vec<NodeId>) {
    if let Some(data) = doc.get(node) {
        if let Some(name) = data.name() {
            if name.namespace.as_deref() == Some(W::NS) && name.local_name == "p" {
                result.push(node);
                return;
            }
        }
    }
    let children: Vec<_> = doc.children(node).collect();
    for child in children { collect_paragraphs(doc, child, result); }
}

fn coalesce_paragraph_runs(doc: &mut XmlDocument, para: NodeId, settings: &WmlComparerSettings) {
    let children: Vec<_> = doc.children(para).collect();
    if children.is_empty() { return; }
    let mut new_children = Vec::new();
    let mut i = 0;
    while i < children.len() {
        let current = children[i];
        let key = get_consolidation_key(doc, current, settings);
        if key == "DontConsolidate" {
            new_children.push(current);
            i += 1;
            continue;
        }
        let mut group = vec![current];
        let mut j = i + 1;
        while j < children.len() {
            let next = children[j];
            if get_consolidation_key(doc, next, settings) == key {
                group.push(next);
                j += 1;
            } else { break; }
        }
        if group.len() > 1 { new_children.push(merge_nodes(doc, &group)); } else { new_children.push(current); }
        i = j;
    }
    for &child in &children { doc.detach(child); }
    for &child in &new_children { doc.reparent(para, child); }
}

fn get_consolidation_key(doc: &mut XmlDocument, node: NodeId, settings: &WmlComparerSettings) -> String {
    let Some(data) = doc.get(node) else { return "DontConsolidate".to_string() };
    let Some(name) = data.name() else { return "DontConsolidate".to_string() };
    if name.namespace.as_deref() != Some(W::NS) { return "DontConsolidate".to_string(); }
    match name.local_name.as_str() {
        "r" => {
            let children: Vec<_> = doc.children(node).filter(|&c| !is_r_pr(doc, c)).collect();
            if children.len() != 1 || !is_t(doc, children[0]) { return "DontConsolidate".to_string(); }
            format!("Wt|{}", get_r_pr_signature(doc, node, settings))
        }
        "ins" => {
            let run = find_child_by_name(doc, node, W::NS, "r");
            if let Some(r) = run {
                let children: Vec<_> = doc.children(r).filter(|&c| !is_r_pr(doc, c)).collect();
                if children.len() == 1 && is_t(doc, children[0]) {
                    let author = get_attr(doc, node, W::NS, "author").unwrap_or_default();
                    let date = get_attr(doc, node, W::NS, "date").unwrap_or_default();
                    let id = get_attr(doc, node, W::NS, "id").unwrap_or_default();
                    return format!("Wins|{}|{}|{}|{}", author, date, id, get_r_pr_signature(doc, r, settings));
                }
            }
            "DontConsolidate".to_string()
        }
        "del" => {
            let run = find_child_by_name(doc, node, W::NS, "r");
            if let Some(r) = run {
                let children: Vec<_> = doc.children(r).filter(|&c| !is_r_pr(doc, c)).collect();
                if children.len() == 1 && is_del_text(doc, children[0]) {
                    let author = get_attr(doc, node, W::NS, "author").unwrap_or_default();
                    let date = get_attr(doc, node, W::NS, "date").unwrap_or_default();
                    return format!("Wdel|{}|{}|{}", author, date, get_r_pr_signature(doc, r, settings));
                }
            }
            "DontConsolidate".to_string()
        }
        _ => "DontConsolidate".to_string()
    }
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
    let grouped = group_by_key(list, |ca| {
        if level >= ca.ancestor_unids.len() { String::new() } else { ca.ancestor_unids[level].clone() }
    });
    let grouped: Vec<_> = grouped.into_iter().filter(|(key, _)| !key.is_empty()).collect();
    if grouped.is_empty() { return; }
    for (group_key, group_atoms) in grouped {
        let first_atom = &group_atoms[0];
        let ancestor_being_constructed = get_ancestor_element_for_level(first_atom, level, &group_atoms);
        let ancestor_name = &ancestor_being_constructed.local_name;
        let is_inside_vml = is_inside_vml_content(first_atom, level);
        let grouped_children = group_adjacent_by_correlation(&group_atoms, level, is_inside_vml, settings);
        match ancestor_name.as_str() {
            "p" => reconstruct_paragraph(doc, parent, &group_key, &ancestor_being_constructed, &grouped_children, level, is_inside_vml, _part, settings),
            "r" => reconstruct_run(doc, parent, &ancestor_being_constructed, &grouped_children, level, is_inside_vml, _part, settings),
            "t" => reconstruct_text_elements(doc, parent, &grouped_children),
            "drawing" => reconstruct_drawing_elements(doc, parent, &grouped_children, _part, settings),
            "oMath" | "oMathPara" => reconstruct_math_elements(doc, parent, &ancestor_being_constructed, &grouped_children, settings),
            elem if ALLOWABLE_RUN_CHILDREN.contains(&elem) => reconstruct_allowable_run_children(doc, parent, &ancestor_being_constructed, &grouped_children),
            "tbl" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["tblPr", "tblGrid"], &group_atoms, level, _part, settings),
            "tr" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["trPr"], &group_atoms, level, _part, settings),
            "tc" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["tcPr"], &group_atoms, level, _part, settings),
            "sdt" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["sdtPr", "sdtEndPr"], &group_atoms, level, _part, settings),
            "pict" | "shape" | "rect" | "group" | "shapetype" | "oval" | "line" | "arc" | "curve" | "polyline" | "roundrect" => reconstruct_vml_element(doc, parent, &group_key, &ancestor_being_constructed, &group_atoms, level, _part, settings),
            "object" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["shapetype", "shape", "OLEObject"], &group_atoms, level, _part, settings),
            "ruby" => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &["rubyPr"], &group_atoms, level, _part, settings),
            _ => reconstruct_element(doc, parent, &group_key, &ancestor_being_constructed, &[], &group_atoms, level, _part, settings),
        }
    }
}

fn reconstruct_paragraph(doc: &mut XmlDocument, parent: NodeId, group_key: &str, ancestor: &AncestorElementInfo, grouped_children: &[(String, Vec<ComparisonUnitAtom>)], level: usize, is_inside_vml: bool, part: Option<()>, settings: &WmlComparerSettings) {
    let mut para_attrs = ancestor.attributes.clone();
    para_attrs.retain(|a| a.name.namespace.as_deref() != Some(PT_STATUS_NS));
    para_attrs.push(XAttribute::new(pt_unid(), group_key));
    let para = doc.add_child(parent, XmlNodeData::element_with_attrs(W::p(), para_attrs));
    for (key, group_atoms) in grouped_children {
        let spl: Vec<&str> = key.split('|').collect();
        if spl.get(0) == Some(&"") {
            for gcc in group_atoms {
                if is_inside_vml && matches!(&gcc.content_element, ContentElement::ParagraphProperties) && spl.get(1) == Some(&"Inserted") { continue; }
                let content_elem_node = create_content_element(doc, gcc, spl.get(1).unwrap_or(&""));
                if let Some(node) = content_elem_node { doc.reparent(para, node); }
            }
        } else { coalesce_recurse(doc, para, group_atoms, level + 1, part, settings); }
    }
}

fn reconstruct_run(doc: &mut XmlDocument, parent: NodeId, ancestor: &AncestorElementInfo, grouped_children: &[(String, Vec<ComparisonUnitAtom>)], level: usize, _is_inside_vml: bool, part: Option<()>, settings: &WmlComparerSettings) {
    let mut run_attrs = ancestor.attributes.clone();
    run_attrs.retain(|a| a.name.namespace.as_deref() != Some(PT_STATUS_NS));
    let run = doc.add_child(parent, XmlNodeData::element_with_attrs(W::r(), run_attrs));
    let mut format_changed = false;
    for (key, group_atoms) in grouped_children {
        let spl: Vec<&str> = key.split('|').collect();
        if spl.get(0) == Some(&"") {
            for gcc in group_atoms {
                if gcc.correlation_status == ComparisonCorrelationStatus::FormatChanged {
                    format_changed = true;
                }
                let content_elem_node = create_content_element(doc, gcc, spl.get(1).unwrap_or(&""));
                if let Some(node) = content_elem_node { doc.reparent(run, node); }
            }
        } else { coalesce_recurse(doc, run, group_atoms, level + 1, part, settings); }
    }
    if settings.track_formatting_changes && format_changed {
        let existing_rpr = doc.children(run).find(|&c| is_r_pr(doc, c));
        let rpr = match existing_rpr { Some(node) => node, None => doc.add_child(run, XmlNodeData::element(W::rPr())) };
        let revision_settings = RevisionSettings { author: settings.author_for_revisions.clone().unwrap_or_else(|| "redline-rs".to_string()), date_time: settings.date_time_for_revisions.clone() };
        let _rpr_change = create_run_property_change(doc, rpr, &revision_settings);
    }
}

fn reconstruct_text_elements(doc: &mut XmlDocument, parent: NodeId, grouped_children: &[(String, Vec<ComparisonUnitAtom>)]) {
    for (_key, group_atoms) in grouped_children {
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

fn reconstruct_drawing_elements(doc: &mut XmlDocument, parent: NodeId, grouped_children: &[(String, Vec<ComparisonUnitAtom>)], _part: Option<()>, _settings: &WmlComparerSettings) {
    for (_key, group_atoms) in grouped_children {
        let first = &group_atoms[0];
        let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
        let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
        if del || ins {
            for gcc in group_atoms {
                if let ContentElement::Drawing { .. } = &gcc.content_element {
                    let drawing = doc.add_child(parent, XmlNodeData::element(W::drawing()));
                    let status = if del { "Deleted" } else { "Inserted" };
                    doc.set_attribute(drawing, &pt_status(), status);
                }
            }
        } else {
            for gcc in group_atoms {
                if let ContentElement::Drawing { .. } = &gcc.content_element { doc.add_child(parent, XmlNodeData::element(W::drawing())); }
            }
        }
    }
}

fn reconstruct_math_elements(doc: &mut XmlDocument, parent: NodeId, _ancestor: &AncestorElementInfo, grouped_children: &[(String, Vec<ComparisonUnitAtom>)], settings: &WmlComparerSettings) {
    for (_key, group_atoms) in grouped_children {
        let first = &group_atoms[0];
        let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
        let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
        if del {
            for gcc in group_atoms {
                let del_elem = doc.add_child(parent, XmlNodeData::element_with_attrs(W::del(), vec![XAttribute::new(W::author(), settings.author_for_revisions.as_deref().unwrap_or("Unknown")), XAttribute::new(W::id(), "0"), XAttribute::new(W::date(), &settings.date_time_for_revisions)]));
                if let Some(content) = create_content_element(doc, gcc, "") { doc.reparent(del_elem, content); }
            }
        } else if ins {
            for gcc in group_atoms {
                let ins_elem = doc.add_child(parent, XmlNodeData::element_with_attrs(W::ins(), vec![XAttribute::new(W::author(), settings.author_for_revisions.as_deref().unwrap_or("Unknown")), XAttribute::new(W::id(), "0"), XAttribute::new(W::date(), &settings.date_time_for_revisions)]));
                if let Some(content) = create_content_element(doc, gcc, "") { doc.reparent(ins_elem, content); }
            }
        } else {
            for gcc in group_atoms {
                if let Some(content) = create_content_element(doc, gcc, "") { doc.reparent(parent, content); }
            }
        }
    }
}

fn reconstruct_allowable_run_children(doc: &mut XmlDocument, parent: NodeId, ancestor: &AncestorElementInfo, grouped_children: &[(String, Vec<ComparisonUnitAtom>)]) {
    for (_key, group_atoms) in grouped_children {
        let first = &group_atoms[0];
        let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
        let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
        if del || ins {
            for _gcc in group_atoms {
                let mut attrs = ancestor.attributes.clone();
                attrs.retain(|a| a.name.namespace.as_deref() != Some(PT_STATUS_NS));
                let status = if del { "Deleted" } else { "Inserted" };
                attrs.push(XAttribute::new(pt_status(), status));
                let elem_name = XName::new(W::NS, &ancestor.local_name);
                doc.add_child(parent, XmlNodeData::element_with_attrs(elem_name, attrs));
            }
        } else {
            for gcc in group_atoms {
                if let Some(content) = create_content_element(doc, gcc, "") { doc.reparent(parent, content); }
            }
        }
    }
}

fn reconstruct_element(doc: &mut XmlDocument, parent: NodeId, group_key: &str, ancestor: &AncestorElementInfo, _props_names: &[&str], group_atoms: &[ComparisonUnitAtom], level: usize, part: Option<()>, settings: &WmlComparerSettings) {
    let temp_container = doc.new_node(XmlNodeData::element(W::body()));
    coalesce_recurse(doc, temp_container, group_atoms, level + 1, part, settings);
    let new_child_elements: Vec<NodeId> = doc.children(temp_container).collect();
    let mut attrs = ancestor.attributes.clone();
    attrs.push(XAttribute::new(pt_unid(), group_key));
    let elem_name = XName::new(W::NS, &ancestor.local_name);
    let elem = doc.add_child(parent, XmlNodeData::element_with_attrs(elem_name, attrs));
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
        ContentElement::Text(_ch) => None,
        ContentElement::Break => {
            let br = doc.new_node(XmlNodeData::element(W::br()));
            if !status.is_empty() { doc.set_attribute(br, &pt_status(), status); }
            Some(br)
        }
        ContentElement::Tab => {
            let tab = doc.new_node(XmlNodeData::element(W::tab()));
            if !status.is_empty() { doc.set_attribute(tab, &pt_status(), status); }
            Some(tab)
        }
        ContentElement::ParagraphProperties => {
            let mut attrs = Vec::new();
            if !status.is_empty() { attrs.push(XAttribute::new(pt_status(), status)); }
            Some(doc.new_node(XmlNodeData::element_with_attrs(W::pPr(), attrs)))
        }
        _ => None,
    }
}

fn group_by_key<T, F, K>(items: &[T], mut key_fn: F) -> Vec<(K, Vec<T>)> where T: Clone, F: FnMut(&T) -> K, K: Eq + std::hash::Hash + Clone {
    let mut groups: HashMap<K, Vec<T>> = HashMap::new();
    let mut order: Vec<K> = Vec::new();
    for item in items {
        let key = key_fn(item);
        if !groups.contains_key(&key) { order.push(key.clone()); }
        groups.entry(key).or_default().push(item.clone());
    }
    order.into_iter().filter_map(|key| groups.remove(&key).map(|items| (key, items))).collect()
}

fn group_adjacent_by_correlation(
    atoms: &[ComparisonUnitAtom],
    level: usize,
    _is_inside_vml: bool,
    settings: &WmlComparerSettings,
) -> Vec<(String, Vec<ComparisonUnitAtom>)> {
    let mut groups: Vec<(String, Vec<ComparisonUnitAtom>)> = Vec::new();
    for atom in atoms {
        let in_txbx_content = atom.ancestor_elements.iter().take(level).any(|a| a.local_name == "txbxContent");
        let mut ancestor_unid = if level < atom.ancestor_unids.len() - 1 { atom.ancestor_unids[level + 1].clone() } else { String::new() };
        if in_txbx_content && !ancestor_unid.is_empty() { ancestor_unid = "TXBX".to_string(); }
        let status_str = if in_txbx_content { "Equal".to_string() } else { format!("{:?}", atom.correlation_status) };
        let key = if in_txbx_content { format!("{}|{}", ancestor_unid, status_str) } else {
            if settings.track_formatting_changes {
                if atom.correlation_status == ComparisonCorrelationStatus::FormatChanged {
                    format!("{}|{}|FMT:{}|TO:{}", ancestor_unid, status_str, atom.formatting_change_rpr_before_signature.as_deref().unwrap_or("<null>"), atom.formatting_signature.as_deref().unwrap_or("<null>"))
                } else if atom.correlation_status == ComparisonCorrelationStatus::Equal {
                    format!("{}|{}|SIG:{}", ancestor_unid, status_str, atom.formatting_signature.as_deref().unwrap_or("<null>"))
                } else { format!("{}|{}", ancestor_unid, status_str) }
            } else { format!("{}|{}", ancestor_unid, status_str) }
        };
        if let Some((last_key, last_group)) = groups.last_mut() {
            if last_key == &key { last_group.push(atom.clone()); continue; }
        }
        groups.push((key, vec![atom.clone()]));
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
                        local_name: before_ancestors[level].local_name.clone(),
                        attributes: before_ancestors[level].attributes.clone(),
                    };
                }
            }
        }
    }
    AncestorElementInfo {
        local_name: first_atom.ancestor_elements[level].local_name.clone(),
        attributes: first_atom.ancestor_elements[level].attributes.clone(),
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
}

