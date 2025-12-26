//! Coalesce - Reconstruct XML tree from comparison atoms
//!
//! This is a faithful line-by-line port of CoalesceRecurse from C# WmlComparer.cs (lines 5161-5738).
//!
//! CRITICAL METHODOLOGY:
//! 1. Translation-first: Port C# code line-by-line, preserving structure and naming
//! 2. No behavior inference: Never guess what C# does from test outputs
//! 3. Tests validate, not guide: Run tests only AFTER implementation complete
//! 4. When stuck: Read C# source, not test expectations
//! 5. Naming convention: Rust functions mirror C# method names
//!
//! The algorithm from C# lines 5161-5738:
//! 1. Group atoms by their AncestorUnids[level] (line 5163-5169)
//! 2. Filter out empty keys (line 5169)
//! 3. For each group, select appropriate element reconstruction (lines 5194-5596)
//! 4. GroupAdjacent by: ancestorUnid[level+1] + correlationStatus + formattingSignature (lines 5231-5290)
//! 5. Special handling for txbxContent: use "TXBX" marker and force Equal status (lines 5236-5267)
//! 6. Special handling for VML: prefer "before" ancestors (lines 5198-5227)
//! 7. Recurse for nested elements (line 5321, 5357, etc.)
//!
//! Key features from C# that MUST be ported exactly:
//! - txbxContent grouping: lines 5236-5267
//! - VML content handling: lines 5198-5227
//! - Formatting signature grouping: lines 5272-5288
//! - GroupAdjacent pattern: lines 5231-5290
//! - Element-specific reconstruction: W.p (5292-5332), W.r (5334-5399), W.t (5401-5426),
//!   W.drawing (5428-5495), M.oMath (5497-5533), AllowableRunChildren (5535-5569),
//!   W.tbl/tr/tc/sdt (5571-5578), VML elements (5579-5590), W._object (5591-5592),
//!   W.ruby (5593-5594), generic elements (5595)

use super::comparison_unit::{ComparisonCorrelationStatus, ComparisonUnitAtom, ContentElement};
use super::settings::WmlComparerSettings;
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

/// VML-related element names (from C# VmlRelatedElements set)
/// See C# lines 7829-7832
static VML_RELATED_ELEMENTS: &[&str] = &[
    "pict",      // W.pict
    "shape",     // VML.shape
    "rect",      // VML.rect
    "group",     // VML.group
    "shapetype", // VML.shapetype
    "oval",      // VML.oval
    "line",      // VML.line
    "arc",       // VML.arc
    "curve",     // VML.curve
    "polyline",  // VML.polyline
    "roundrect", // VML.roundrect
];

/// Allowable run children that can have pt:Status (from C# AllowableRunChildren)
/// See C# lines 5535-5569
static ALLOWABLE_RUN_CHILDREN: &[&str] = &[
    "br",
    "tab",
    "sym",
    "ptab",
    "cr",
    "dayShort",
    "dayLong",
    "monthShort",
    "monthLong",
    "yearShort",
    "yearLong",
];

/// Port of C# Coalesce() entry point (lines 7977-7991)
///
/// Creates a new document with w:document root, populates w:body from atoms,
/// then runs cleanup (MoveLastSectPrToChildOfBody).
pub struct CoalesceResult {
    pub document: XmlDocument,
    pub root: NodeId,
}

pub fn coalesce(atoms: &[ComparisonUnitAtom], settings: &WmlComparerSettings) -> CoalesceResult {
    let mut doc = XmlDocument::new();
    
    // Create w:document root with namespaces (C# line 7981-7983)
    let doc_root = doc.add_root(XmlNodeData::element_with_attrs(
        W::document(),
        vec![
            XAttribute::new(
                XName::new("http://www.w3.org/2000/xmlns/", "w"),
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            ),
            XAttribute::new(
                XName::new("http://www.w3.org/2000/xmlns/", "pt14"),
                PT_STATUS_NS,
            ),
        ],
    ));
    
    // Create w:body (C# line 7984)
    let body = doc.add_child(doc_root, XmlNodeData::element(W::body()));
    
    // Coalesce atoms into body children (C# line 7980)
    coalesce_recurse(&mut doc, body, atoms, 0, None, settings);
    
    // Cleanup: move last sectPr to child of body (C# line 7987)
    move_last_sect_pr_to_child_of_body(&mut doc, doc_root);
    
    // TODO: C# line 7988-7989: WmlOrderElementsPerStandard
    // For now, we skip this as it's a separate concern
    
    CoalesceResult {
        document: doc,
        root: doc_root,
    }
}

/// Port of C# CoalesceRecurse (lines 5161-5599)
///
/// This is the core recursive tree reconstruction function.
/// CRITICAL: This must match C# line-by-line to ensure identical behavior.
///
/// Parameters:
/// - doc: The XML document being built
/// - parent: The parent node to append children to
/// - list: The list of ComparisonUnitAtom to process
/// - level: Current tree level (0 = deepest ancestor, increases as we descend)
/// - part: OpenXmlPart (Rust: not used yet, pass None)
/// - settings: Comparer settings
fn coalesce_recurse(
    doc: &mut XmlDocument,
    parent: NodeId,
    list: &[ComparisonUnitAtom],
    level: usize,
    _part: Option<()>, // TODO: port OpenXmlPart handling for drawings
    settings: &WmlComparerSettings,
) {
    // C# lines 5163-5169: Group by AncestorUnids[level], filter out empty keys
    let grouped = group_by_key(list, |ca| {
        if level >= ca.ancestor_unids.len() {
            String::new()
        } else {
            ca.ancestor_unids[level].clone()
        }
    });
    
    let grouped: Vec<_> = grouped
        .into_iter()
        .filter(|(key, _)| !key.is_empty())
        .collect();
    
    // C# lines 5171-5173: If no deeper children, return null
    if grouped.is_empty() {
        return;
    }
    
    // C# lines 5194-5598: Process each group
    for (group_key, group_atoms) in grouped {
        let first_atom = &group_atoms[0];
        
        // C# lines 5198-5227: VML handling - prefer "before" ancestors
        let ancestor_being_constructed = get_ancestor_element_for_level(first_atom, level, &group_atoms);
        
        let ancestor_name = &ancestor_being_constructed.local_name;
        
        // Determine if we're inside VML content (C# lines 5202-5210)
        let is_inside_vml = is_inside_vml_content(first_atom, level);
        
        // C# lines 5229-5290: GroupAdjacent by child unid, status, and formatting signature
        let grouped_children = group_adjacent_by_correlation(
            &group_atoms,
            level,
            is_inside_vml,
            settings,
        );
        
        // Reconstruct the appropriate element type
        match ancestor_name.as_str() {
            // C# lines 5292-5332: W.p
            "p" => {
                reconstruct_paragraph(
                    doc,
                    parent,
                    &group_key,
                    &ancestor_being_constructed,
                    &grouped_children,
                    level,
                    is_inside_vml,
                    _part,
                    settings,
                );
            }
            // C# lines 5334-5399: W.r
            "r" => {
                reconstruct_run(
                    doc,
                    parent,
                    &ancestor_being_constructed,
                    &grouped_children,
                    level,
                    is_inside_vml,
                    _part,
                    settings,
                );
            }
            // C# lines 5401-5426: W.t (text elements)
            "t" => {
                reconstruct_text_elements(
                    doc,
                    parent,
                    &grouped_children,
                );
            }
            // C# lines 5428-5495: W.drawing
            "drawing" => {
                reconstruct_drawing_elements(
                    doc,
                    parent,
                    &grouped_children,
                    _part,
                    settings,
                );
            }
            // C# lines 5497-5533: M.oMath and M.oMathPara
            "oMath" | "oMathPara" => {
                reconstruct_math_elements(
                    doc,
                    parent,
                    &ancestor_being_constructed,
                    &grouped_children,
                    settings,
                );
            }
            // C# lines 5535-5569: AllowableRunChildren
            elem if ALLOWABLE_RUN_CHILDREN.contains(&elem) => {
                reconstruct_allowable_run_children(
                    doc,
                    parent,
                    &ancestor_being_constructed,
                    &grouped_children,
                );
            }
            // C# lines 5571-5578: Table elements
            "tbl" => {
                reconstruct_element(
                    doc,
                    parent,
                    &group_key,
                    &ancestor_being_constructed,
                    &["tblPr", "tblGrid"],
                    &group_atoms,
                    level,
                    _part,
                    settings,
                );
            }
            "tr" => {
                reconstruct_element(
                    doc,
                    parent,
                    &group_key,
                    &ancestor_being_constructed,
                    &["trPr"],
                    &group_atoms,
                    level,
                    _part,
                    settings,
                );
            }
            "tc" => {
                reconstruct_element(
                    doc,
                    parent,
                    &group_key,
                    &ancestor_being_constructed,
                    &["tcPr"],
                    &group_atoms,
                    level,
                    _part,
                    settings,
                );
            }
            "sdt" => {
                reconstruct_element(
                    doc,
                    parent,
                    &group_key,
                    &ancestor_being_constructed,
                    &["sdtPr", "sdtEndPr"],
                    &group_atoms,
                    level,
                    _part,
                    settings,
                );
            }
            // C# lines 5579-5590: VML elements
            "pict" | "shape" | "rect" | "group" | "shapetype" | "oval" | "line" | "arc" | "curve" | "polyline" | "roundrect" => {
                reconstruct_vml_element(
                    doc,
                    parent,
                    &group_key,
                    &ancestor_being_constructed,
                    &group_atoms,
                    level,
                    _part,
                    settings,
                );
            }
            // C# lines 5591-5592: W._object
            "object" => {
                reconstruct_element(
                    doc,
                    parent,
                    &group_key,
                    &ancestor_being_constructed,
                    &["shapetype", "shape", "OLEObject"],
                    &group_atoms,
                    level,
                    _part,
                    settings,
                );
            }
            // C# lines 5593-5594: W.ruby
            "ruby" => {
                reconstruct_element(
                    doc,
                    parent,
                    &group_key,
                    &ancestor_being_constructed,
                    &["rubyPr"],
                    &group_atoms,
                    level,
                    _part,
                    settings,
                );
            }
            // C# line 5595: Generic element
            _ => {
                reconstruct_element(
                    doc,
                    parent,
                    &group_key,
                    &ancestor_being_constructed,
                    &[],
                    &group_atoms,
                    level,
                    _part,
                    settings,
                );
            }
        }
    }
}

/// Helper struct to represent an ancestor element's information
#[derive(Clone, Debug)]
struct AncestorElementInfo {
    local_name: String,
    attributes: Vec<XAttribute>,
}

/// Port of C# lines 5198-5227: Get the ancestor element for reconstruction
///
/// For VML content, prefer "before" ancestors to preserve original document structure.
/// This ensures proper round-trip when rejecting revisions.
fn get_ancestor_element_for_level(
    first_atom: &ComparisonUnitAtom,
    level: usize,
    group_atoms: &[ComparisonUnitAtom],
) -> AncestorElementInfo {
    // Check if ANY ancestor (not just current level) is VML-related (C# lines 5202-5210)
    let mut is_inside_vml = false;
    for i in 0..=level {
        if i < first_atom.ancestor_elements.len() {
            if is_vml_related_element(&first_atom.ancestor_elements[i].local_name) {
                is_inside_vml = true;
                break;
            }
        }
    }
    
    // C# lines 5212-5227: Try to find an atom with AncestorElementsBefore
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
    
    // Default: use first atom's current ancestors (C# line 5226)
    AncestorElementInfo {
        local_name: first_atom.ancestor_elements[level].local_name.clone(),
        attributes: first_atom.ancestor_elements[level].attributes.clone(),
    }
}

/// Check if an element name is VML-related (C# lines 7829-7832)
fn is_vml_related_element(name: &str) -> bool {
    VML_RELATED_ELEMENTS.contains(&name)
}

/// Check if an atom is inside VML content (C# lines 5202-5210)
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

/// Port of C# lines 5231-5290: GroupAdjacent by correlation key
///
/// Groups adjacent atoms by:
/// 1. ancestorUnid[level+1] (child unid) - with special "TXBX" marker for txbxContent
/// 2. correlationStatus - with special Equal forcing for txbxContent
/// 3. formattingSignature - if tracking formatting changes
///
/// Returns Vec<(key_string, Vec<atom>)> preserving adjacency order
fn group_adjacent_by_correlation(
    atoms: &[ComparisonUnitAtom],
    level: usize,
    is_inside_vml: bool,
    settings: &WmlComparerSettings,
) -> Vec<(String, Vec<ComparisonUnitAtom>)> {
    let mut groups: Vec<(String, Vec<ComparisonUnitAtom>)> = Vec::new();
    
    for atom in atoms {
        // C# lines 5234-5244: Check for txbxContent ancestor
        let in_txbx_content = {
            let mut found = false;
            for i in 0..level {
                if i < atom.ancestor_elements.len() {
                    if atom.ancestor_elements[i].local_name == "txbxContent" {
                        found = true;
                        break;
                    }
                }
            }
            found
        };
        
        // C# lines 5246-5257: Build grouping key
        let mut ancestor_unid = if level < atom.ancestor_unids.len() - 1 {
            atom.ancestor_unids[level + 1].clone()
        } else {
            String::new()
        };
        
        // C# lines 5254-5257: For txbxContent, use "TXBX" marker
        if in_txbx_content && !ancestor_unid.is_empty() {
            ancestor_unid = "TXBX".to_string();
        }
        
        let status_for_grouping = atom.correlation_status;
        let status_str = if in_txbx_content {
            "Equal"
        } else {
            match status_for_grouping {
                ComparisonCorrelationStatus::Equal => "Equal",
                ComparisonCorrelationStatus::Inserted => "Inserted",
                ComparisonCorrelationStatus::Deleted => "Deleted",
                ComparisonCorrelationStatus::FormatChanged => "FormatChanged",
                // Other statuses treated as unknown during grouping
                ComparisonCorrelationStatus::Nil 
                | ComparisonCorrelationStatus::Normal 
                | ComparisonCorrelationStatus::Unknown 
                | ComparisonCorrelationStatus::Group => "Unknown",
            }
        };
        
        // C# lines 5263-5288: For txbxContent, skip formatting signature logic
        let key = if in_txbx_content {
            format!("{}|{}", ancestor_unid, status_str)
        } else {
            // C# lines 5272-5288: Add formatting signature if tracking format changes
            if settings.track_formatting_changes {
                if atom.correlation_status == ComparisonCorrelationStatus::FormatChanged {
                    let before_sig = atom.formatting_change_rpr_before_signature.as_deref().unwrap_or("<null>");
                    let after_sig = atom.formatting_signature.as_deref().unwrap_or("<null>");
                    format!("{}|{}|FMT:{}|TO:{}", ancestor_unid, status_str, before_sig, after_sig)
                } else if atom.correlation_status == ComparisonCorrelationStatus::Equal {
                    let sig = atom.formatting_signature.as_deref().unwrap_or("<null>");
                    format!("{}|{}|SIG:{}", ancestor_unid, status_str, sig)
                } else {
                    format!("{}|{}", ancestor_unid, status_str)
                }
            } else {
                format!("{}|{}", ancestor_unid, status_str)
            }
        };
        
        // GroupAdjacent: append to last group if key matches, otherwise create new group
        if let Some((last_key, last_group)) = groups.last_mut() {
            if last_key == &key {
                last_group.push(atom.clone());
                continue;
            }
        }
        groups.push((key, vec![atom.clone()]));
    }
    
    groups
}

/// Port of C# lines 5292-5332: Reconstruct paragraph (W.p)
fn reconstruct_paragraph(
    doc: &mut XmlDocument,
    parent: NodeId,
    group_key: &str,
    ancestor: &AncestorElementInfo,
    grouped_children: &[(String, Vec<ComparisonUnitAtom>)],
    level: usize,
    is_inside_vml: bool,
    part: Option<()>,
    settings: &WmlComparerSettings,
) {
    // Create w:p element (C# lines 5326-5329)
    let mut para_attrs = ancestor.attributes.clone();
    para_attrs.retain(|a| a.name.namespace.as_deref() != Some(PT_STATUS_NS));
    para_attrs.push(XAttribute::new(pt_unid(), group_key));
    
    let para = doc.add_child(parent, XmlNodeData::element_with_attrs(W::p(), para_attrs));
    
    // Process grouped children (C# lines 5294-5324)
    for (key, group_atoms) in grouped_children {
        let spl: Vec<&str> = key.split('|').collect();
        
        // C# lines 5298-5318: If child_unid is empty, add content directly
        if spl.get(0) == Some(&"") {
            for gcc in group_atoms {
                // C# lines 5302-5305: Skip Inserted pPr in VML
                if is_inside_vml
                    && matches!(&gcc.content_element, ContentElement::ParagraphProperties)
                    && spl.get(1) == Some(&"Inserted")
                {
                    continue;
                }
                
                // C# lines 5307-5317: Use "before" content element for VML
                let content_elem_node = if is_inside_vml && gcc.content_element_before.is_some() {
                    // TODO: use content_element_before
                    create_content_element(doc, gcc, spl.get(1).unwrap_or(&""))
                } else {
                    create_content_element(doc, gcc, spl.get(1).unwrap_or(&""))
                };
                
                if let Some(node) = content_elem_node {
                    doc.reparent(para, node);
                }
            }
        } else {
            // C# line 5321: Recurse for deeper elements
            coalesce_recurse(doc, para, group_atoms, level + 1, part, settings);
        }
    }
}

/// Port of C# lines 5334-5399: Reconstruct run (W.r)
fn reconstruct_run(
    doc: &mut XmlDocument,
    parent: NodeId,
    ancestor: &AncestorElementInfo,
    grouped_children: &[(String, Vec<ComparisonUnitAtom>)],
    level: usize,
    is_inside_vml: bool,
    part: Option<()>,
    settings: &WmlComparerSettings,
) {
    // Create w:r element (C# lines 5366-5369)
    let mut run_attrs = ancestor.attributes.clone();
    run_attrs.retain(|a| a.name.namespace.as_deref() != Some(PT_STATUS_NS));
    
    let run = doc.add_child(parent, XmlNodeData::element_with_attrs(W::r(), run_attrs));
    
    // C# lines 5362-5364: Copy rPr if present
    // TODO: port rPr copying from ancestorBeingConstructed
    
    // Process grouped children (C# lines 5336-5360)
    for (key, group_atoms) in grouped_children {
        let spl: Vec<&str> = key.split('|').collect();
        
        if spl.get(0) == Some(&"") {
            for gcc in group_atoms {
                let content_elem_node = if is_inside_vml && gcc.content_element_before.is_some() {
                    create_content_element(doc, gcc, spl.get(1).unwrap_or(&""))
                } else {
                    create_content_element(doc, gcc, spl.get(1).unwrap_or(&""))
                };
                
                if let Some(node) = content_elem_node {
                    doc.reparent(run, node);
                }
            }
        } else {
            coalesce_recurse(doc, run, group_atoms, level + 1, part, settings);
        }
    }
    
    // C# lines 5371-5396: Add w:rPrChange if format changed
    if settings.track_formatting_changes {
        // TODO: port formatting change tracking
        // This requires checking group_atoms for FormattingChangeRPrBefore
        // and creating w:rPrChange element with before formatting
    }
}

/// Port of C# lines 5401-5426: Reconstruct text elements (W.t)
fn reconstruct_text_elements(
    doc: &mut XmlDocument,
    parent: NodeId,
    grouped_children: &[(String, Vec<ComparisonUnitAtom>)],
) {
    for (_key, group_atoms) in grouped_children {
        // C# lines 5406-5422: Concatenate text and create w:t or w:delText
        let text_of_text_element: String = group_atoms
            .iter()
            .filter_map(|gce| {
                if let ContentElement::Text(ch) = gce.content_element {
                    Some(ch)
                } else {
                    None
                }
            })
            .collect();
        
        if text_of_text_element.is_empty() {
            continue;
        }
        
        let first = &group_atoms[0];
        let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
        let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
        
        let elem_name = if del { W::delText() } else { W::t() };
        let mut attrs = Vec::new();
        
        // C# lines 5412, 5418, 5421: Add xml:space="preserve" if needed
        if needs_xml_space(&text_of_text_element) {
            attrs.push(XAttribute::new(
                XName::new("http://www.w3.org/XML/1998/namespace", "space"),
                "preserve",
            ));
        }
        
        // C# lines 5411, 5416: Add pt:Status
        if del {
            attrs.push(XAttribute::new(pt_status(), "Deleted"));
        } else if ins {
            attrs.push(XAttribute::new(pt_status(), "Inserted"));
        }
        
        let text_elem = if attrs.is_empty() {
            doc.add_child(parent, XmlNodeData::element(elem_name))
        } else {
            doc.add_child(parent, XmlNodeData::element_with_attrs(elem_name, attrs))
        };
        
        doc.add_child(text_elem, XmlNodeData::Text(text_of_text_element));
    }
}

/// Port of C# lines 5693-5699: GetXmlSpaceAttribute
fn needs_xml_space(text: &str) -> bool {
    if text.is_empty() {
        return false;
    }
    let chars: Vec<char> = text.chars().collect();
    chars[0].is_whitespace() || chars[chars.len() - 1].is_whitespace()
}

/// Port of C# lines 5428-5495: Reconstruct drawing elements
fn reconstruct_drawing_elements(
    doc: &mut XmlDocument,
    parent: NodeId,
    grouped_children: &[(String, Vec<ComparisonUnitAtom>)],
    _part: Option<()>, // TODO: port relationship handling
    _settings: &WmlComparerSettings,
) {
    for (_key, group_atoms) in grouped_children {
        let first = &group_atoms[0];
        let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
        let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
        
        if del || ins {
            for gcc in group_atoms {
                // Create drawing element with pt:Status
                if let ContentElement::Drawing { .. } = &gcc.content_element {
                    let drawing = doc.add_child(parent, XmlNodeData::element(W::drawing()));
                    let status = if del { "Deleted" } else { "Inserted" };
                    doc.set_attribute(drawing, &pt_status(), status);
                    
                    // TODO: C# lines 5442-5452, 5464-5472: MoveRelatedPartsToDestination
                    // This requires porting the relationship/package handling
                }
            }
        } else {
            // Equal status: still need to copy related parts (C# lines 5476-5491)
            for gcc in group_atoms {
                if let ContentElement::Drawing { .. } = &gcc.content_element {
                    let _drawing = doc.add_child(parent, XmlNodeData::element(W::drawing()));
                    // TODO: Copy related parts for Equal status too
                }
            }
        }
    }
}

/// Port of C# lines 5497-5533: Reconstruct math elements (M.oMath, M.oMathPara)
fn reconstruct_math_elements(
    doc: &mut XmlDocument,
    parent: NodeId,
    ancestor: &AncestorElementInfo,
    grouped_children: &[(String, Vec<ComparisonUnitAtom>)],
    settings: &WmlComparerSettings,
) {
    for (_key, group_atoms) in grouped_children {
        let first = &group_atoms[0];
        let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
        let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
        
        if del {
            // C# lines 5506-5513: Wrap in w:del
            for gcc in group_atoms {
                let del_elem = doc.add_child(
                    parent,
                    XmlNodeData::element_with_attrs(
                        W::del(),
                        vec![
                            XAttribute::new(W::author(), settings.author_for_revisions.as_deref().unwrap_or("Unknown")),
                            XAttribute::new(W::id(), "0"), // TODO: use s_MaxId++
                            XAttribute::new(W::date(), &settings.date_time_for_revisions),
                        ],
                    ),
                );
                // Add content element to del
                if let Some(content) = create_content_element(doc, gcc, "") {
                    doc.reparent(del_elem, content);
                }
            }
        } else if ins {
            // C# lines 5517-5524: Wrap in w:ins
            for gcc in group_atoms {
                let ins_elem = doc.add_child(
                    parent,
                    XmlNodeData::element_with_attrs(
                        W::ins(),
                        vec![
                            XAttribute::new(W::author(), settings.author_for_revisions.as_deref().unwrap_or("Unknown")),
                            XAttribute::new(W::id(), "0"), // TODO: use s_MaxId++
                            XAttribute::new(W::date(), &settings.date_time_for_revisions),
                        ],
                    ),
                );
                // Add content element to ins
                if let Some(content) = create_content_element(doc, gcc, "") {
                    doc.reparent(ins_elem, content);
                }
            }
        } else {
            // C# lines 5528: Just add content element
            for gcc in group_atoms {
                if let Some(content) = create_content_element(doc, gcc, "") {
                    doc.reparent(parent, content);
                }
            }
        }
    }
}

/// Port of C# lines 5535-5569: Reconstruct allowable run children
fn reconstruct_allowable_run_children(
    doc: &mut XmlDocument,
    parent: NodeId,
    ancestor: &AncestorElementInfo,
    grouped_children: &[(String, Vec<ComparisonUnitAtom>)],
) {
    for (_key, group_atoms) in grouped_children {
        let first = &group_atoms[0];
        let del = first.correlation_status == ComparisonCorrelationStatus::Deleted;
        let ins = first.correlation_status == ComparisonCorrelationStatus::Inserted;
        
        if del || ins {
            // C# lines 5544-5550, 5553-5559: Create element with pt:Status
            for _gcc in group_atoms {
                let mut attrs = ancestor.attributes.clone();
                attrs.retain(|a| a.name.namespace.as_deref() != Some(PT_STATUS_NS));
                let status = if del { "Deleted" } else { "Inserted" };
                attrs.push(XAttribute::new(pt_status(), status));
                
                let elem_name = XName::new(W::NS, &ancestor.local_name);
                doc.add_child(parent, XmlNodeData::element_with_attrs(elem_name, attrs));
            }
        } else {
            // C# line 5564: Just add content element
            for gcc in group_atoms {
                if let Some(content) = create_content_element(doc, gcc, "") {
                    doc.reparent(parent, content);
                }
            }
        }
    }
}

/// Port of C# lines 5701-5723: ReconstructElement (generic version)
fn reconstruct_element(
    doc: &mut XmlDocument,
    parent: NodeId,
    group_key: &str,
    ancestor: &AncestorElementInfo,
    props_names: &[&str],
    group_atoms: &[ComparisonUnitAtom],
    level: usize,
    part: Option<()>,
    settings: &WmlComparerSettings,
) {
    // Recurse to get child elements (C# line 5704)
    let temp_container = doc.add_child(parent, XmlNodeData::element(W::body())); // Temporary
    coalesce_recurse(doc, temp_container, group_atoms, level + 1, part, settings);
    let new_child_elements: Vec<NodeId> = doc.children(temp_container).collect();
    doc.detach(temp_container);
    
    // Create reconstructed element (C# lines 5718-5720)
    let mut attrs = ancestor.attributes.clone();
    attrs.push(XAttribute::new(pt_unid(), group_key));
    
    let elem_name = XName::new(W::NS, &ancestor.local_name);
    let elem = doc.add_child(parent, XmlNodeData::element_with_attrs(elem_name, attrs));
    
    // Add property elements first (C# lines 5706-5716)
    for prop_name in props_names {
        // TODO: Copy property elements from ancestorBeingConstructed
        // For now, we skip this as it requires more context about the original element
    }
    
    // Add child elements
    for child in new_child_elements {
        doc.reparent(elem, child);
    }
}

/// Port of C# lines 5736-5743: ReconstructVmlElement
fn reconstruct_vml_element(
    doc: &mut XmlDocument,
    parent: NodeId,
    group_key: &str,
    ancestor: &AncestorElementInfo,
    group_atoms: &[ComparisonUnitAtom],
    level: usize,
    part: Option<()>,
    settings: &WmlComparerSettings,
) {
    // Same as ReconstructElement but preserves VML property children
    // TODO: C# line 5740: GetVmlPropertyChildren
    // For now, we use the generic reconstruct_element
    reconstruct_element(doc, parent, group_key, ancestor, &[], group_atoms, level, part, settings);
}

/// Port of C# lines 4997-5010: MoveLastSectPrToChildOfBody
///
/// Moves the last w:sectPr from w:p/w:pPr/w:sectPr to w:body/w:sectPr
fn move_last_sect_pr_to_child_of_body(doc: &mut XmlDocument, doc_root: NodeId) {
    // Find w:body (C# line 5001)
    let body = doc.children(doc_root).find(|&child| {
        doc.get(child)
            .and_then(|d| d.name())
            .map(|n| n == &W::body())
            .unwrap_or(false)
    });
    
    if body.is_none() {
        return;
    }
    let body = body.unwrap();
    
    // Find last w:p with w:pPr/w:sectPr (C# lines 4999-5004)
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
    
    // Move sectPr to body (C# lines 5006-5009)
    if let (Some(_para), Some(sect_pr)) = (last_para_with_sect_pr, sect_pr_node) {
        doc.reparent(body, sect_pr);
    }
}

/// Helper: Create a content element from an atom
fn create_content_element(
    doc: &mut XmlDocument,
    atom: &ComparisonUnitAtom,
    status: &str,
) -> Option<NodeId> {
    match &atom.content_element {
        ContentElement::Text(ch) => {
            // Text is handled in reconstruct_text_elements
            None
        }
        ContentElement::Break => {
            let temp_root = doc.add_root(XmlNodeData::element(W::br()));
            let br = doc.add_child(temp_root, XmlNodeData::element(W::br()));
            if !status.is_empty() {
                doc.set_attribute(br, &pt_status(), status);
            }
            Some(br)
        }
        ContentElement::Tab => {
            let temp_root = doc.add_root(XmlNodeData::element(W::tab()));
            let tab = doc.add_child(temp_root, XmlNodeData::element(W::tab()));
            if !status.is_empty() {
                doc.set_attribute(tab, &pt_status(), status);
            }
            Some(tab)
        }
        _ => None,
    }
}

/// Helper: Group items by a key function, preserving order
fn group_by_key<T, F, K>(items: &[T], mut key_fn: F) -> Vec<(K, Vec<T>)>
where
    T: Clone,
    F: FnMut(&T) -> K,
    K: Eq + std::hash::Hash + Clone,
{
    let mut groups: HashMap<K, Vec<T>> = HashMap::new();
    let mut order: Vec<K> = Vec::new();
    
    for item in items {
        let key = key_fn(item);
        if !groups.contains_key(&key) {
            order.push(key.clone());
        }
        groups.entry(key).or_default().push(item.clone());
    }
    
    order
        .into_iter()
        .filter_map(|key| groups.remove(&key).map(|items| (key, items)))
        .collect()
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
