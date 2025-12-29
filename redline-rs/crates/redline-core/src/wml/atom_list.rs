use crate::package::OoxmlPackage;
use crate::wml::comparison_unit::{AncestorInfo, ComparisonUnitAtom, ContentElement, generate_unid};
use crate::wml::drawing_identity::compute_drawing_identity;
use crate::wml::formatting::{compute_normalized_rpr, compute_formatting_signature};
use crate::wml::settings::WmlComparerSettings;
use crate::xml::arena::XmlDocument;
use crate::xml::builder::serialize_subtree;
use crate::xml::namespaces::{M, MC, O, PT, V, W, W10};
use crate::xml::node::XmlNodeData;
use indextree::NodeId;
use sha1::{Digest, Sha1};
use std::collections::HashSet;

fn allowable_run_children() -> HashSet<String> {
    [
        "br", "drawing", "cr", "dayLong", "dayShort", "footnoteReference", "endnoteReference",
        "monthLong", "monthShort", "noBreakHyphen", "pgNum", "ptab", "softHyphen", "sym",
        "tab", "yearLong", "yearShort", "fldChar", "instrText",
    ]
    .iter()
    .map(|s| s.to_string())
    .collect()
}

fn allowable_run_children_math() -> HashSet<String> {
    ["oMathPara", "oMath"].iter().map(|s| s.to_string()).collect()
}

fn elements_to_throw_away() -> HashSet<String> {
    [
        "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd",
        "lastRenderedPageBreak", "proofErr", "tblPr", "sectPr", "permEnd", "permStart",
        "footnoteRef", "endnoteRef", "separator", "continuationSeparator",
    ]
    .iter()
    .map(|s| s.to_string())
    .collect()
}

struct RecursionInfo {
    element_name: &'static str,
    namespace: &'static str,
    child_props_to_skip: &'static [&'static str],
}

const RECURSION_ELEMENTS: &[RecursionInfo] = &[
    RecursionInfo { element_name: "del", namespace: W::NS, child_props_to_skip: &[] },
    RecursionInfo { element_name: "ins", namespace: W::NS, child_props_to_skip: &[] },
    RecursionInfo { element_name: "tbl", namespace: W::NS, child_props_to_skip: &["tblPr", "tblGrid", "tblPrEx"] },
    RecursionInfo { element_name: "tr", namespace: W::NS, child_props_to_skip: &["trPr", "tblPrEx"] },
    RecursionInfo { element_name: "tc", namespace: W::NS, child_props_to_skip: &["tcPr", "tblPrEx"] },
    RecursionInfo { element_name: "pict", namespace: W::NS, child_props_to_skip: &["shapetype"] },
    RecursionInfo { element_name: "group", namespace: V::NS, child_props_to_skip: &["fill", "stroke", "shadow", "path", "formulas", "handles", "lock", "extrusion"] },
    RecursionInfo { element_name: "shape", namespace: V::NS, child_props_to_skip: &["fill", "stroke", "shadow", "textpath", "path", "formulas", "handles", "imagedata", "lock", "extrusion", "wrap"] },
    RecursionInfo { element_name: "rect", namespace: V::NS, child_props_to_skip: &["fill", "stroke", "shadow", "textpath", "path", "formulas", "handles", "lock", "extrusion"] },
    RecursionInfo { element_name: "textbox", namespace: V::NS, child_props_to_skip: &[] },
    RecursionInfo { element_name: "lock", namespace: O::NS, child_props_to_skip: &[] },
    RecursionInfo { element_name: "txbxContent", namespace: W::NS, child_props_to_skip: &[] },
    RecursionInfo { element_name: "wrap", namespace: W10::NS, child_props_to_skip: &[] },
    RecursionInfo { element_name: "sdt", namespace: W::NS, child_props_to_skip: &["sdtPr", "sdtEndPr"] },
    RecursionInfo { element_name: "sdtContent", namespace: W::NS, child_props_to_skip: &[] },
    RecursionInfo { element_name: "hyperlink", namespace: W::NS, child_props_to_skip: &[] },
    RecursionInfo { element_name: "fldSimple", namespace: W::NS, child_props_to_skip: &[] },
    RecursionInfo { element_name: "shapetype", namespace: V::NS, child_props_to_skip: &["stroke", "path", "fill", "shadow", "formulas", "handles"] },
    RecursionInfo { element_name: "smartTag", namespace: W::NS, child_props_to_skip: &["smartTagPr"] },
    RecursionInfo { element_name: "ruby", namespace: W::NS, child_props_to_skip: &["rubyPr"] },
];

pub fn assign_unid_to_all_elements(doc: &mut XmlDocument, content_parent: NodeId) {
    let pt_unid = PT::Unid();
    let descendants: Vec<NodeId> = doc.descendants(content_parent).collect();
    
    for node_id in descendants {
        if let Some(data) = doc.get(node_id) {
            if data.is_element() {
                let has_unid = data
                    .attributes()
                    .map(|attrs| attrs.iter().any(|a| a.name == pt_unid))
                    .unwrap_or(false);
                
                if !has_unid {
                    let unid = generate_unid();
                    doc.set_attribute(node_id, &pt_unid, &unid);
                }
            }
        }
    }
}

/// Create a list of comparison unit atoms from a document.
///
/// This version does not resolve image relationships - drawings are hashed by XML structure only.
/// For proper image identity (SHA1 of binary content), use `create_comparison_unit_atom_list_with_package`.
pub fn create_comparison_unit_atom_list(
    doc: &mut XmlDocument,
    content_parent: NodeId,
    part_name: &str,
    settings: &WmlComparerSettings,
) -> Vec<ComparisonUnitAtom> {
    create_comparison_unit_atom_list_with_package(doc, content_parent, part_name, settings, None)
}

/// Create a list of comparison unit atoms from a document with package access.
///
/// This version resolves image relationships and computes SHA1 of binary image content
/// for stable drawing identity. This is the recommended version when package access is available.
///
/// # Arguments
/// * `doc` - The XML document
/// * `content_parent` - The parent node to start from
/// * `part_name` - The name of the part (e.g., "word/document.xml")
/// * `settings` - Comparer settings
/// * `package` - Optional OOXML package for resolving image relationships
pub fn create_comparison_unit_atom_list_with_package(
    doc: &mut XmlDocument,
    content_parent: NodeId,
    part_name: &str,
    settings: &WmlComparerSettings,
    package: Option<&OoxmlPackage>,
) -> Vec<ComparisonUnitAtom> {
    let mut atoms = Vec::new();
    let allowable = allowable_run_children();
    let allowable_math = allowable_run_children_math();
    let throw_away = elements_to_throw_away();
    
    create_atom_list_recurse(
        doc,
        content_parent,
        &mut atoms,
        part_name,
        &allowable,
        &allowable_math,
        &throw_away,
        settings,
        None,
        package,
    );
    
    atoms
}

fn create_atom_list_recurse(
    doc: &mut XmlDocument,
    node: NodeId,
    atoms: &mut Vec<ComparisonUnitAtom>,
    part_name: &str,
    allowable: &HashSet<String>,
    allowable_math: &HashSet<String>,
    throw_away: &HashSet<String>,
    settings: &WmlComparerSettings,
    current_formatting_signature: Option<String>,
    package: Option<&OoxmlPackage>,
) {
    let Some(data) = doc.get(node) else { return };
    let Some(name) = data.name() else { return };
    
    let ns = name.namespace.as_deref();
    let local = name.local_name.as_str();
    
    if ns == Some(W::NS) && (local == "body" || local == "footnote" || local == "endnote") {
        let children: Vec<_> = doc.children(node).collect();
        for child in children {
            create_atom_list_recurse(doc, child, atoms, part_name, allowable, allowable_math, throw_away, settings, None, package);
        }
        return;
    }
    
    if ns == Some(W::NS) && local == "p" {
        let children: Vec<_> = doc.children(node).collect();
        for child in children {
            if let Some(child_data) = doc.get(child) {
                if let Some(child_name) = child_data.name() {
                    if !(child_name.namespace.as_deref() == Some(W::NS) && child_name.local_name == "pPr") {
                        create_atom_list_recurse(doc, child, atoms, part_name, allowable, allowable_math, throw_away, settings, None, package);
                    }
                }
            }
        }
        
        let ancestors = build_ancestor_chain(doc, node);
        let mut atom = ComparisonUnitAtom::new(ContentElement::ParagraphProperties, ancestors, part_name, settings);
        atom.formatting_signature = None;
        atoms.push(atom);
        return;
    }
    
    if ns == Some(W::NS) && local == "r" {
        let normalized_rpr = compute_normalized_rpr(doc, node, settings);
        let formatting_signature = compute_formatting_signature(doc, normalized_rpr);
        
        let children: Vec<_> = doc.children(node).collect();
        for child in children {
            if let Some(child_data) = doc.get(child) {
                if let Some(child_name) = child_data.name() {
                    if !(child_name.namespace.as_deref() == Some(W::NS) && child_name.local_name == "rPr") {
                        create_atom_list_recurse(
                            doc, 
                            child, 
                            atoms, 
                            part_name, 
                            allowable, 
                            allowable_math, 
                            throw_away, 
                            settings, 
                            formatting_signature.clone(),
                            package,
                        );
                    }
                }
            }
        }
        return;
    }
    
    if ns == Some(W::NS) && (local == "t" || local == "delText") {
        let text = extract_text_content(doc, node);
        let ancestors = build_ancestor_chain(doc, node);
        
        for ch in text.chars() {
            let mut atom = ComparisonUnitAtom::new(
                ContentElement::Text(ch),
                ancestors.clone(),
                part_name,
                settings,
            );
            atom.formatting_signature = current_formatting_signature.clone();
            atoms.push(atom);
        }
        return;
    }
    
    if (ns == Some(W::NS) && allowable.contains(local)) || 
       (ns == Some(M::NS) && allowable_math.contains(local)) {
        let ancestors = build_ancestor_chain(doc, node);
        let content = create_content_element_with_package(doc, node, ns.unwrap_or(""), local, part_name, package);
        let mut atom = ComparisonUnitAtom::new(content, ancestors, part_name, settings);
        atom.formatting_signature = current_formatting_signature.clone();
        atoms.push(atom);
        return;
    }
    
    if ns == Some(W::NS) && local == "object" {
        let ancestors = build_ancestor_chain(doc, node);
        let hash = compute_element_hash(doc, node);
        let mut atom = ComparisonUnitAtom::new(
            ContentElement::Object { hash },
            ancestors,
            part_name,
            settings,
        );
        atom.formatting_signature = current_formatting_signature.clone();
        atoms.push(atom);
        return;
    }
    
    if let Some(re) = find_recursion_element(ns, local) {
        let skip_props: HashSet<&str> = re.child_props_to_skip.iter().copied().collect();
        
        let children: Vec<_> = doc.children(node).collect();
        for child in children {
            if let Some(child_data) = doc.get(child) {
                if let Some(child_name) = child_data.name() {
                    if !skip_props.contains(child_name.local_name.as_str()) {
                        create_atom_list_recurse(
                            doc, 
                            child, 
                            atoms, 
                            part_name, 
                            allowable, 
                            allowable_math, 
                            throw_away, 
                            settings, 
                            current_formatting_signature.clone(),
                            package,
                        );
                    }
                }
            }
        }
        return;
    }
    
    if ns == Some(W::NS) && throw_away.contains(local) {
        return;
    }
    
    // Handle mc:AlternateContent - only process one branch to avoid duplicates
    // Word documents often contain both DrawingML (mc:Choice) and VML (mc:Fallback) 
    // representations of the same content (like textboxes). We prefer Fallback for
    // compatibility, matching C# WmlComparer FlattenAlternateContent behavior.
    // See C# WmlComparer.cs:500-541
    if ns == Some(MC::NS) && local == "AlternateContent" {
        let children: Vec<_> = doc.children(node).collect();
        
        // First try mc:Fallback (VML representation - more compatible)
        let fallback_node = children.iter().find(|&&child| {
            doc.get(child)
                .and_then(|d| d.name())
                .map(|n| n.namespace.as_deref() == Some(MC::NS) && n.local_name == "Fallback")
                .unwrap_or(false)
        }).copied();
        
        if let Some(fallback) = fallback_node {
            let fallback_children: Vec<_> = doc.children(fallback).collect();
            for fallback_child in fallback_children {
                create_atom_list_recurse(
                    doc, fallback_child, atoms, part_name, 
                    allowable, allowable_math, throw_away, 
                    settings, current_formatting_signature.clone(),
                    package,
                );
            }
            return;
        }
        
        // No Fallback found, try mc:Choice
        let choice_node = children.iter().find(|&&child| {
            doc.get(child)
                .and_then(|d| d.name())
                .map(|n| n.namespace.as_deref() == Some(MC::NS) && n.local_name == "Choice")
                .unwrap_or(false)
        }).copied();
        
        if let Some(choice) = choice_node {
            let choice_children: Vec<_> = doc.children(choice).collect();
            for choice_child in choice_children {
                create_atom_list_recurse(
                    doc, choice_child, atoms, part_name, 
                    allowable, allowable_math, throw_away, 
                    settings, current_formatting_signature.clone(),
                    package,
                );
            }
            return;
        }
        
        // Neither found - skip this AlternateContent entirely
        return;
    }
    
    let children: Vec<_> = doc.children(node).collect();
    for child in children {
        create_atom_list_recurse(doc, child, atoms, part_name, allowable, allowable_math, throw_away, settings, current_formatting_signature.clone(), package);
    }
}

fn find_recursion_element(ns: Option<&str>, local: &str) -> Option<&'static RecursionInfo> {
    RECURSION_ELEMENTS.iter().find(|re| {
        Some(re.namespace) == ns && re.element_name == local
    })
}

fn build_ancestor_chain(doc: &XmlDocument, node: NodeId) -> Vec<AncestorInfo> {
    let pt_unid = PT::Unid();
    let mut ancestors = Vec::new();
    
    for ancestor_id in doc.ancestors(node) {
        if let Some(data) = doc.get(ancestor_id) {
            if let Some(name) = data.name() {
                let ns = name.namespace.as_deref();
                let local = name.local_name.as_str();
                
                if ns == Some(W::NS) && (local == "body" || local == "footnotes" || local == "endnotes") {
                    break;
                }
                
                // Filter out namespace declarations - they should only appear on the root element
                // during serialization, not on every descendant element
                let attrs = std::sync::Arc::new(
                    data.attributes()
                        .map(|a| a.iter()
                            .filter(|attr| {
                                // Keep attribute if it's NOT a namespace declaration
                                // Namespace declarations have namespace = "http://www.w3.org/2000/xmlns/"
                                // or are the "xmlns" attribute with no namespace
                                let is_xmlns_ns = attr.name.namespace.as_deref() == Some("http://www.w3.org/2000/xmlns/");
                                let is_xmlns_attr = attr.name.namespace.is_none() && attr.name.local_name == "xmlns";
                                !is_xmlns_ns && !is_xmlns_attr
                            })
                            .cloned()
                            .collect::<Vec<_>>())
                        .unwrap_or_default()
                );
                let unid = attrs.iter()
                    .find(|a| a.name == pt_unid)
                    .map(|a| a.value.clone())
                    .unwrap_or_default();
                
                // For run elements (w:r), capture the rPr child element for reconstruction
                // This is critical for preserving formatting in comparison output
                // C# equivalent: ancestorBeingConstructed.Element(W.rPr) in CoalesceRecurse
                let rpr_xml = if ns == Some(W::NS) && local == "r" {
                    extract_rpr_xml(doc, ancestor_id)
                } else {
                    None
                };
                
                ancestors.push(AncestorInfo {
                    node_id: ancestor_id,
                    namespace: name.namespace.clone(),
                    local_name: local.to_string(),
                    unid,
                    attributes: attrs,
                    has_merged_cells: false,  // Stub - not implemented yet
                    rpr_xml,
                });
            }
        }
    }
    
    ancestors.reverse();
    ancestors
}

/// Extract and serialize the w:rPr child element from a run element.
/// Returns None if no rPr is present.
fn extract_rpr_xml(doc: &XmlDocument, run_id: NodeId) -> Option<String> {
    // Find the w:rPr child element
    for child in doc.children(run_id) {
        if let Some(child_data) = doc.get(child) {
            if let Some(child_name) = child_data.name() {
                if child_name.namespace.as_deref() == Some(W::NS) && child_name.local_name == "rPr" {
                    // Serialize the rPr element and all its children
                    return Some(serialize_element_tree(doc, child));
                }
            }
        }
    }
    None
}

/// Serialize an XML element and its descendants to a string.
/// This produces a compact XML representation without unnecessary whitespace.
fn serialize_element_tree(doc: &XmlDocument, node: NodeId) -> String {
    let mut result = String::new();
    serialize_element_recursive(doc, node, &mut result);
    result
}

fn serialize_element_recursive(doc: &XmlDocument, node: NodeId, result: &mut String) {
    let Some(data) = doc.get(node) else { return };
    
    match data {
        XmlNodeData::Element { name, attributes } => {
            result.push('<');
            
            // Add namespace prefix
            if let Some(ns) = &name.namespace {
                if let Some(prefix) = get_prefix_for_ns(ns) {
                    result.push_str(prefix);
                    result.push(':');
                }
            }
            result.push_str(&name.local_name);
            
            // Add attributes (skip internal pt: namespace attributes)
            for attr in attributes {
                // Skip PowerTools internal tracking attributes
                if attr.name.namespace.as_deref() == Some("http://powertools.codeplex.com/2011") {
                    continue;
                }
                
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
                result.push_str("/>");
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

fn get_prefix_for_ns(ns: &str) -> Option<&'static str> {
    match ns {
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main" => Some("w"),
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships" => Some("r"),
        "http://schemas.openxmlformats.org/markup-compatibility/2006" => Some("mc"),
        "http://www.w3.org/XML/1998/namespace" => Some("xml"),
        _ => None,
    }
}

fn escape_xml_attr(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn escape_xml_text(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn extract_text_content(doc: &XmlDocument, node: NodeId) -> String {
    let mut text = String::new();
    for child in doc.children(node) {
        if let Some(XmlNodeData::Text(t)) = doc.get(child) {
            text.push_str(t);
        }
    }
    text
}

/// Create a content element with package access for proper drawing identity.
/// This version uses the drawing identity module to compute SHA1 of image content.
fn create_content_element_with_package(
    doc: &XmlDocument, 
    node: NodeId, 
    ns: &str, 
    local: &str,
    part_name: &str,
    package: Option<&OoxmlPackage>,
) -> ContentElement {
    match (ns, local) {
        (W::NS, "br") => ContentElement::Break,
        (W::NS, "tab") => ContentElement::Tab,
        (W::NS, "cr") => ContentElement::Break,
        (W::NS, "footnoteReference") => {
            let id = get_attribute_value(doc, node, W::NS, "id").unwrap_or_default();
            ContentElement::FootnoteReference { id }
        }
        (W::NS, "endnoteReference") => {
            let id = get_attribute_value(doc, node, W::NS, "id").unwrap_or_default();
            ContentElement::EndnoteReference { id }
        }
        (W::NS, "drawing") => {
            // Use the new drawing identity module for stable SHA1-based identity
            let hash = compute_drawing_identity(doc, node, package, part_name);
            // Serialize the full element content so it can be reconstructed
            let element_xml = serialize_subtree(doc, node).unwrap_or_default();
            ContentElement::Drawing { hash, element_xml }
        }
        (W::NS, "sym") => {
            let font = get_attribute_value(doc, node, W::NS, "font").unwrap_or_default();
            let char_code = get_attribute_value(doc, node, W::NS, "char").unwrap_or_default();
            ContentElement::Symbol { font, char_code }
        }
        (W::NS, "fldChar") => {
            let fld_char_type = get_attribute_value(doc, node, W::NS, "fldCharType").unwrap_or_default();
            match fld_char_type.as_str() {
                "begin" => ContentElement::FieldBegin,
                "separate" => ContentElement::FieldSeparator,
                "end" => ContentElement::FieldEnd,
                _ => ContentElement::Unknown { name: format!("fldChar:{}", fld_char_type) },
            }
        }
        (M::NS, "oMath") | (M::NS, "oMathPara") => {
            let hash = compute_element_hash(doc, node);
            let element_xml = serialize_subtree(doc, node).unwrap_or_default();
            ContentElement::Math { hash, element_xml }
        }
        _ => ContentElement::Unknown { name: local.to_string() },
    }
}

fn get_attribute_value(doc: &XmlDocument, node: NodeId, ns: &str, local: &str) -> Option<String> {
    let data = doc.get(node)?;
    let attrs = data.attributes()?;
    
    attrs.iter()
        .find(|a| a.name.local_name == local && a.name.namespace.as_deref() == Some(ns))
        .or_else(|| attrs.iter().find(|a| a.name.local_name == local && a.name.namespace.is_none()))
        .map(|a| a.value.clone())
}

fn compute_element_hash(doc: &XmlDocument, node: NodeId) -> String {
    let mut hasher = Sha1::new();
    hash_element_recursive(doc, node, &mut hasher);
    let result = hasher.finalize();
    format!("{:x}", result)
}

fn hash_element_recursive(doc: &XmlDocument, node: NodeId, hasher: &mut Sha1) {
    let Some(data) = doc.get(node) else { return };
    
    match data {
        XmlNodeData::Element { name, attributes } => {
            hasher.update(name.local_name.as_bytes());
            let pt_unid = PT::Unid();
            let pt_sha1 = PT::SHA1Hash();
            for attr in attributes {
                if attr.name != pt_unid && attr.name != pt_sha1 {
                    hasher.update(attr.name.local_name.as_bytes());
                    hasher.update(attr.value.as_bytes());
                }
            }
            for child in doc.children(node) {
                hash_element_recursive(doc, child, hasher);
            }
        }
        XmlNodeData::Text(text) => {
            hasher.update(text.as_bytes());
        }
        _ => {}
    }
}
