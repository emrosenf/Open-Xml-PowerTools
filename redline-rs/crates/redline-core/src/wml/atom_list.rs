use crate::wml::comparison_unit::{AncestorInfo, ComparisonUnitAtom, ContentElement, generate_unid};
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{M, O, PT, V, W, W10};
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

pub fn create_comparison_unit_atom_list(
    doc: &XmlDocument,
    content_parent: NodeId,
    part_name: &str,
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
    );
    
    atoms
}

fn create_atom_list_recurse(
    doc: &XmlDocument,
    node: NodeId,
    atoms: &mut Vec<ComparisonUnitAtom>,
    part_name: &str,
    allowable: &HashSet<String>,
    allowable_math: &HashSet<String>,
    throw_away: &HashSet<String>,
) {
    let Some(data) = doc.get(node) else { return };
    let Some(name) = data.name() else { return };
    
    let ns = name.namespace.as_deref();
    let local = name.local_name.as_str();
    
    if ns == Some(W::NS) && (local == "body" || local == "footnote" || local == "endnote") {
        for child in doc.children(node) {
            create_atom_list_recurse(doc, child, atoms, part_name, allowable, allowable_math, throw_away);
        }
        return;
    }
    
    if ns == Some(W::NS) && local == "p" {
        for child in doc.children(node) {
            if let Some(child_data) = doc.get(child) {
                if let Some(child_name) = child_data.name() {
                    if !(child_name.namespace.as_deref() == Some(W::NS) && child_name.local_name == "pPr") {
                        create_atom_list_recurse(doc, child, atoms, part_name, allowable, allowable_math, throw_away);
                    }
                }
            }
        }
        
        let ancestors = build_ancestor_chain(doc, node);
        let atom = ComparisonUnitAtom::new(ContentElement::ParagraphProperties, ancestors, part_name);
        atoms.push(atom);
        return;
    }
    
    if ns == Some(W::NS) && local == "r" {
        for child in doc.children(node) {
            if let Some(child_data) = doc.get(child) {
                if let Some(child_name) = child_data.name() {
                    if !(child_name.namespace.as_deref() == Some(W::NS) && child_name.local_name == "rPr") {
                        create_atom_list_recurse(doc, child, atoms, part_name, allowable, allowable_math, throw_away);
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
            let atom = ComparisonUnitAtom::new(
                ContentElement::Text(ch),
                ancestors.clone(),
                part_name,
            );
            atoms.push(atom);
        }
        return;
    }
    
    if (ns == Some(W::NS) && allowable.contains(local)) || 
       (ns == Some(M::NS) && allowable_math.contains(local)) {
        let ancestors = build_ancestor_chain(doc, node);
        let content = create_content_element(doc, node, ns.unwrap_or(""), local);
        let atom = ComparisonUnitAtom::new(content, ancestors, part_name);
        atoms.push(atom);
        return;
    }
    
    if ns == Some(W::NS) && local == "object" {
        let ancestors = build_ancestor_chain(doc, node);
        let hash = compute_element_hash(doc, node);
        let atom = ComparisonUnitAtom::new(
            ContentElement::Object { hash },
            ancestors,
            part_name,
        );
        atoms.push(atom);
        return;
    }
    
    if let Some(re) = find_recursion_element(ns, local) {
        let skip_props: HashSet<&str> = re.child_props_to_skip.iter().copied().collect();
        
        for child in doc.children(node) {
            if let Some(child_data) = doc.get(child) {
                if let Some(child_name) = child_data.name() {
                    if !skip_props.contains(child_name.local_name.as_str()) {
                        create_atom_list_recurse(doc, child, atoms, part_name, allowable, allowable_math, throw_away);
                    }
                }
            }
        }
        return;
    }
    
    if ns == Some(W::NS) && throw_away.contains(local) {
        return;
    }
    
    for child in doc.children(node) {
        create_atom_list_recurse(doc, child, atoms, part_name, allowable, allowable_math, throw_away);
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
                
                let unid = data
                    .attributes()
                    .and_then(|attrs| attrs.iter().find(|a| a.name == pt_unid))
                    .map(|a| a.value.clone())
                    .unwrap_or_default();
                
                ancestors.push(AncestorInfo {
                    node_id: ancestor_id,
                    local_name: local.to_string(),
                    unid,
                });
            }
        }
    }
    
    ancestors.reverse();
    ancestors
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

fn create_content_element(doc: &XmlDocument, node: NodeId, ns: &str, local: &str) -> ContentElement {
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
            let hash = compute_element_hash(doc, node);
            ContentElement::Drawing { hash }
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
            ContentElement::Math { hash }
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

#[cfg(test)]
mod tests {
    use super::*;
    
    #[test]
    fn test_allowable_run_children() {
        let allowable = allowable_run_children();
        assert!(allowable.contains("br"));
        assert!(allowable.contains("drawing"));
        assert!(allowable.contains("tab"));
        assert!(!allowable.contains("t"));
    }
    
    #[test]
    fn test_elements_to_throw_away() {
        let throw_away = elements_to_throw_away();
        assert!(throw_away.contains("bookmarkStart"));
        assert!(throw_away.contains("proofErr"));
        assert!(!throw_away.contains("p"));
    }
}
