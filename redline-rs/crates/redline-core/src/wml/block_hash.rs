//! Block-level content hashing for document correlation
//!
//! Faithful port of C# WmlComparer.HashBlockLevelContent and CloneBlockLevelContentForHashing.
//! These functions compute CorrelatedSHA1Hash for paragraphs, tables, and rows to enable
//! efficient block-level matching during document comparison.

use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{PT, W};
use crate::xml::node::XmlNodeData;
use crate::xml::xname::XName;
use indextree::NodeId;
use sha1::{Digest, Sha1};
use std::collections::HashMap;

fn is_rsid_attribute(name: &XName) -> bool {
    if let Some(ns) = &name.namespace {
        if ns == W::NS {
            let local = &name.local_name;
            return local == "rsid"
                || local == "rsidDel"
                || local == "rsidP"
                || local == "rsidR"
                || local == "rsidRDefault"
                || local == "rsidRPr"
                || local == "rsidSect"
                || local == "rsidTr";
        }
    }
    false
}

fn is_pt_namespace(name: &XName) -> bool {
    name.namespace.as_deref() == Some(PT::NS)
}

fn should_skip_element(name: &XName) -> bool {
    let ns = name.namespace.as_deref();
    let local = &name.local_name;

    if ns == Some(W::NS) {
        return local == "bookmarkStart"
            || local == "bookmarkEnd"
            || local == "pPr"
            || local == "rPr";
    }

    false
}

fn clone_element_for_hashing(
    doc: &XmlDocument,
    node: NodeId,
    output: &mut String,
    settings: &HashingSettings,
) {
    let Some(data) = doc.get(node) else { return };

    match data {
        XmlNodeData::Element { name, attributes } => {
            if should_skip_element(name) {
                return;
            }

            let local = &name.local_name;
            let ns = name.namespace.as_deref();

            if ns == Some(W::NS) && local == "p" {
                clone_paragraph_for_hashing(doc, node, output, settings);
                return;
            }

            if ns == Some(W::NS) && local == "r" {
                clone_run_for_hashing(doc, node, output, settings);
                return;
            }

            if ns == Some(W::NS) && local == "tbl" {
                output.push_str("<w:tbl>");
                for child in doc.children(node) {
                    if let Some(child_data) = doc.get(child) {
                        if let Some(child_name) = child_data.name() {
                            if child_name.local_name == "tr"
                                && child_name.namespace.as_deref() == Some(W::NS)
                            {
                                clone_element_for_hashing(doc, child, output, settings);
                            }
                        }
                    }
                }
                output.push_str("</w:tbl>");
                return;
            }

            if ns == Some(W::NS) && local == "tr" {
                output.push_str("<w:tr>");
                for child in doc.children(node) {
                    if let Some(child_data) = doc.get(child) {
                        if let Some(child_name) = child_data.name() {
                            if child_name.local_name == "tc"
                                && child_name.namespace.as_deref() == Some(W::NS)
                            {
                                clone_element_for_hashing(doc, child, output, settings);
                            }
                        }
                    }
                }
                output.push_str("</w:tr>");
                return;
            }

            if ns == Some(W::NS) && local == "tc" {
                output.push_str("<w:tc>");
                for child in doc.children(node) {
                    clone_element_for_hashing(doc, child, output, settings);
                }
                output.push_str("</w:tc>");
                return;
            }

            if ns == Some(W::NS) && local == "tcPr" {
                output.push_str("<w:tcPr>");
                for child in doc.children(node) {
                    if let Some(child_data) = doc.get(child) {
                        if let Some(child_name) = child_data.name() {
                            if child_name.local_name == "gridSpan" {
                                clone_element_for_hashing(doc, child, output, settings);
                            }
                        }
                    }
                }
                output.push_str("</w:tcPr>");
                return;
            }

            if ns == Some(W::NS) && local == "gridSpan" {
                let val = attributes
                    .iter()
                    .find(|a| a.name.local_name == "val")
                    .map(|a| a.value.as_str())
                    .unwrap_or("");
                output.push_str(&format!("<w:gridSpan val=\"{}\"/>", val));
                return;
            }

            if ns == Some(W::NS) && (local == "pict" || local == "drawing") {
                let has_textbox = has_descendant_txbx_content(doc, node);
                if has_textbox {
                    output.push_str(&format!("<w:{}>", local));
                    for txbx_node in find_descendant_txbx_content(doc, node) {
                        clone_element_for_hashing(doc, txbx_node, output, settings);
                    }
                    output.push_str(&format!("</w:{}>", local));
                    return;
                }
            }

            if ns == Some(W::NS) && local == "txbxContent" {
                output.push_str("<w:txbxContent>");
                for child in doc.children(node) {
                    clone_element_for_hashing(doc, child, output, settings);
                }
                output.push_str("</w:txbxContent>");
                return;
            }

            output.push('<');
            if let Some(ns_prefix) = get_prefix_for_namespace(ns) {
                output.push_str(ns_prefix);
                output.push(':');
            }
            output.push_str(local);

            for attr in attributes {
                if is_rsid_attribute(&attr.name) || is_pt_namespace(&attr.name) {
                    continue;
                }
                output.push(' ');
                output.push_str(&attr.name.local_name);
                output.push_str("=\"");
                output.push_str(&escape_xml_attr(&attr.value));
                output.push('"');
            }

            let children: Vec<_> = doc.children(node).collect();
            if children.is_empty() {
                output.push_str("/>");
            } else {
                output.push('>');
                for child in children {
                    clone_element_for_hashing(doc, child, output, settings);
                }
                output.push_str("</");
                if let Some(ns_prefix) = get_prefix_for_namespace(ns) {
                    output.push_str(ns_prefix);
                    output.push(':');
                }
                output.push_str(local);
                output.push('>');
            }
        }
        XmlNodeData::Text(text) => {
            let mut normalized = text.clone();
            if settings.case_insensitive {
                normalized = normalized.to_uppercase();
            }
            if settings.conflate_spaces {
                normalized = normalized.replace(' ', "\u{00a0}");
            }
            output.push_str(&escape_xml_text(&normalized));
        }
        _ => {}
    }
}

fn clone_paragraph_for_hashing(
    doc: &XmlDocument,
    node: NodeId,
    output: &mut String,
    settings: &HashingSettings,
) {
    output.push_str("<w:p>");

    let mut text_buffer = String::new();
    let mut has_content = false;

    for child in doc.children(node) {
        if let Some(child_data) = doc.get(child) {
            if let Some(name) = child_data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "pPr" {
                    continue;
                }
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "r" {
                    if is_text_only_run(doc, child) {
                        let run_text = extract_run_text(doc, child);
                        text_buffer.push_str(&run_text);
                    } else {
                        flush_text_buffer(&mut text_buffer, output, settings);
                        clone_element_for_hashing(doc, child, output, settings);
                    }
                    has_content = true;
                } else {
                    flush_text_buffer(&mut text_buffer, output, settings);
                    clone_element_for_hashing(doc, child, output, settings);
                    has_content = true;
                }
            }
        }
    }

    flush_text_buffer(&mut text_buffer, output, settings);
    let _ = has_content;

    output.push_str("</w:p>");
}

fn clone_run_for_hashing(
    doc: &XmlDocument,
    node: NodeId,
    output: &mut String,
    settings: &HashingSettings,
) {
    if settings.track_formatting_changes {
        output.push_str("<w:r>");
        for child in doc.children(node) {
            if let Some(child_data) = doc.get(child) {
                if let Some(name) = child_data.name() {
                    if name.namespace.as_deref() == Some(W::NS) && name.local_name == "rPr" {
                        continue;
                    }
                }
            }
            clone_element_for_hashing(doc, child, output, settings);
        }
        output.push_str("</w:r>");
    } else {
        for child in doc.children(node) {
            if let Some(child_data) = doc.get(child) {
                if let Some(name) = child_data.name() {
                    if name.namespace.as_deref() == Some(W::NS) && name.local_name == "rPr" {
                        continue;
                    }
                }
            }
            clone_element_for_hashing(doc, child, output, settings);
        }
    }
}

fn is_text_only_run(doc: &XmlDocument, run_node: NodeId) -> bool {
    let mut has_t = false;
    let mut has_other = false;

    for child in doc.children(run_node) {
        if let Some(child_data) = doc.get(child) {
            if let Some(name) = child_data.name() {
                if name.namespace.as_deref() == Some(W::NS) {
                    match name.local_name.as_str() {
                        "t" => has_t = true,
                        "rPr" => {}
                        _ => has_other = true,
                    }
                }
            }
        }
    }

    has_t && !has_other
}

fn extract_run_text(doc: &XmlDocument, run_node: NodeId) -> String {
    let mut text = String::new();
    for child in doc.children(run_node) {
        if let Some(child_data) = doc.get(child) {
            if let Some(name) = child_data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "t" {
                    for text_child in doc.children(child) {
                        if let Some(XmlNodeData::Text(t)) = doc.get(text_child) {
                            text.push_str(t);
                        }
                    }
                }
            }
        }
    }
    text
}

fn flush_text_buffer(buffer: &mut String, output: &mut String, settings: &HashingSettings) {
    if !buffer.is_empty() {
        let mut text = std::mem::take(buffer);
        if settings.case_insensitive {
            text = text.to_uppercase();
        }
        if settings.conflate_spaces {
            text = text.replace(' ', "\u{00a0}");
        }
        output.push_str("<w:r><w:t>");
        output.push_str(&escape_xml_text(&text));
        output.push_str("</w:t></w:r>");
    }
}

fn has_descendant_txbx_content(doc: &XmlDocument, node: NodeId) -> bool {
    for desc in doc.descendants(node) {
        if let Some(data) = doc.get(desc) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "txbxContent" {
                    return true;
                }
            }
        }
    }
    false
}

fn find_descendant_txbx_content(doc: &XmlDocument, node: NodeId) -> Vec<NodeId> {
    let mut result = Vec::new();
    for desc in doc.descendants(node) {
        if let Some(data) = doc.get(desc) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "txbxContent" {
                    result.push(desc);
                }
            }
        }
    }
    result
}

fn get_prefix_for_namespace(ns: Option<&str>) -> Option<&'static str> {
    match ns {
        Some(W::NS) => Some("w"),
        Some("urn:schemas-microsoft-com:vml") => Some("v"),
        Some("urn:schemas-microsoft-com:office:office") => Some("o"),
        Some("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing") => {
            Some("wp")
        }
        Some("http://schemas.openxmlformats.org/drawingml/2006/main") => Some("a"),
        _ => None,
    }
}

fn escape_xml_text(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => result.push_str("&amp;"),
            '<' => result.push_str("&lt;"),
            '>' => result.push_str("&gt;"),
            _ => result.push(c),
        }
    }
    result
}

fn escape_xml_attr(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => result.push_str("&amp;"),
            '<' => result.push_str("&lt;"),
            '>' => result.push_str("&gt;"),
            '"' => result.push_str("&quot;"),
            _ => result.push(c),
        }
    }
    result
}

#[derive(Default)]
pub struct HashingSettings {
    pub case_insensitive: bool,
    pub conflate_spaces: bool,
    pub track_formatting_changes: bool,
}

pub fn clone_block_level_content_for_hashing(
    doc: &XmlDocument,
    node: NodeId,
    settings: &HashingSettings,
) -> String {
    let mut output = String::new();
    clone_element_for_hashing(doc, node, &mut output, settings);
    output
}

pub fn compute_block_hash(doc: &XmlDocument, node: NodeId, settings: &HashingSettings) -> String {
    let xml_string = clone_block_level_content_for_hashing(doc, node, settings);
    let mut hasher = Sha1::new();
    hasher.update(xml_string.as_bytes());
    let result = hasher.finalize();
    format!("{:x}", result)
}

pub fn hash_block_level_content(
    source_doc: &mut XmlDocument,
    source_root: NodeId,
    after_proc_doc: &XmlDocument,
    after_proc_root: NodeId,
    settings: &HashingSettings,
) {
    let pt_unid = PT::Unid();
    let pt_correlated_hash = PT::CorrelatedSHA1Hash();

    let mut source_unid_map: HashMap<String, NodeId> = HashMap::new();
    for node in source_doc.descendants(source_root) {
        if let Some(data) = source_doc.get(node) {
            if let Some(name) = data.name() {
                let ns = name.namespace.as_deref();
                let local = &name.local_name;
                if ns == Some(W::NS) && (local == "p" || local == "tbl" || local == "tr" || local == "txbxContent") {
                    if let Some(attrs) = data.attributes() {
                        if let Some(unid_attr) = attrs.iter().find(|a| a.name == pt_unid) {
                            source_unid_map.insert(unid_attr.value.clone(), node);
                        }
                    }
                }
            }
        }
    }

    for node in after_proc_doc.descendants(after_proc_root) {
        if let Some(data) = after_proc_doc.get(node) {
            if let Some(name) = data.name() {
                let ns = name.namespace.as_deref();
                let local = &name.local_name;
                if ns == Some(W::NS) && (local == "p" || local == "tbl" || local == "tr" || local == "txbxContent") {
                    let sha1_hash = compute_block_hash(after_proc_doc, node, settings);

                    if let Some(attrs) = data.attributes() {
                        if let Some(unid_attr) = attrs.iter().find(|a| a.name == pt_unid) {
                            if let Some(&source_node) = source_unid_map.get(&unid_attr.value) {
                                source_doc.set_attribute(
                                    source_node,
                                    &pt_correlated_hash,
                                    &sha1_hash,
                                );
                            }
                        }
                    }
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_escape_xml_text() {
        assert_eq!(escape_xml_text("hello"), "hello");
        assert_eq!(escape_xml_text("a & b"), "a &amp; b");
        assert_eq!(escape_xml_text("<tag>"), "&lt;tag&gt;");
    }

    #[test]
    fn test_escape_xml_attr() {
        assert_eq!(escape_xml_attr("hello"), "hello");
        assert_eq!(escape_xml_attr("a\"b"), "a&quot;b");
    }

    #[test]
    fn test_is_rsid_attribute() {
        assert!(is_rsid_attribute(&W::rsid()));
        assert!(is_rsid_attribute(&W::rsidR()));
        assert!(!is_rsid_attribute(&W::id()));
    }
}
