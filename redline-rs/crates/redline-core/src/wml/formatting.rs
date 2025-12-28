//! Formatting change detection for WmlComparer
//!
//! This module contains logic for detecting and tracking formatting changes
//! during document comparison. It is a faithful port of the formatting-related
//! code from C# OpenXmlPowerTools WmlComparer.cs (lines 8298-8466).
//!
//! Key components:
//! - `compute_normalized_rpr`: Normalizes run properties for comparison
//! - `compute_formatting_signature`: Generates SHA1 hash of normalized formatting
//! - Allowed formatting properties filter
//! - Attribute cleanup for consistent comparison

use crate::hash::sha1::sha1_hash_string;
use crate::wml::settings::WmlComparerSettings;
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{W, PT};
use crate::xml::node::XmlNodeData;
use indextree::NodeId;
use std::collections::HashSet;

/// Formatting properties considered for comparison (C# WmlComparer.cs:8314-8336)
static ALLOWED_FORMATTING_PROPERTIES: &[&str] = &[
    "b", "bCs", "i", "iCs", "u",
    "sz", "szCs", "color", "rFonts",
    "highlight", "strike", "dstrike",
    "caps", "smallCaps",
];

/// Properties where attribute values are semantically significant (C# WmlComparer.cs:8338-8345)
static PROPS_WITH_VALUES: &[&str] = &[
    "u", "color", "sz", "szCs", "rFonts", "highlight",
];

/// Font attribute names preserved for rFonts elements
static FONT_ATTRIBUTES: &[&str] = &["ascii", "hAnsi", "cs", "eastAsia"];

fn clone_element_deep_in_place(doc: &mut XmlDocument, source: NodeId, parent: Option<NodeId>) -> NodeId {
    let source_data = doc.get(source).expect("Source node must exist").clone();
    
    let cloned = match parent {
        Some(p) => doc.add_child(p, source_data),
        None => {
            let id = doc.new_node(XmlNodeData::element(W::rPr()));
            if let Some(mut_data) = doc.get_mut(id) {
                *mut_data = source_data;
            }
            id
        }
    };
    
    let children: Vec<_> = doc.children(source).collect();
    for child in children {
        clone_element_deep_in_place(doc, child, Some(cloned));
    }
    
    cloned
}

/// Compute normalized run properties for formatting comparison.
/// This is a faithful port of C# ComputeNormalizedRPr (WmlComparer.cs lines 8380-8466).
///
/// # Purpose
/// Normalizes run properties (rPr) by:
/// 1. Filtering to only allowed formatting properties
/// 2. Removing internal/tracking attributes (pt:*, rsid*, etc.)
/// 3. Keeping only meaningful attributes for value-based properties
///
/// # Arguments
/// * `doc` - The XML document containing the elements
/// * `run_node` - The w:r (run) element node
/// * `settings` - Comparer settings (checks TrackFormattingChanges flag)
///
/// # Returns
/// * `Some(NodeId)` - Normalized rPr element if formatting tracking is enabled
/// * `None` - If tracking is disabled or no relevant formatting found
pub fn compute_normalized_rpr(
    doc: &mut XmlDocument,
    run_node: NodeId,
    settings: &WmlComparerSettings,
) -> Option<NodeId> {
    if !settings.track_formatting_changes {
        return None;
    }

    let rpr_node = doc.children(run_node)
        .find(|&child| {
            doc.get(child)
                .and_then(|data| data.name())
                .map(|name| name == &W::rPr())
                .unwrap_or(false)
        })?;

    let clone_node = clone_element_deep_in_place(doc, rpr_node, None);

    let allowed_set: HashSet<&str> = ALLOWED_FORMATTING_PROPERTIES.iter().copied().collect();
    let props_with_values_set: HashSet<&str> = PROPS_WITH_VALUES.iter().copied().collect();
    let font_attrs_set: HashSet<&str> = FONT_ATTRIBUTES.iter().copied().collect();
    
    let elements_to_remove: Vec<NodeId> = doc.children(clone_node)
        .filter(|&child| {
            doc.get(child)
                .and_then(|data| data.name())
                .map(|name| !allowed_set.contains(name.local_name.as_str()))
                .unwrap_or(false)
        })
        .collect();

    for elem in elements_to_remove {
        doc.remove(elem);
    }

    let mut all_descendants: Vec<NodeId> = vec![clone_node];
    all_descendants.extend(doc.descendants(clone_node));
    
    for node_id in all_descendants {
        if let Some(data) = doc.get(node_id) {
            if let XmlNodeData::Element { name, .. } = data {
                let element_name = name.local_name.clone();
                
                if let Some(data_mut) = doc.get_mut(node_id) {
                    if let XmlNodeData::Element { attributes, .. } = data_mut {
                        let mut attrs_to_remove = Vec::new();
                        
                        for attr in attributes.iter() {
                            let is_pt_namespace = attr.name.namespace.as_deref() == Some(PT::NS);
                            let is_xmlns = attr.name.namespace.as_deref() == Some("http://www.w3.org/2000/xmlns/");
                            
                            if is_pt_namespace
                                || is_xmlns
                                || attr.name.local_name == "Unid"
                                || attr.name.local_name.to_lowercase().starts_with("rsid") 
                            {
                                attrs_to_remove.push(attr.name.clone());
                            } else if props_with_values_set.contains(element_name.as_str()) {
                                let should_keep = attr.name.local_name == "val" 
                                    || font_attrs_set.contains(attr.name.local_name.as_str());
                                
                                if !should_keep {
                                    attrs_to_remove.push(attr.name.clone());
                                }
                            }
                        }
                        
                        for attr_name in attrs_to_remove {
                            doc.remove_attribute(node_id, &attr_name);
                        }
                    }
                }
            }
        }
    }

    let has_content = doc.children(clone_node).next().is_some() 
        || doc.get(clone_node)
            .and_then(|data| data.attributes())
            .map(|attrs| !attrs.is_empty())
            .unwrap_or(false);

    if has_content {
        Some(clone_node)
    } else {
        None
    }
}

/// Compute formatting signature from normalized rPr.
/// This is a faithful port of C# FormattingSignature computation (WmlComparer.cs lines 8376-8377).
///
/// # Purpose
/// Creates a string representation of normalized formatting for comparison.
/// Uses SHA1 hashing for efficient comparison.
///
/// # Arguments
/// * `doc` - The XML document
/// * `normalized_rpr` - The normalized rPr node (from compute_normalized_rpr)
///
/// # Returns
/// * String representation suitable for comparison (serialized XML)
pub fn compute_formatting_signature(
    doc: &XmlDocument,
    normalized_rpr: Option<NodeId>,
) -> Option<String> {
    normalized_rpr.map(|node| serialize_node_no_formatting(doc, node))
}

fn serialize_node_no_formatting(doc: &XmlDocument, node: NodeId) -> String {
    let mut result = String::new();
    serialize_node_recursive(doc, node, &mut result);
    result
}

fn serialize_node_recursive(doc: &XmlDocument, node: NodeId, result: &mut String) {
    if let Some(data) = doc.get(node) {
        match data {
            XmlNodeData::Element { name, attributes } => {
                result.push('<');
                if name.namespace.is_some() {
                    result.push_str("w:");
                }
                result.push_str(&name.local_name);
                
                for attr in attributes {
                    result.push(' ');
                    if attr.name.namespace.is_some() {
                        result.push_str("w:");
                    }
                    result.push_str(&attr.name.local_name);
                    result.push_str("=\"");
                    result.push_str(&attr.value);
                    result.push('"');
                }
                
                let has_children = doc.children(node).next().is_some();
                if has_children {
                    result.push('>');
                    for child in doc.children(node) {
                        serialize_node_recursive(doc, child, result);
                    }
                    result.push_str("</");
                    if name.namespace.is_some() {
                        result.push_str("w:");
                    }
                    result.push_str(&name.local_name);
                    result.push('>');
                } else {
                    result.push_str(" />");
                }
            }
            XmlNodeData::Text(text) => {
                result.push_str(text);
            }
            XmlNodeData::CData(text) => {
                result.push_str("<![CDATA[");
                result.push_str(text);
                result.push_str("]]>");
            }
            XmlNodeData::Comment(text) => {
                result.push_str("<!--");
                result.push_str(text);
                result.push_str("-->");
            }
            XmlNodeData::ProcessingInstruction { target, data } => {
                result.push_str("<?");
                result.push_str(target);
                result.push(' ');
                result.push_str(data);
                result.push_str("?>");
            }
        }
    }
}

/// Compute formatting signature with SHA1 hash.
/// Alternative version that returns a hash instead of the full XML string.
///
/// # Arguments
/// * `doc` - The XML document
/// * `normalized_rpr` - The normalized rPr node
///
/// # Returns
/// * SHA1 hash of the formatting signature
pub fn compute_formatting_signature_hash(
    doc: &XmlDocument,
    normalized_rpr: Option<NodeId>,
) -> Option<String> {
    compute_formatting_signature(doc, normalized_rpr)
        .map(|sig| sha1_hash_string(&sig))
}

/// Check if two formatting signatures differ.
/// Helper function for determining if a formatting change occurred.
///
/// # Arguments
/// * `before` - Formatting signature from "before" document
/// * `after` - Formatting signature from "after" document
///
/// # Returns
/// * `true` if formatting differs, `false` if same or both None
pub fn formatting_differs(before: &Option<String>, after: &Option<String>) -> bool {
    match (before, after) {
        (None, None) => false,
        (None, Some(_)) => true,
        (Some(_), None) => true,
        (Some(b), Some(a)) => b != a,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::arena::XmlDocument;
    use crate::xml::node::XmlNodeData;
    use crate::xml::namespaces::W;

    #[test]
    fn test_allowed_properties_list() {
        assert_eq!(ALLOWED_FORMATTING_PROPERTIES.len(), 14);
        assert!(ALLOWED_FORMATTING_PROPERTIES.contains(&"b"));
        assert!(ALLOWED_FORMATTING_PROPERTIES.contains(&"i"));
        assert!(ALLOWED_FORMATTING_PROPERTIES.contains(&"u"));
        assert!(ALLOWED_FORMATTING_PROPERTIES.contains(&"sz"));
        assert!(ALLOWED_FORMATTING_PROPERTIES.contains(&"color"));
    }

    #[test]
    fn test_props_with_values_list() {
        assert_eq!(PROPS_WITH_VALUES.len(), 6);
        assert!(PROPS_WITH_VALUES.contains(&"u"));
        assert!(PROPS_WITH_VALUES.contains(&"color"));
        assert!(PROPS_WITH_VALUES.contains(&"sz"));
        assert!(PROPS_WITH_VALUES.contains(&"rFonts"));
    }

    #[test]
    fn test_compute_normalized_rpr_disabled() {
        let mut doc = XmlDocument::new();
        let root = doc.add_root(XmlNodeData::element(W::r()));
        
        let settings = WmlComparerSettings::default().with_track_formatting(false);
        
        let result = compute_normalized_rpr(&mut doc, root, &settings);
        assert!(result.is_none());
    }

    #[test]
    fn test_formatting_differs() {
        assert!(!formatting_differs(&None, &None));
        assert!(formatting_differs(&None, &Some("test".to_string())));
        assert!(formatting_differs(&Some("test".to_string()), &None));
        assert!(!formatting_differs(&Some("same".to_string()), &Some("same".to_string())));
        assert!(formatting_differs(&Some("before".to_string()), &Some("after".to_string())));
    }

    #[test]
    fn test_compute_formatting_signature() {
        let mut doc = XmlDocument::new();
        let rpr = doc.add_root(XmlNodeData::element(W::rPr()));
        
        let sig = compute_formatting_signature(&doc, Some(rpr));
        assert!(sig.is_some());
        assert!(sig.unwrap().contains("rPr"));
    }
}
