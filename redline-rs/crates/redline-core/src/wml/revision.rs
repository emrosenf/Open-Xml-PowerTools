use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{W, W16DU};
use crate::xml::node::XmlNodeData;
use crate::xml::xname::{XAttribute, XName};
use indextree::NodeId;
use std::sync::atomic::{AtomicI32, Ordering};

static REVISION_ID_COUNTER: AtomicI32 = AtomicI32::new(0);

#[derive(Debug, Clone)]
pub struct RevisionSettings {
    pub author: String,
    pub date_time: String,
}

impl Default for RevisionSettings {
    fn default() -> Self {
        Self {
            author: "redline-rs".to_string(),
            date_time: chrono::Utc::now().format("%Y-%m-%dT%H:%M:%SZ").to_string(),
        }
    }
}

pub fn reset_revision_id_counter(value: i32) {
    REVISION_ID_COUNTER.store(value, Ordering::SeqCst);
}

pub fn get_next_revision_id() -> i32 {
    REVISION_ID_COUNTER.fetch_add(1, Ordering::SeqCst)
}

pub fn find_max_revision_id(doc: &XmlDocument, start: NodeId) -> i32 {
    let mut max_id = 0;
    
    for node_id in doc.descendants(start) {
        if let Some(data) = doc.get(node_id) {
            if let Some(attrs) = data.attributes() {
                for attr in attrs {
                    if attr.name.local_name == "id" && attr.name.namespace.as_deref() == Some(W::NS) {
                        if let Ok(id) = attr.value.parse::<i32>() {
                            max_id = max_id.max(id);
                        }
                    }
                }
            }
        }
    }
    
    max_id
}

pub fn create_insertion(
    doc: &mut XmlDocument,
    parent: NodeId,
    settings: &RevisionSettings,
) -> NodeId {
    let rev_id = get_next_revision_id();
    let attrs = vec![
        XAttribute::new(W::id(), &rev_id.to_string()),  // w:id MUST come first per ECMA-376
        XAttribute::new(W::author(), &settings.author),
        XAttribute::new(W::date(), &settings.date_time),
        // Add w16du:dateUtc for modern Word timezone handling
        XAttribute::new(W16DU::dateUtc(), &settings.date_time),
    ];
    
    doc.add_child(parent, XmlNodeData::element_with_attrs(W::ins(), attrs))
}

pub fn create_deletion(
    doc: &mut XmlDocument,
    parent: NodeId,
    settings: &RevisionSettings,
) -> NodeId {
    let rev_id = get_next_revision_id();
    let attrs = vec![
        XAttribute::new(W::id(), &rev_id.to_string()),  // w:id MUST come first per ECMA-376
        XAttribute::new(W::author(), &settings.author),
        XAttribute::new(W::date(), &settings.date_time),
        // Add w16du:dateUtc for modern Word timezone handling
        XAttribute::new(W16DU::dateUtc(), &settings.date_time),
    ];
    
    doc.add_child(parent, XmlNodeData::element_with_attrs(W::del(), attrs))
}

pub fn create_run_property_change(
    doc: &mut XmlDocument,
    parent: NodeId,
    settings: &RevisionSettings,
) -> NodeId {
    let rev_id = get_next_revision_id();
    let attrs = vec![
        XAttribute::new(W::id(), &rev_id.to_string()),  // w:id MUST come first per ECMA-376
        XAttribute::new(W::author(), &settings.author),
        XAttribute::new(W::date(), &settings.date_time),
    ];
    
    doc.add_child(parent, XmlNodeData::element_with_attrs(W::rPrChange(), attrs))
}

pub fn create_paragraph_property_change(
    doc: &mut XmlDocument,
    parent: NodeId,
    settings: &RevisionSettings,
) -> NodeId {
    let rev_id = get_next_revision_id();
    let attrs = vec![
        XAttribute::new(W::id(), &rev_id.to_string()),  // w:id MUST come first per ECMA-376
        XAttribute::new(W::author(), &settings.author),
        XAttribute::new(W::date(), &settings.date_time),
    ];
    
    doc.add_child(parent, XmlNodeData::element_with_attrs(W::pPrChange(), attrs))
}

pub fn create_text_run(doc: &mut XmlDocument, parent: NodeId, text: &str) -> NodeId {
    let run_id = doc.add_child(parent, XmlNodeData::element(W::r()));
    
    let needs_preserve = text.starts_with(' ') || text.ends_with(' ');
    let text_attrs = if needs_preserve {
        vec![XAttribute::new(XName::new("http://www.w3.org/XML/1998/namespace", "space"), "preserve")]
    } else {
        vec![]
    };
    
    let text_element = doc.add_child(run_id, XmlNodeData::element_with_attrs(W::t(), text_attrs));
    doc.add_child(text_element, XmlNodeData::Text(text.to_string()));
    
    run_id
}

pub fn create_paragraph(doc: &mut XmlDocument, parent: NodeId) -> NodeId {
    doc.add_child(parent, XmlNodeData::element(W::p()))
}

pub fn is_revision_element(name: &XName) -> bool {
    name.namespace.as_deref() == Some(W::NS) 
        && matches!(name.local_name.as_str(), "ins" | "del")
}

pub fn is_insertion(name: &XName) -> bool {
    name.namespace.as_deref() == Some(W::NS) && name.local_name == "ins"
}

pub fn is_deletion(name: &XName) -> bool {
    name.namespace.as_deref() == Some(W::NS) && name.local_name == "del"
}

pub fn is_format_change(name: &XName) -> bool {
    name.namespace.as_deref() == Some(W::NS) 
        && matches!(name.local_name.as_str(), "rPrChange" | "pPrChange")
}

static REVISION_ELEMENT_TAGS: &[&str] = &[
    "ins", "del", "rPrChange", "pPrChange", "sectPrChange",
    "tblPrChange", "tblGridChange", "trPrChange", "tcPrChange",
    "cellIns", "cellDel", "cellMerge",
    "customXmlInsRangeStart", "customXmlDelRangeStart",
    "customXmlMoveFromRangeStart", "customXmlMoveToRangeStart",
    "moveFrom", "moveTo", "moveFromRangeStart", "moveToRangeStart",
    "numberingChange",
];

pub fn is_revision_element_tag(local_name: &str) -> bool {
    REVISION_ELEMENT_TAGS.contains(&local_name)
}

pub fn fix_up_revision_ids(doc: &mut XmlDocument, start: NodeId) {
    let mut revision_nodes = Vec::new();
    
    for node_id in doc.descendants(start) {
        if let Some(data) = doc.get(node_id) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) 
                    && is_revision_element_tag(&name.local_name) 
                {
                    if let Some(attrs) = data.attributes() {
                        if attrs.iter().any(|a| a.name.local_name == "id" && a.name.namespace.as_deref() == Some(W::NS)) {
                            revision_nodes.push(node_id);
                        }
                    }
                }
            }
        }
    }
    
    let mut next_id = 0;
    for node_id in revision_nodes {
        if let Some(data) = doc.get_mut(node_id) {
            if let Some(attrs) = data.attributes_mut() {
                for attr in attrs.iter_mut() {
                    if attr.name.local_name == "id" && attr.name.namespace.as_deref() == Some(W::NS) {
                        attr.value = next_id.to_string();
                        next_id += 1;
                        break;
                    }
                }
            }
        }
    }
}

#[derive(Debug, Clone, Default)]
pub struct RevisionCounts {
    pub insertions: usize,
    pub deletions: usize,
    pub format_changes: usize,
}

impl RevisionCounts {
    pub fn total(&self) -> usize {
        self.insertions + self.deletions + self.format_changes
    }
}

/// Count revisions by grouping adjacent revision elements with the same type and attributes
/// This matches the C# GetRevisions algorithm which uses GroupAdjacent
pub fn count_revisions(doc: &XmlDocument, start: NodeId) -> RevisionCounts {
    let mut counts = RevisionCounts::default();
    
    // First pass: collect all paragraphs and their revision children
    // The C# groups by correlation status + revision attributes (author, date)
    // For simplicity, we group adjacent revision elements of the same type with same author/date
    
    count_revisions_in_subtree(doc, start, &mut counts, None);
    
    counts
}

/// Get a grouping key for a revision element (type + author + date, excluding id)
fn get_revision_key(doc: &XmlDocument, node: NodeId) -> Option<String> {
    let data = doc.get(node)?;
    let name = data.name()?;
    
    if name.namespace.as_deref() != Some(W::NS) {
        return None;
    }
    
    let rev_type = match name.local_name.as_str() {
        "ins" => "ins",
        "del" => "del",
        "rPrChange" => "rPrChange",
        "pPrChange" => "pPrChange",
        _ => return None,
    };
    
    // Build key from type + author + date (but NOT id)
    let mut key = rev_type.to_string();
    if let Some(attrs) = data.attributes() {
        for attr in attrs {
            if attr.name.namespace.as_deref() == Some(W::NS) {
                match attr.name.local_name.as_str() {
                    "author" => key.push_str(&format!("|author={}", attr.value)),
                    "date" => key.push_str(&format!("|date={}", attr.value)),
                    _ => {}
                }
            }
        }
    }
    Some(key)
}

/// Recursively count revisions, grouping adjacent siblings with same key
fn count_revisions_in_subtree(
    doc: &XmlDocument,
    node: NodeId,
    counts: &mut RevisionCounts,
    _parent_key: Option<&str>,
) {
    let children: Vec<_> = doc.children(node).collect();
    let mut last_key: Option<String> = None;
    
    for child in children {
        if let Some(key) = get_revision_key(doc, child) {
            // Only count if this is a new revision group (different key from last)
            let is_new_group = last_key.as_ref() != Some(&key);
            
            if is_new_group {
                // Count this as a new revision
                if key.starts_with("ins") {
                    counts.insertions += 1;
                } else if key.starts_with("del") {
                    counts.deletions += 1;
                } else if key.starts_with("rPrChange") || key.starts_with("pPrChange") {
                    counts.format_changes += 1;
                }
            }
            
            last_key = Some(key);
            
            // Don't recurse into revision elements - their content is part of the revision
        } else {
            // Not a revision element - reset the grouping and recurse
            last_key = None;
            count_revisions_in_subtree(doc, child, counts, None);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn revision_settings_default() {
        let settings = RevisionSettings::default();
        assert_eq!(settings.author, "redline-rs");
        assert!(!settings.date_time.is_empty());
    }

    #[test]
    fn revision_id_counter() {
        reset_revision_id_counter(100);
        assert_eq!(get_next_revision_id(), 100);
        assert_eq!(get_next_revision_id(), 101);
        assert_eq!(get_next_revision_id(), 102);
    }

    #[test]
    fn revision_id_starts_at_zero() {
        reset_revision_id_counter(0);
        assert_eq!(get_next_revision_id(), 0);
        assert_eq!(get_next_revision_id(), 1);
        assert_eq!(get_next_revision_id(), 2);
    }

    #[test]
    fn is_revision_element_checks() {
        assert!(is_insertion(&W::ins()));
        assert!(!is_insertion(&W::del()));
        assert!(is_deletion(&W::del()));
        assert!(!is_deletion(&W::ins()));
        assert!(is_format_change(&W::rPrChange()));
        assert!(is_format_change(&W::pPrChange()));
        assert!(!is_format_change(&W::ins()));
    }

    #[test]
    fn create_text_run_works() {
        let mut doc = XmlDocument::new();
        let root = doc.add_root(XmlNodeData::element(W::body()));
        let run = create_text_run(&mut doc, root, "Hello");
        
        assert!(doc.get(run).is_some());
        let run_data = doc.get(run).unwrap();
        assert_eq!(run_data.name().map(|n| &n.local_name), Some(&"r".to_string()));
    }
}
