//! Extract change metadata from Word documents with revision markup.
//!
//! This module walks a document that contains `w:ins`, `w:del`, and `w:rPrChange`
//! elements and extracts structured change data suitable for UI display.

use super::types::{WmlChange, WmlChangeType, WmlWordCount};
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::W;
use crate::xml::node::XmlNodeData;
use indextree::NodeId;

/// Context tracked while walking the document tree
#[derive(Debug, Clone, Default)]
struct ExtractionContext {
    /// Current paragraph index (0-based)
    paragraph_index: usize,
    /// Current table row index (if inside a table)
    table_row_index: Option<usize>,
    /// Current table cell index (if inside a table cell)
    table_cell_index: Option<usize>,
    /// Whether we're inside a table
    in_table: bool,
    /// Whether we're inside a footnote
    in_footnote: bool,
    /// Whether we're inside an endnote
    in_endnote: bool,
    /// Whether we're inside a textbox
    in_textbox: bool,
}

/// Extract all changes from a document that contains revision markup.
///
/// This function walks the document tree and extracts each `w:ins`, `w:del`,
/// and `w:rPrChange` element into a structured `WmlChange` record.
///
/// # Arguments
/// * `doc` - The XML document to extract changes from
/// * `root` - The root node to start extraction from (usually document body)
/// * `default_author` - Default author name if not specified on revision
/// * `default_date` - Default date/time if not specified on revision
///
/// # Returns
/// A vector of `WmlChange` records in document order
pub fn extract_changes_from_document(
    doc: &XmlDocument,
    root: NodeId,
    default_author: Option<&str>,
    default_date: Option<&str>,
) -> Vec<WmlChange> {
    let mut changes = Vec::new();
    let mut context = ExtractionContext::default();

    walk_node(
        doc,
        root,
        &mut context,
        &mut changes,
        default_author,
        default_date,
    );

    changes
}

/// Recursively walk the document tree and extract changes
fn walk_node(
    doc: &XmlDocument,
    node: NodeId,
    context: &mut ExtractionContext,
    changes: &mut Vec<WmlChange>,
    default_author: Option<&str>,
    default_date: Option<&str>,
) {
    let Some(data) = doc.get(node) else { return };
    let Some(name) = data.name() else {
        // Text or other non-element node - recurse into children
        for child in doc.children(node) {
            walk_node(doc, child, context, changes, default_author, default_date);
        }
        return;
    };

    // Track context based on element type
    let is_w_ns = name.namespace.as_deref() == Some(W::NS);
    let local = name.local_name.as_str();

    // Update context based on structural elements
    if is_w_ns {
        match local {
            "p" => {
                context.paragraph_index += 1;
            }
            "tbl" => {
                context.in_table = true;
                context.table_row_index = None;
            }
            "tr" => {
                context.table_row_index = Some(context.table_row_index.map(|i| i + 1).unwrap_or(0));
                context.table_cell_index = None;
            }
            "tc" => {
                context.table_cell_index =
                    Some(context.table_cell_index.map(|i| i + 1).unwrap_or(0));
            }
            "txbxContent" => {
                context.in_textbox = true;
            }
            "footnote" => {
                context.in_footnote = true;
            }
            "endnote" => {
                context.in_endnote = true;
            }
            _ => {}
        }
    }

    // Check for revision elements
    if is_w_ns {
        match local {
            "ins" => {
                let change = extract_insertion(doc, node, context, default_author, default_date);
                changes.push(change);
                // Don't recurse into w:ins - we've handled its content
                return;
            }
            "del" => {
                let change = extract_deletion(doc, node, context, default_author, default_date);
                changes.push(change);
                // Don't recurse into w:del - we've handled its content
                return;
            }
            "rPrChange" => {
                let change =
                    extract_format_change(doc, node, context, default_author, default_date);
                changes.push(change);
                // Don't recurse into w:rPrChange
                return;
            }
            _ => {}
        }
    }

    // Recurse into children
    for child in doc.children(node) {
        walk_node(doc, child, context, changes, default_author, default_date);
    }

    // Reset context when leaving structural elements
    if is_w_ns {
        match local {
            "tbl" => {
                context.in_table = false;
                context.table_row_index = None;
                context.table_cell_index = None;
            }
            "txbxContent" => {
                context.in_textbox = false;
            }
            "footnote" => {
                context.in_footnote = false;
            }
            "endnote" => {
                context.in_endnote = false;
            }
            _ => {}
        }
    }
}

/// Extract an insertion change from a w:ins element
fn extract_insertion(
    doc: &XmlDocument,
    node: NodeId,
    context: &ExtractionContext,
    default_author: Option<&str>,
    default_date: Option<&str>,
) -> WmlChange {
    let (revision_id, author, date_time) =
        extract_revision_attrs(doc, node, default_author, default_date);
    let text = extract_text_content(doc, node);
    let word_count = count_words(&text);

    WmlChange {
        change_type: WmlChangeType::TextInserted,
        revision_id,
        paragraph_index: Some(context.paragraph_index),
        table_row_index: context.table_row_index,
        table_cell_index: context.table_cell_index,
        old_text: None,
        new_text: Some(text),
        word_count: Some(WmlWordCount {
            deleted: 0,
            inserted: word_count,
        }),
        format_description: None,
        author,
        date_time,
        in_footnote: context.in_footnote,
        in_endnote: context.in_endnote,
        in_table: context.in_table,
        in_textbox: context.in_textbox,
    }
}

/// Extract a deletion change from a w:del element
fn extract_deletion(
    doc: &XmlDocument,
    node: NodeId,
    context: &ExtractionContext,
    default_author: Option<&str>,
    default_date: Option<&str>,
) -> WmlChange {
    let (revision_id, author, date_time) =
        extract_revision_attrs(doc, node, default_author, default_date);
    let text = extract_deleted_text_content(doc, node);
    let word_count = count_words(&text);

    WmlChange {
        change_type: WmlChangeType::TextDeleted,
        revision_id,
        paragraph_index: Some(context.paragraph_index),
        table_row_index: context.table_row_index,
        table_cell_index: context.table_cell_index,
        old_text: Some(text),
        new_text: None,
        word_count: Some(WmlWordCount {
            deleted: word_count,
            inserted: 0,
        }),
        format_description: None,
        author,
        date_time,
        in_footnote: context.in_footnote,
        in_endnote: context.in_endnote,
        in_table: context.in_table,
        in_textbox: context.in_textbox,
    }
}

/// Extract a format change from a w:rPrChange element
fn extract_format_change(
    doc: &XmlDocument,
    node: NodeId,
    context: &ExtractionContext,
    default_author: Option<&str>,
    default_date: Option<&str>,
) -> WmlChange {
    let (revision_id, author, date_time) =
        extract_revision_attrs(doc, node, default_author, default_date);

    WmlChange {
        change_type: WmlChangeType::FormatChanged,
        revision_id,
        paragraph_index: Some(context.paragraph_index),
        table_row_index: context.table_row_index,
        table_cell_index: context.table_cell_index,
        old_text: None,
        new_text: None,
        word_count: None,
        format_description: Some("Format changed".to_string()),
        author,
        date_time,
        in_footnote: context.in_footnote,
        in_endnote: context.in_endnote,
        in_table: context.in_table,
        in_textbox: context.in_textbox,
    }
}

/// Extract revision attributes (id, author, date) from an element
fn extract_revision_attrs(
    doc: &XmlDocument,
    node: NodeId,
    default_author: Option<&str>,
    default_date: Option<&str>,
) -> (i32, Option<String>, Option<String>) {
    let data = doc.get(node);
    let attrs = data.and_then(|d| d.attributes()).unwrap_or(&[]);

    let mut revision_id = 0;
    let mut author = default_author.map(|s| s.to_string());
    let mut date_time = default_date.map(|s| s.to_string());

    for attr in attrs {
        let local = attr.name.local_name.as_str();
        let is_w_ns = attr.name.namespace.as_deref() == Some(W::NS);

        if is_w_ns {
            match local {
                "id" => {
                    revision_id = attr.value.parse().unwrap_or(0);
                }
                "author" => {
                    author = Some(attr.value.clone());
                }
                "date" => {
                    date_time = Some(attr.value.clone());
                }
                _ => {}
            }
        }
    }

    (revision_id, author, date_time)
}

/// Extract text content from w:t elements inside a node (for insertions)
fn extract_text_content(doc: &XmlDocument, node: NodeId) -> String {
    let mut text = String::new();
    collect_text_recursive(doc, node, &mut text, false);
    text
}

/// Extract text content from w:delText elements inside a node (for deletions)
fn extract_deleted_text_content(doc: &XmlDocument, node: NodeId) -> String {
    let mut text = String::new();
    collect_text_recursive(doc, node, &mut text, true);
    text
}

/// Recursively collect text from w:t or w:delText elements
fn collect_text_recursive(doc: &XmlDocument, node: NodeId, text: &mut String, is_deletion: bool) {
    let Some(data) = doc.get(node) else { return };

    if let Some(name) = data.name() {
        let is_w_ns = name.namespace.as_deref() == Some(W::NS);
        let local = name.local_name.as_str();

        // Check if this is a text element we want
        let is_text_element = if is_deletion {
            is_w_ns && local == "delText"
        } else {
            is_w_ns && local == "t"
        };

        if is_text_element {
            // Get text content from children
            for child in doc.children(node) {
                if let Some(child_data) = doc.get(child) {
                    if let XmlNodeData::Text(child_text) = child_data {
                        text.push_str(child_text);
                    }
                }
            }
            return;
        }
    }

    // Recurse into children
    for child in doc.children(node) {
        collect_text_recursive(doc, child, text, is_deletion);
    }
}

/// Count words in a text string
fn count_words(text: &str) -> usize {
    text.split_whitespace().count()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::parser;

    #[test]
    fn test_count_words() {
        assert_eq!(count_words(""), 0);
        assert_eq!(count_words("hello"), 1);
        assert_eq!(count_words("hello world"), 2);
        assert_eq!(count_words("  hello   world  "), 2);
    }

    #[test]
    fn test_extract_insertions() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r>
                        <w:ins w:id="1" w:author="Test Author" w:date="2025-01-01T10:00:00Z">
                            <w:r>
                                <w:t>inserted text</w:t>
                            </w:r>
                        </w:ins>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>"#;

        let doc = parser::parse_bytes(xml.as_bytes()).unwrap();
        let body = crate::wml::find_document_body(&doc).unwrap();
        let changes = extract_changes_from_document(&doc, body, None, None);

        assert_eq!(changes.len(), 1);
        assert_eq!(changes[0].change_type, WmlChangeType::TextInserted);
        assert_eq!(changes[0].revision_id, 1);
        assert_eq!(changes[0].new_text.as_deref(), Some("inserted text"));
        assert_eq!(changes[0].author.as_deref(), Some("Test Author"));
    }

    #[test]
    fn test_extract_deletions() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r>
                        <w:del w:id="2" w:author="Test Author">
                            <w:r>
                                <w:delText>deleted text</w:delText>
                            </w:r>
                        </w:del>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>"#;

        let doc = parser::parse_bytes(xml.as_bytes()).unwrap();
        let body = crate::wml::find_document_body(&doc).unwrap();
        let changes = extract_changes_from_document(&doc, body, None, None);

        assert_eq!(changes.len(), 1);
        assert_eq!(changes[0].change_type, WmlChangeType::TextDeleted);
        assert_eq!(changes[0].revision_id, 2);
        assert_eq!(changes[0].old_text.as_deref(), Some("deleted text"));
    }

    #[test]
    fn test_extract_multiple_changes() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:del w:id="1"><w:r><w:delText>old</w:delText></w:r></w:del>
                    <w:ins w:id="2"><w:r><w:t>new</w:t></w:r></w:ins>
                </w:p>
            </w:body>
        </w:document>"#;

        let doc = parser::parse_bytes(xml.as_bytes()).unwrap();
        let body = crate::wml::find_document_body(&doc).unwrap();
        let changes = extract_changes_from_document(&doc, body, None, None);

        assert_eq!(changes.len(), 2);
        assert_eq!(changes[0].change_type, WmlChangeType::TextDeleted);
        assert_eq!(changes[1].change_type, WmlChangeType::TextInserted);
    }
}
