//! MarkupSimplifier - Port of C# MarkupSimplifier.cs

use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::W;
use indextree::NodeId;

#[derive(Debug, Clone, Default)]
pub struct SimplifyMarkupSettings {
    pub accept_revisions: bool,
    pub normalize_xml: bool,
    pub remove_bookmarks: bool,
    pub remove_comments: bool,
    pub remove_content_controls: bool,
    pub remove_end_and_foot_notes: bool,
    pub remove_field_codes: bool,
    pub remove_go_back_bookmark: bool,
    pub remove_hyperlinks: bool,
    pub remove_last_rendered_page_break: bool,
    pub remove_markup_for_document_comparison: bool,
    pub remove_permissions: bool,
    pub remove_proof: bool,
    pub remove_rsid_info: bool,
    pub remove_smart_tags: bool,
    pub remove_soft_hyphens: bool,
    pub remove_web_hidden: bool,
    pub replace_tabs_with_spaces: bool,
}

pub fn simplify_markup(doc: &mut XmlDocument, root: NodeId, settings: &SimplifyMarkupSettings) {
    if settings.remove_markup_for_document_comparison {
        let mut settings_copy = settings.clone();
        settings_copy.remove_rsid_info = true;
        remove_elements_for_document_comparison(doc, root);
    }

    if settings.remove_rsid_info {
        remove_rsid_info_in_settings(doc, root);
    }

    simplify_markup_for_part(doc, root, settings);
}

fn remove_rsid_info_in_settings(doc: &mut XmlDocument, root: NodeId) {
    let rsids_to_remove: Vec<NodeId> = doc
        .descendants(root)
        .filter(|&node_id| {
            doc.get(node_id)
                .and_then(|data| data.name())
                .map(|n| *n == W::rsids())
                .unwrap_or(false)
        })
        .collect();

    for rsid_node in rsids_to_remove {
        doc.remove(rsid_node);
    }
}

fn remove_elements_for_document_comparison(doc: &mut XmlDocument, root: NodeId) {
    let goback_bookmarks: Vec<NodeId> = doc
        .descendants(root)
        .filter(|&node_id| {
            if let Some(data) = doc.get(node_id) {
                if let Some(name) = data.name() {
                    if *name == W::bookmark_start() {
                        if let Some(attrs) = data.attributes() {
                            return attrs
                                .iter()
                                .any(|attr| attr.name == W::name() && attr.value == "_GoBack");
                        }
                    }
                }
            }
            false
        })
        .collect();

    for bookmark_start in goback_bookmarks {
        if let Some(data) = doc.get(bookmark_start) {
            if let Some(attrs) = data.attributes() {
                let bookmark_id = attrs
                    .iter()
                    .find(|a| a.name == W::id())
                    .map(|a| a.value.clone());

                if let Some(id) = bookmark_id {
                    let bookmark_ends: Vec<NodeId> = doc
                        .descendants(root)
                        .filter(|&node_id| {
                            if let Some(data) = doc.get(node_id) {
                                if let Some(name) = data.name() {
                                    if *name == W::bookmark_end() {
                                        if let Some(attrs) = data.attributes() {
                                            return attrs.iter().any(|attr| {
                                                attr.name == W::id() && attr.value == id
                                            });
                                        }
                                    }
                                }
                            }
                            false
                        })
                        .collect();

                    for end in bookmark_ends {
                        doc.remove(end);
                    }
                }
            }
        }
        doc.remove(bookmark_start);
    }
}

fn simplify_markup_for_part(
    doc: &mut XmlDocument,
    root: NodeId,
    settings: &SimplifyMarkupSettings,
) {
    if settings.remove_rsid_info {
        remove_rsid_transform(doc, root);
    }

    if settings.remove_content_controls {
        remove_content_controls(doc, root);
    }

    if settings.remove_smart_tags {
        remove_smart_tags(doc, root);
    }

    if settings.remove_hyperlinks {
        remove_hyperlinks(doc, root);
    }
}

fn remove_rsid_transform(doc: &mut XmlDocument, root: NodeId) {
    let nodes: Vec<NodeId> = doc.descendants(root).collect();

    for node_id in nodes {
        if let Some(data) = doc.get_mut(node_id) {
            if let Some(attrs) = data.attributes_mut() {
                attrs.retain(|attr| !attr.name.local_name.starts_with("rsid"));
            }
        }
    }
}

fn remove_content_controls(doc: &mut XmlDocument, root: NodeId) {
    let sdt_nodes: Vec<NodeId> = doc
        .descendants(root)
        .filter(|&node_id| {
            doc.get(node_id)
                .and_then(|data| data.name())
                .map(|n| *n == W::sdt())
                .unwrap_or(false)
        })
        .collect();

    for sdt in sdt_nodes {
        let sdt_content_children: Vec<NodeId> = doc
            .children(sdt)
            .filter(|&child_id| {
                doc.get(child_id)
                    .and_then(|data| data.name())
                    .map(|n| *n == W::sdt_content())
                    .unwrap_or(false)
            })
            .flat_map(|sdt_content| doc.children(sdt_content).collect::<Vec<_>>())
            .collect();

        if let Some(parent) = doc.parent(sdt) {
            for child in sdt_content_children {
                doc.reparent(parent, child);
            }
        }
        doc.remove(sdt);
    }
}

fn remove_smart_tags(doc: &mut XmlDocument, root: NodeId) {
    let smart_tag_nodes: Vec<NodeId> = doc
        .descendants(root)
        .filter(|&node_id| {
            doc.get(node_id)
                .and_then(|data| data.name())
                .map(|n| *n == W::smart_tag())
                .unwrap_or(false)
        })
        .collect();

    for smart_tag in smart_tag_nodes {
        let children: Vec<NodeId> = doc.children(smart_tag).collect();

        if let Some(parent) = doc.parent(smart_tag) {
            for child in children {
                doc.reparent(parent, child);
            }
        }
        doc.remove(smart_tag);
    }
}

/// Remove hyperlinks by unwrapping their content
///
/// Corresponds to C# MarkupSimplifier.cs line 370-372:
/// ```csharp
/// if (settings.RemoveHyperlinks && (element.Name == W.hyperlink))
///     return element.Elements();
/// ```
fn remove_hyperlinks(doc: &mut XmlDocument, root: NodeId) {
    let hyperlink_nodes: Vec<NodeId> = doc
        .descendants(root)
        .filter(|&node_id| {
            doc.get(node_id)
                .and_then(|data| data.name())
                .map(|n| *n == W::hyperlink())
                .unwrap_or(false)
        })
        .collect();

    for hyperlink in hyperlink_nodes {
        let children: Vec<NodeId> = doc.children(hyperlink).collect();

        if let Some(parent) = doc.parent(hyperlink) {
            for child in children {
                doc.reparent(parent, child);
            }
        }
        doc.remove(hyperlink);
    }
}

pub fn merge_adjacent_superfluous_runs(_doc: &mut XmlDocument, _root: NodeId) {
    todo!("Port from C# MarkupSimplifier.cs:MergeAdjacentSuperfluousRuns")
}

pub fn transform_element_to_single_character_runs(_doc: &mut XmlDocument, _root: NodeId) {
    todo!("Port from C# MarkupSimplifier.cs:TransformElementToSingleCharacterRuns")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simplify_markup_settings_default() {
        let settings = SimplifyMarkupSettings::default();
        assert!(!settings.accept_revisions);
        assert!(!settings.remove_bookmarks);
    }
}
