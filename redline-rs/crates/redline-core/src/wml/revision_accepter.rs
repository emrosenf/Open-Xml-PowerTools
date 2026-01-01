use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{M, W};
use crate::xml::node::XmlNodeData;
use crate::xml::xname::XAttribute;
use indextree::NodeId;
use std::collections::HashSet;

pub fn accept_revisions(source: &XmlDocument, source_root: NodeId) -> XmlDocument {
    process_revisions(source, source_root, &AcceptAllStrategy)
}

pub fn accept_revisions_by_id(
    source: &XmlDocument,
    source_root: NodeId,
    revision_ids: &[i32],
) -> XmlDocument {
    let ids: HashSet<i32> = revision_ids.iter().cloned().collect();
    process_revisions(source, source_root, &AcceptByIdStrategy { ids })
}

pub fn reject_revisions_by_id(
    source: &XmlDocument,
    source_root: NodeId,
    revision_ids: &[i32],
) -> XmlDocument {
    let ids: HashSet<i32> = revision_ids.iter().cloned().collect();
    process_revisions(source, source_root, &RejectByIdStrategy { ids })
}

trait RevisionStrategy {
    fn decide(&self, name: &str, id: Option<i32>) -> RevisionAction;
    fn should_remove_deleted_containers(&self) -> bool {
        false
    }
}

// Removed duplicate definitions

enum RevisionAction {
    Keep,   // Keep element as is
    Remove, // Remove element and its content
    Unwrap, // Remove element wrapper, keep content
}

struct AcceptAllStrategy;

impl RevisionStrategy for AcceptAllStrategy {
    fn decide(&self, name: &str, _id: Option<i32>) -> RevisionAction {
        match name {
            "ins" | "moveTo" => RevisionAction::Unwrap,
            "del"
            | "delText"
            | "delInstrText"
            | "moveFrom"
            | "pPrChange"
            | "rPrChange"
            | "tblPrChange"
            | "tblGridChange"
            | "tcPrChange"
            | "trPrChange"
            | "tblPrExChange"
            | "sectPrChange"
            | "numberingChange"
            | "cellIns"
            | "customXmlDelRangeStart"
            | "customXmlDelRangeEnd"
            | "customXmlInsRangeStart"
            | "customXmlInsRangeEnd"
            | "customXmlMoveFromRangeStart"
            | "customXmlMoveFromRangeEnd"
            | "customXmlMoveToRangeStart"
            | "customXmlMoveToRangeEnd"
            | "moveFromRangeStart"
            | "moveFromRangeEnd"
            | "moveToRangeStart"
            | "moveToRangeEnd" => RevisionAction::Remove,
            _ => RevisionAction::Keep,
        }
    }
}

struct AcceptByIdStrategy {
    ids: HashSet<i32>,
}

impl RevisionStrategy for AcceptByIdStrategy {
    fn decide(&self, name: &str, id: Option<i32>) -> RevisionAction {
        let is_target = id.map_or(false, |i| self.ids.contains(&i));
        if !is_target {
            return RevisionAction::Keep;
        }

        match name {
            "ins" => RevisionAction::Unwrap,
            "del" => RevisionAction::Remove,
            "rPrChange" | "pPrChange" | "sectPrChange" => RevisionAction::Remove, // Accepting format change = keeping current format (removing history)
            // Handle other change types as needed
            _ => RevisionAction::Keep,
        }
    }
}

struct RejectByIdStrategy {
    ids: HashSet<i32>,
}

impl RevisionStrategy for RejectByIdStrategy {
    fn decide(&self, name: &str, id: Option<i32>) -> RevisionAction {
        let is_target = id.map_or(false, |i| self.ids.contains(&i));
        if !is_target {
            return RevisionAction::Keep;
        }

        match name {
            "ins" => RevisionAction::Remove, // Reject insertion = remove content
            "del" => RevisionAction::Unwrap, // Reject deletion = restore content
            "rPrChange" | "pPrChange" => RevisionAction::Remove, // Reject format change = TODO: revert format? For now just remove marker?
            // NOTE: Reverting format is complex (need to merge old properties).
            // For MVP, removing the change marker keeps current format (Accept behavior).
            // To truly reject, we would need to apply the properties inside rPrChange to the parent.
            // This implementation currently treats Reject Format as Accept Format (limitation).
            _ => RevisionAction::Keep,
        }
    }
}

fn process_revisions<S: RevisionStrategy>(
    source: &XmlDocument,
    source_root: NodeId,
    strategy: &S,
) -> XmlDocument {
    let mut result = XmlDocument::new();

    if let Some(children) = transform_node(source, source_root, &mut result, None, strategy) {
        if children.len() == 1 {
            result.set_root(Some(children[0]));
        }
    }

    result
}

fn get_id_attr(data: &XmlNodeData) -> Option<i32> {
    data.attributes()?.iter().find_map(|attr| {
        if attr.name.local_name == "id" && attr.name.namespace.as_deref() == Some(W::NS) {
            attr.value.parse().ok()
        } else {
            None
        }
    })
}

fn transform_node<S: RevisionStrategy>(
    source: &XmlDocument,
    node_id: NodeId,
    result: &mut XmlDocument,
    parent: Option<NodeId>,
    strategy: &S,
) -> Option<Vec<NodeId>> {
    let data = source.get(node_id)?;

    match data {
        XmlNodeData::Text(text) => {
            let new_id = if let Some(p) = parent {
                result.add_child(p, XmlNodeData::Text(text.clone()))
            } else {
                result.add_root(XmlNodeData::Text(text.clone()))
            };
            Some(vec![new_id])
        }
        XmlNodeData::CData(text) => {
            let new_id = if let Some(p) = parent {
                result.add_child(p, XmlNodeData::CData(text.clone()))
            } else {
                result.add_root(XmlNodeData::CData(text.clone()))
            };
            Some(vec![new_id])
        }
        XmlNodeData::Comment(text) => {
            let new_id = if let Some(p) = parent {
                result.add_child(p, XmlNodeData::Comment(text.clone()))
            } else {
                result.add_root(XmlNodeData::Comment(text.clone()))
            };
            Some(vec![new_id])
        }
        XmlNodeData::ProcessingInstruction {
            target,
            data: pi_data,
        } => {
            let new_id = if let Some(p) = parent {
                result.add_child(
                    p,
                    XmlNodeData::ProcessingInstruction {
                        target: target.clone(),
                        data: pi_data.clone(),
                    },
                )
            } else {
                result.add_root(XmlNodeData::ProcessingInstruction {
                    target: target.clone(),
                    data: pi_data.clone(),
                })
            };
            Some(vec![new_id])
        }
        XmlNodeData::Element { name, attributes } => {
            let local = &name.local_name;
            let ns = name.namespace.as_deref();

            let id = get_id_attr(data);

            // Check strategy for this element
            let action = if ns == Some(W::NS) {
                strategy.decide(local, id)
            } else {
                RevisionAction::Keep
            };

            match action {
                RevisionAction::Remove => return None,
                RevisionAction::Unwrap => {
                    let mut unwrapped = Vec::new();
                    for child in source.children(node_id) {
                        if let Some(children) =
                            transform_node(source, child, result, parent, strategy)
                        {
                            unwrapped.extend(children);
                        }
                    }
                    return if unwrapped.is_empty() {
                        None
                    } else {
                        Some(unwrapped)
                    };
                }
                RevisionAction::Keep => {
                    // Continue with normal processing
                }
            }

            if ns == Some(W::NS) && local == "tr" && is_deleted_table_row(source, node_id) {
                if strategy.should_remove_deleted_containers() {
                    return None;
                }
            }

            if ns == Some(M::NS) && local == "f" && has_deleted_math_control(source, node_id) {
                if strategy.should_remove_deleted_containers() {
                    return None;
                }
            }

            let filtered_attrs = filter_rsid_attributes(attributes);

            let new_id = if let Some(p) = parent {
                result.add_child(
                    p,
                    XmlNodeData::element_with_attrs(name.clone(), filtered_attrs),
                )
            } else {
                result.add_root(XmlNodeData::element_with_attrs(
                    name.clone(),
                    filtered_attrs,
                ))
            };

            for child in source.children(node_id) {
                transform_node(source, child, result, Some(new_id), strategy);
            }

            Some(vec![new_id])
        }
    }
}

fn filter_rsid_attributes(attributes: &[XAttribute]) -> Vec<XAttribute> {
    // Note: We preserve rsid*, paraId, and textId attributes during revision acceptance.
    // These are important for document identity and should appear in the output.
    // They are stripped only during hash computation (in block_hash.rs) for comparison purposes.
    attributes.to_vec()
}

fn is_deleted_table_row(doc: &XmlDocument, tr_node: NodeId) -> bool {
    for child in doc.children(tr_node) {
        if let Some(data) = doc.get(child) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "trPr" {
                    for pr_child in doc.children(child) {
                        if let Some(pr_data) = doc.get(pr_child) {
                            if let Some(pr_name) = pr_data.name() {
                                if pr_name.namespace.as_deref() == Some(W::NS)
                                    && pr_name.local_name == "del"
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    false
}

fn has_deleted_math_control(doc: &XmlDocument, mf_node: NodeId) -> bool {
    for child in doc.children(mf_node) {
        if let Some(data) = doc.get(child) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(M::NS) && name.local_name == "fPr" {
                    for fpr_child in doc.children(child) {
                        if let Some(fpr_data) = doc.get(fpr_child) {
                            if let Some(fpr_name) = fpr_data.name() {
                                if fpr_name.namespace.as_deref() == Some(M::NS)
                                    && fpr_name.local_name == "ctrlPr"
                                {
                                    for ctrl_child in doc.children(fpr_child) {
                                        if let Some(ctrl_data) = doc.get(ctrl_child) {
                                            if let Some(ctrl_name) = ctrl_data.name() {
                                                if ctrl_name.namespace.as_deref() == Some(W::NS)
                                                    && ctrl_name.local_name == "del"
                                                {
                                                    return true;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::xname::XName;

    #[test]
    fn accept_revisions_removes_del() {
        let mut source = XmlDocument::new();
        let body = source.add_root(XmlNodeData::element(W::body()));
        let para = source.add_child(body, XmlNodeData::element(W::p()));
        source.add_child(para, XmlNodeData::element(W::del()));
        source.add_child(para, XmlNodeData::element(W::r()));

        let result = accept_revisions(&source, body);

        let result_body = result.root().unwrap();
        let children: Vec<_> = result.children(result_body).collect();
        assert_eq!(children.len(), 1);

        let para_children: Vec<_> = result.children(children[0]).collect();
        assert_eq!(para_children.len(), 1);

        let run_data = result.get(para_children[0]).unwrap();
        assert_eq!(run_data.name().map(|n| n.local_name.as_str()), Some("r"));
    }

    #[test]
    fn accept_revisions_unwraps_ins() {
        let mut source = XmlDocument::new();
        let body = source.add_root(XmlNodeData::element(W::body()));
        let ins = source.add_child(body, XmlNodeData::element(W::ins()));
        source.add_child(ins, XmlNodeData::element(W::r()));

        let result = accept_revisions(&source, body);

        let result_body = result.root().unwrap();
        let children: Vec<_> = result.children(result_body).collect();
        assert_eq!(children.len(), 1);

        let run_data = result.get(children[0]).unwrap();
        assert_eq!(run_data.name().map(|n| n.local_name.as_str()), Some("r"));
    }

    #[test]
    fn filter_rsid_preserves_all_attrs() {
        // Note: rsid, paraId, and textId attributes are now preserved during revision acceptance.
        // They are stripped only during hash computation (in block_hash.rs) for comparison purposes.
        let attrs = vec![
            XAttribute::new(XName::new(W::NS, "rsidR"), "00123456"),
            XAttribute::new(XName::new(W::NS, "val"), "test"),
        ];

        let filtered = filter_rsid_attributes(&attrs);

        // All attributes should be preserved
        assert_eq!(filtered.len(), 2);
        assert!(filtered.iter().any(|a| a.name.local_name == "rsidR"));
        assert!(filtered.iter().any(|a| a.name.local_name == "val"));
    }

    #[test]
    fn accept_revisions_preserves_pt_unid() {
        use crate::xml::namespaces::PT;

        let mut source = XmlDocument::new();
        let body = source.add_root(XmlNodeData::element(W::body()));

        // Add pt:Unid to the body element
        let pt_unid = PT::Unid();
        source.set_attribute(body, &pt_unid, "body-unid-123");

        // Add a paragraph with pt:Unid
        let para = source.add_child(body, XmlNodeData::element(W::p()));
        source.set_attribute(para, &pt_unid, "para-unid-456");

        // Add a run with pt:Unid
        let run = source.add_child(para, XmlNodeData::element(W::r()));
        source.set_attribute(run, &pt_unid, "run-unid-789");

        let result = accept_revisions(&source, body);

        // Verify the result document has the pt:Unid attributes preserved
        let result_body = result.root().unwrap();
        let body_attrs = result.get(result_body).unwrap().attributes().unwrap();
        assert!(
            body_attrs
                .iter()
                .any(|a| a.name == pt_unid && a.value == "body-unid-123"),
            "Body should have pt:Unid preserved"
        );

        let para_node = result.children(result_body).next().unwrap();
        let para_attrs = result.get(para_node).unwrap().attributes().unwrap();
        assert!(
            para_attrs
                .iter()
                .any(|a| a.name == pt_unid && a.value == "para-unid-456"),
            "Paragraph should have pt:Unid preserved"
        );

        let run_node = result.children(para_node).next().unwrap();
        let run_attrs = result.get(run_node).unwrap().attributes().unwrap();
        assert!(
            run_attrs
                .iter()
                .any(|a| a.name == pt_unid && a.value == "run-unid-789"),
            "Run should have pt:Unid preserved"
        );
    }

    #[test]
    fn accept_revisions_preserves_pt_unid_with_document_root() {
        use crate::xml::namespaces::PT;

        // This test mirrors the actual usage in comparer.rs:
        // accept_revisions(&doc1, doc1_root) where doc1_root is w:document

        let mut source = XmlDocument::new();
        // Create document structure: w:document > w:body > w:p > w:r
        let document = source.add_root(XmlNodeData::element(XName::new(W::NS, "document")));
        let body = source.add_child(document, XmlNodeData::element(W::body()));
        let para = source.add_child(body, XmlNodeData::element(W::p()));
        let run = source.add_child(para, XmlNodeData::element(W::r()));

        // Add pt:Unid to all elements
        let pt_unid = PT::Unid();
        source.set_attribute(document, &pt_unid, "doc-unid-000");
        source.set_attribute(body, &pt_unid, "body-unid-123");
        source.set_attribute(para, &pt_unid, "para-unid-456");
        source.set_attribute(run, &pt_unid, "run-unid-789");

        // Accept revisions starting from document root (like comparer.rs does)
        let result = accept_revisions(&source, document);

        // Verify the result document has the pt:Unid attributes preserved
        let result_document = result.root().unwrap();
        let doc_attrs = result.get(result_document).unwrap().attributes().unwrap();
        assert!(
            doc_attrs
                .iter()
                .any(|a| a.name == pt_unid && a.value == "doc-unid-000"),
            "Document should have pt:Unid preserved"
        );

        let result_body = result.children(result_document).next().unwrap();
        let body_attrs = result.get(result_body).unwrap().attributes().unwrap();
        assert!(
            body_attrs
                .iter()
                .any(|a| a.name == pt_unid && a.value == "body-unid-123"),
            "Body should have pt:Unid preserved"
        );

        let para_node = result.children(result_body).next().unwrap();
        let para_attrs = result.get(para_node).unwrap().attributes().unwrap();
        assert!(
            para_attrs
                .iter()
                .any(|a| a.name == pt_unid && a.value == "para-unid-456"),
            "Paragraph should have pt:Unid preserved"
        );

        let run_node = result.children(para_node).next().unwrap();
        let run_attrs = result.get(run_node).unwrap().attributes().unwrap();
        assert!(
            run_attrs
                .iter()
                .any(|a| a.name == pt_unid && a.value == "run-unid-789"),
            "Run should have pt:Unid preserved"
        );
    }
}
