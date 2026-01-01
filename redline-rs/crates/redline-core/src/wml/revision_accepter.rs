use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{M, W};
use crate::xml::node::XmlNodeData;
use crate::xml::xname::XAttribute;
use indextree::NodeId;
use once_cell::sync::Lazy;
use std::collections::HashSet;

static ELEMENTS_TO_REMOVE: Lazy<HashSet<&'static str>> = Lazy::new(|| {
    let mut s = HashSet::new();
    s.insert("del");
    s.insert("delText");
    s.insert("delInstrText");
    s.insert("moveFrom");
    s.insert("pPrChange");
    s.insert("rPrChange");
    s.insert("tblPrChange");
    s.insert("tblGridChange");
    s.insert("tcPrChange");
    s.insert("trPrChange");
    s.insert("tblPrExChange");
    s.insert("sectPrChange");
    s.insert("numberingChange");
    s.insert("cellIns");
    s.insert("customXmlDelRangeStart");
    s.insert("customXmlDelRangeEnd");
    s.insert("customXmlInsRangeStart");
    s.insert("customXmlInsRangeEnd");
    s.insert("customXmlMoveFromRangeStart");
    s.insert("customXmlMoveFromRangeEnd");
    s.insert("customXmlMoveToRangeStart");
    s.insert("customXmlMoveToRangeEnd");
    s.insert("moveFromRangeStart");
    s.insert("moveFromRangeEnd");
    s.insert("moveToRangeStart");
    s.insert("moveToRangeEnd");
    s
});

static ELEMENTS_TO_UNWRAP: Lazy<HashSet<&'static str>> = Lazy::new(|| {
    let mut s = HashSet::new();
    s.insert("ins");
    s.insert("moveTo");
    s
});

pub fn accept_revisions(source: &XmlDocument, source_root: NodeId) -> XmlDocument {
    let mut result = XmlDocument::new();

    if let Some(children) = transform_node(source, source_root, &mut result, None) {
        if children.len() == 1 {
            result.set_root(Some(children[0]));
        }
    }

    result
}

fn transform_node(
    source: &XmlDocument,
    node_id: NodeId,
    result: &mut XmlDocument,
    parent: Option<NodeId>,
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

            if ns == Some(W::NS) && ELEMENTS_TO_REMOVE.contains(local.as_str()) {
                return None;
            }

            if ns == Some(W::NS) && ELEMENTS_TO_UNWRAP.contains(local.as_str()) {
                let mut unwrapped = Vec::new();
                for child in source.children(node_id) {
                    if let Some(children) = transform_node(source, child, result, parent) {
                        unwrapped.extend(children);
                    }
                }
                return if unwrapped.is_empty() {
                    None
                } else {
                    Some(unwrapped)
                };
            }

            if ns == Some(W::NS) && local == "tr" && is_deleted_table_row(source, node_id) {
                return None;
            }

            if ns == Some(M::NS) && local == "f" && has_deleted_math_control(source, node_id) {
                return None;
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
                transform_node(source, child, result, Some(new_id));
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
