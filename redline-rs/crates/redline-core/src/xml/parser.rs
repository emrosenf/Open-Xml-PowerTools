use super::arena::XmlDocument;
use super::node::XmlNodeData;
use super::xname::{XAttribute, XName};
use crate::error::{RedlineError, Result};

pub fn parse(xml: &str) -> Result<XmlDocument> {
    parse_bytes(xml.as_bytes())
}

pub fn parse_bytes(bytes: &[u8]) -> Result<XmlDocument> {
    let doc = roxmltree::Document::parse_with_options(
        std::str::from_utf8(bytes).map_err(|e| RedlineError::XmlParse {
            message: e.to_string(),
            location: "input".to_string(),
        })?,
        roxmltree::ParsingOptions {
            allow_dtd: true,
            ..Default::default()
        },
    )
    .map_err(|e| RedlineError::XmlParse {
        message: e.to_string(),
        location: format!("line {}", e.pos().row),
    })?;

    let mut xml_doc = XmlDocument::new();
    
    if doc.root_element().parent().is_some() {
        build_tree(doc.root_element(), &mut xml_doc, None);
    }

    Ok(xml_doc)
}

fn build_tree(
    node: roxmltree::Node,
    doc: &mut XmlDocument,
    parent: Option<indextree::NodeId>,
) {
    let node_data = match node.node_type() {
        roxmltree::NodeType::Element => {
            let name = XName::new(
                node.tag_name().namespace().unwrap_or(""),
                node.tag_name().name(),
            );
            
            let mut attributes: Vec<XAttribute> = node
                .attributes()
                .map(|attr| {
                    XAttribute::new(
                        XName::new(attr.namespace().unwrap_or(""), attr.name()),
                        attr.value(),
                    )
                })
                .collect();
            
            // Capture namespace declarations as attributes (xmlns:prefix="uri")
            // roxmltree separates these from regular attributes
            for ns in node.namespaces() {
                if let Some(prefix) = ns.name() {
                    // This is xmlns:prefix="uri"
                    attributes.push(XAttribute::new(
                        XName::new("http://www.w3.org/2000/xmlns/", prefix),
                        ns.uri(),
                    ));
                } else {
                    // This is xmlns="uri" (default namespace)
                    attributes.push(XAttribute::new(
                        XName::local("xmlns"),
                        ns.uri(),
                    ));
                }
            }
            
            XmlNodeData::Element { name, attributes }
        }
        roxmltree::NodeType::Text => {
            if let Some(text) = node.text() {
                XmlNodeData::Text(text.to_string())
            } else {
                return;
            }
        }
        roxmltree::NodeType::Comment => {
            if let Some(text) = node.text() {
                XmlNodeData::Comment(text.to_string())
            } else {
                return;
            }
        }
        roxmltree::NodeType::PI => {
            XmlNodeData::ProcessingInstruction {
                target: node.pi()
                    .map(|pi| pi.target.to_string())
                    .unwrap_or_default(),
                data: node.pi()
                    .and_then(|pi| pi.value.map(|s| s.to_string()))
                    .unwrap_or_default(),
            }
        }
        _ => return,
    };

    let new_id = match parent {
        Some(parent_id) => doc.add_child(parent_id, node_data),
        None => doc.add_root(node_data),
    };

    for child in node.children() {
        build_tree(child, doc, Some(new_id));
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_simple_xml() {
        let xml = r#"<root><child attr="value">text</child></root>"#;
        let doc = parse(xml).unwrap();
        
        assert!(doc.root().is_some());
    }

    #[test]
    fn parse_xml_with_namespaces() {
        let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>
        </w:document>"#;
        
        let doc = parse(xml).unwrap();
        assert!(doc.root().is_some());
    }

    #[test]
    fn parse_preserves_attribute_order() {
        let xml = r#"<root a="1" b="2" c="3" d="4"/>"#;
        let doc = parse(xml).unwrap();
        
        let root_id = doc.root().unwrap();
        let root_data = doc.get(root_id).unwrap();
        let attrs = root_data.attributes().unwrap();
        
        assert_eq!(attrs.len(), 4);
        assert_eq!(attrs[0].name.local_name, "a");
        assert_eq!(attrs[1].name.local_name, "b");
        assert_eq!(attrs[2].name.local_name, "c");
        assert_eq!(attrs[3].name.local_name, "d");
    }
}
