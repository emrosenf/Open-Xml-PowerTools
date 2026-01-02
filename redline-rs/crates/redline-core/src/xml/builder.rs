use super::arena::XmlDocument;
use super::node::XmlNodeData;
use super::xname::{XAttribute, XName};
use crate::error::Result;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;
use std::collections::{HashMap, HashSet};
use std::io::Cursor;

pub fn serialize(doc: &XmlDocument) -> Result<String> {
    let bytes = serialize_bytes(doc)?;
    Ok(String::from_utf8(bytes).expect("XML should be valid UTF-8"))
}

/// Serialize a subtree starting from a specific node (no XML declaration)
pub fn serialize_subtree(doc: &XmlDocument, node_id: indextree::NodeId) -> Result<String> {
    use std::io::Cursor;

    let mut writer = Writer::new(Cursor::new(Vec::new()));

    let Some(node_data) = doc.get(node_id) else {
        return Ok(String::new());
    };

    match node_data {
        XmlNodeData::Element { name, attributes } => {
            let mut merged_attrs = attributes.clone();
            let mut declared: HashSet<XName> = merged_attrs
                .iter()
                .filter(|attr| is_xmlns_attr(attr))
                .map(|attr| attr.name.clone())
                .collect();

            let inherited = collect_ancestor_namespace_attrs(doc, node_id, &mut declared);
            merged_attrs.extend(inherited);

            let mut namespace_map = NamespaceMap::new();
            extend_namespace_map(&mut namespace_map, &merged_attrs);

            write_element_with_attrs(doc, node_id, name, &merged_attrs, &mut writer, &namespace_map)?;
        }
        _ => {
            let namespace_map = NamespaceMap::new();
            write_node(doc, node_id, &mut writer, &namespace_map)?;
        }
    }

    let bytes = writer.into_inner().into_inner();
    Ok(String::from_utf8(bytes).expect("XML should be valid UTF-8"))
}

pub fn serialize_bytes(doc: &XmlDocument) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    
    writer
        .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .map_err(|e| crate::error::RedlineError::XmlWrite(e.to_string()))?;

    if let Some(root_id) = doc.root() {
        let mut namespace_map = NamespaceMap::new();
        if let Some(root_data) = doc.get(root_id) {
            if let Some(attrs) = root_data.attributes() {
                extend_namespace_map(&mut namespace_map, attrs);
            }
        }
        write_node(doc, root_id, &mut writer, &namespace_map)?;
    }

    Ok(writer.into_inner().into_inner())
}

type NamespaceMap = HashMap<String, String>;

fn is_xmlns_attr(attr: &XAttribute) -> bool {
    (attr.name.namespace.is_none() && attr.name.local_name == "xmlns")
        || attr.name.namespace.as_deref() == Some("http://www.w3.org/2000/xmlns/")
}

fn collect_ancestor_namespace_attrs(
    doc: &XmlDocument,
    node_id: indextree::NodeId,
    declared: &mut HashSet<XName>,
) -> Vec<XAttribute> {
    let mut collected = Vec::new();
    let mut ancestors = doc.ancestors(node_id);
    ancestors.next(); // skip the node itself

    for ancestor_id in ancestors {
        let Some(data) = doc.get(ancestor_id) else {
            continue;
        };
        let Some(attrs) = data.attributes() else {
            continue;
        };

        for attr in attrs {
            if is_xmlns_attr(attr) && !declared.contains(&attr.name) {
                declared.insert(attr.name.clone());
                collected.push(attr.clone());
            }
        }
    }

    collected
}

fn extend_namespace_map(namespace_map: &mut NamespaceMap, attributes: &[super::xname::XAttribute]) {
    for attr in attributes {
        let Some(ns) = &attr.name.namespace else {
            if attr.name.local_name == "xmlns" {
                // Default namespace declaration.
                namespace_map.entry(attr.value.clone()).or_insert_with(String::new);
            }
            continue;
        };

        if ns == "http://www.w3.org/2000/xmlns/" {
            // Namespace declaration: xmlns:prefix="uri".
            namespace_map
                .entry(attr.value.clone())
                .or_insert_with(|| attr.name.local_name.clone());
        }
    }
}

fn prefix_for_namespace<'a>(namespace: &str, namespace_map: &'a NamespaceMap) -> &'a str {
    if let Some(prefix) = namespace_map.get(namespace) {
        return prefix.as_str();
    }

    get_prefix(namespace)
}

fn prefix_for_attribute<'a>(namespace: &str, namespace_map: &'a NamespaceMap) -> &'a str {
    if namespace == "http://www.w3.org/2000/xmlns/" {
        return "xmlns";
    }

    if let Some(prefix) = namespace_map.get(namespace) {
        if !prefix.is_empty() {
            return prefix.as_str();
        }
    }

    get_prefix(namespace)
}

fn write_node<W: std::io::Write>(
    doc: &XmlDocument,
    node_id: indextree::NodeId,
    writer: &mut Writer<W>,
    namespace_map: &NamespaceMap,
) -> Result<()> {
    let Some(node_data) = doc.get(node_id) else {
        return Ok(());
    };

    match node_data {
        XmlNodeData::Element { name, attributes } => {
            write_element_with_attrs(doc, node_id, name, attributes, writer, namespace_map)?;
        }
        XmlNodeData::Text(text) => {
            writer
                .write_event(Event::Text(BytesText::new(text)))
                .map_err(|e| crate::error::RedlineError::XmlWrite(e.to_string()))?;
        }
        XmlNodeData::CData(text) => {
            writer
                .write_event(Event::CData(quick_xml::events::BytesCData::new(text)))
                .map_err(|e| crate::error::RedlineError::XmlWrite(e.to_string()))?;
        }
        XmlNodeData::Comment(text) => {
            writer
                .write_event(Event::Comment(BytesText::new(text)))
                .map_err(|e| crate::error::RedlineError::XmlWrite(e.to_string()))?;
        }
        XmlNodeData::ProcessingInstruction { target, data } => {
            let pi_content = if data.is_empty() {
                target.clone()
            } else {
                format!("{} {}", target, data)
            };
            writer
                .write_event(Event::PI(quick_xml::events::BytesPI::new(&pi_content)))
                .map_err(|e| crate::error::RedlineError::XmlWrite(e.to_string()))?;
        }
    }

    Ok(())
}

fn write_element_with_attrs<W: std::io::Write>(
    doc: &XmlDocument,
    node_id: indextree::NodeId,
    name: &XName,
    attributes: &[XAttribute],
    writer: &mut Writer<W>,
    namespace_map: &NamespaceMap,
) -> Result<()> {
    let mut scoped_map = namespace_map.clone();
    extend_namespace_map(&mut scoped_map, attributes);

    let tag_name = if let Some(ns) = &name.namespace {
        let prefix = prefix_for_namespace(ns, &scoped_map);
        if prefix.is_empty() {
            name.local_name.clone()
        } else {
            format!("{}:{}", prefix, &name.local_name)
        }
    } else {
        name.local_name.clone()
    };

    let mut elem = BytesStart::new(&tag_name);

    for attr in attributes {
        let attr_name = if let Some(ns) = &attr.name.namespace {
            let prefix = prefix_for_attribute(ns, &scoped_map);
            if prefix.is_empty() {
                attr.name.local_name.clone()
            } else {
                format!("{}:{}", prefix, &attr.name.local_name)
            }
        } else {
            attr.name.local_name.clone()
        };
        elem.push_attribute((attr_name.as_str(), attr.value.as_str()));
    }

    let children: Vec<_> = doc.children(node_id).collect();

    if children.is_empty() {
        writer
            .write_event(Event::Empty(elem))
            .map_err(|e| crate::error::RedlineError::XmlWrite(e.to_string()))?;
    } else {
        writer
            .write_event(Event::Start(elem))
            .map_err(|e| crate::error::RedlineError::XmlWrite(e.to_string()))?;

        for child_id in children {
            write_node(doc, child_id, writer, &scoped_map)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new(&tag_name)))
            .map_err(|e| crate::error::RedlineError::XmlWrite(e.to_string()))?;
    }

    Ok(())
}

fn get_prefix(namespace: &str) -> &'static str {
    match namespace {
        // WordprocessingML
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main" => "w",
        "http://schemas.microsoft.com/office/word/2010/wordml" => "w14",
        "http://schemas.microsoft.com/office/word/2012/wordml" => "w15",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" => "wps",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" => "wpg",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" => "wp14",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" => "wpc",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingInk" => "wpi",
        "http://schemas.microsoft.com/office/word/2023/wordml/word16du" => "w16du",
        // SpreadsheetML
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main" => "x",
        // PresentationML
        "http://schemas.openxmlformats.org/presentationml/2006/main" => "p",
        // DrawingML
        "http://schemas.openxmlformats.org/drawingml/2006/main" => "a",
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" => "wp",
        "http://schemas.openxmlformats.org/drawingml/2006/picture" => "pic",
        "http://schemas.openxmlformats.org/drawingml/2006/chart" => "c",
        // Office Math
        "http://schemas.openxmlformats.org/officeDocument/2006/math" => "m",
        // Relationships
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships" => "r",
        // Markup Compatibility
        "http://schemas.openxmlformats.org/markup-compatibility/2006" => "mc",
        // PowerTools
        "http://powertools.codeplex.com/2011" => "pt",
        // xmlns namespace for namespace declarations (xmlns:mc="...", etc.)
        "http://www.w3.org/2000/xmlns/" => "xmlns",
        // xml namespace for xml:space, xml:lang, etc.
        "http://www.w3.org/XML/1998/namespace" => "xml",
        _ => "ns",
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::xname::XName;

    #[test]
    fn serialize_simple_document() {
        let mut doc = XmlDocument::new();
        let root = doc.add_root(XmlNodeData::element(XName::local("root")));
        doc.add_child(root, XmlNodeData::text("content"));
        
        let xml = serialize(&doc).unwrap();
        assert!(xml.contains("<root>content</root>"));
    }

    #[test]
    fn serialize_empty_element() {
        let mut doc = XmlDocument::new();
        doc.add_root(XmlNodeData::element(XName::local("empty")));
        
        let xml = serialize(&doc).unwrap();
        assert!(xml.contains("<empty/>"));
    }
}
