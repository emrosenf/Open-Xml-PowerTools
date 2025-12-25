use super::xname::{XAttribute, XName};
use indextree::NodeId;

#[derive(Clone, Debug)]
pub enum XmlNodeData {
    Element {
        name: XName,
        attributes: Vec<XAttribute>,
    },
    Text(String),
    CData(String),
    Comment(String),
    ProcessingInstruction { target: String, data: String },
}

impl XmlNodeData {
    pub fn element(name: XName) -> Self {
        Self::Element {
            name,
            attributes: Vec::new(),
        }
    }

    pub fn element_with_attrs(name: XName, attributes: Vec<XAttribute>) -> Self {
        Self::Element { name, attributes }
    }

    pub fn text(content: &str) -> Self {
        Self::Text(content.to_string())
    }

    pub fn is_element(&self) -> bool {
        matches!(self, Self::Element { .. })
    }

    pub fn is_text(&self) -> bool {
        matches!(self, Self::Text(_))
    }

    pub fn name(&self) -> Option<&XName> {
        match self {
            Self::Element { name, .. } => Some(name),
            _ => None,
        }
    }

    pub fn attributes(&self) -> Option<&[XAttribute]> {
        match self {
            Self::Element { attributes, .. } => Some(attributes),
            _ => None,
        }
    }

    pub fn attributes_mut(&mut self) -> Option<&mut Vec<XAttribute>> {
        match self {
            Self::Element { attributes, .. } => Some(attributes),
            _ => None,
        }
    }

    pub fn text_content(&self) -> Option<&str> {
        match self {
            Self::Text(s) | Self::CData(s) => Some(s),
            _ => None,
        }
    }
}

pub struct XmlNode<'a> {
    pub id: NodeId,
    pub data: &'a XmlNodeData,
}

impl<'a> XmlNode<'a> {
    pub fn new(id: NodeId, data: &'a XmlNodeData) -> Self {
        Self { id, data }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn element_node_creation() {
        let name = XName::new("http://example.com", "test");
        let node = XmlNodeData::element(name.clone());
        assert!(node.is_element());
        assert_eq!(node.name(), Some(&name));
    }

    #[test]
    fn text_node_creation() {
        let node = XmlNodeData::text("Hello, World!");
        assert!(node.is_text());
        assert_eq!(node.text_content(), Some("Hello, World!"));
    }
}
