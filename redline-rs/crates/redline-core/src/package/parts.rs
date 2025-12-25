use crate::xml::XmlDocument;

pub struct Part {
    pub path: String,
    pub content_type: String,
    pub content: PartContent,
}

pub enum PartContent {
    Xml(XmlDocument),
    Binary(Vec<u8>),
}

impl Part {
    pub fn xml(path: &str, content_type: &str, doc: XmlDocument) -> Self {
        Self {
            path: path.to_string(),
            content_type: content_type.to_string(),
            content: PartContent::Xml(doc),
        }
    }

    pub fn binary(path: &str, content_type: &str, data: Vec<u8>) -> Self {
        Self {
            path: path.to_string(),
            content_type: content_type.to_string(),
            content: PartContent::Binary(data),
        }
    }

    pub fn is_xml(&self) -> bool {
        matches!(self.content, PartContent::Xml(_))
    }

    pub fn as_xml(&self) -> Option<&XmlDocument> {
        match &self.content {
            PartContent::Xml(doc) => Some(doc),
            _ => None,
        }
    }

    pub fn as_xml_mut(&mut self) -> Option<&mut XmlDocument> {
        match &mut self.content {
            PartContent::Xml(doc) => Some(doc),
            _ => None,
        }
    }
}
