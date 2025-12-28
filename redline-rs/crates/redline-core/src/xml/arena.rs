use super::node::XmlNodeData;
use super::xname::{XAttribute, XName};
use indextree::{Arena, NodeId};

pub struct XmlDocument {
    arena: Arena<XmlNodeData>,
    root: Option<NodeId>,
}

impl XmlDocument {
    pub fn new() -> Self {
        Self {
            arena: Arena::new(),
            root: None,
        }
    }

    pub fn root(&self) -> Option<NodeId> {
        self.root
    }

    pub fn set_root(&mut self, root: Option<NodeId>) {
        self.root = root;
    }

    pub fn get(&self, id: NodeId) -> Option<&XmlNodeData> {
        self.arena.get(id).map(|node| node.get())
    }

    pub fn get_mut(&mut self, id: NodeId) -> Option<&mut XmlNodeData> {
        self.arena.get_mut(id).map(|node| node.get_mut())
    }

    pub fn add_root(&mut self, data: XmlNodeData) -> NodeId {
        let id = self.arena.new_node(data);
        self.root = Some(id);
        id
    }

    pub fn new_node(&mut self, data: XmlNodeData) -> NodeId {
        self.arena.new_node(data)
    }

    pub fn add_child(&mut self, parent: NodeId, data: XmlNodeData) -> NodeId {
        let child = self.arena.new_node(data);
        parent.append(child, &mut self.arena);
        child
    }

    pub fn add_before(&mut self, sibling: NodeId, data: XmlNodeData) -> NodeId {
        let new_node = self.arena.new_node(data);
        sibling.insert_before(new_node, &mut self.arena);
        new_node
    }

    pub fn add_after(&mut self, sibling: NodeId, data: XmlNodeData) -> NodeId {
        let new_node = self.arena.new_node(data);
        sibling.insert_after(new_node, &mut self.arena);
        new_node
    }

    pub fn remove(&mut self, node: NodeId) {
        node.remove(&mut self.arena);
    }

    pub fn set_attribute(&mut self, node: NodeId, name: &XName, value: &str) {
        if let Some(node_data) = self.get_mut(node) {
            if let Some(attrs) = node_data.attributes_mut() {
                if let Some(attr) = attrs.iter_mut().find(|a| &a.name == name) {
                    attr.value = value.to_string();
                } else {
                    attrs.push(XAttribute::new(name.clone(), value));
                }
            }
        }
    }

    pub fn remove_attribute(&mut self, node: NodeId, name: &XName) {
        if let Some(node_data) = self.get_mut(node) {
            if let Some(attrs) = node_data.attributes_mut() {
                attrs.retain(|a| &a.name != name);
            }
        }
    }

    pub fn children(&self, parent: NodeId) -> impl Iterator<Item = NodeId> + '_ {
        parent.children(&self.arena)
    }

    pub fn descendants(&self, node: NodeId) -> impl Iterator<Item = NodeId> + '_ {
        node.descendants(&self.arena)
    }

    pub fn parent(&self, node: NodeId) -> Option<NodeId> {
        self.arena.get(node)?.parent()
    }

    pub fn ancestors(&self, node: NodeId) -> impl Iterator<Item = NodeId> + '_ {
        node.ancestors(&self.arena)
    }

    /// Detach a node from its current parent (but keep its children)
    pub fn detach(&mut self, node: NodeId) {
        node.detach(&mut self.arena);
    }

    /// Insert an existing node before a sibling (detaching it from any previous parent)
    pub fn insert_before(&mut self, sibling: NodeId, node: NodeId) {
        node.detach(&mut self.arena);
        sibling.insert_before(node, &mut self.arena);
    }

    /// Insert an existing node after a sibling (detaching it from any previous parent)
    pub fn insert_after(&mut self, sibling: NodeId, node: NodeId) {
        node.detach(&mut self.arena);
        sibling.insert_after(node, &mut self.arena);
    }

    /// Append a child node to a parent (detaching it from any previous parent)
    pub fn reparent(&mut self, parent: NodeId, child: NodeId) {
        child.detach(&mut self.arena);
        parent.append(child, &mut self.arena);
    }

    pub fn elements_by_name<'a>(
        &'a self,
        parent: NodeId,
        name: &'a XName,
    ) -> impl Iterator<Item = NodeId> + 'a {
        self.children(parent).filter(move |&child_id| {
            self.get(child_id)
                .and_then(|data| data.name())
                .map(|n| n == name)
                .unwrap_or(false)
        })
    }

    /// Find the first child element with the given name
    pub fn find_child(&self, parent: NodeId, name: &XName) -> Option<NodeId> {
        self.children(parent).find(|&child_id| {
            self.get(child_id)
                .and_then(|data| data.name())
                .map(|n| n == name)
                .unwrap_or(false)
        })
    }

    /// Get the name of a node (if it's an element)
    pub fn name(&self, node: NodeId) -> Option<&XName> {
        self.get(node).and_then(|data| data.name())
    }

    /// Get an attribute value as a String
    pub fn get_attribute_string(&self, node: NodeId, name: &XName) -> Option<String> {
        self.get(node)
            .and_then(|data| data.attributes())
            .and_then(|attrs| {
                attrs.iter()
                    .find(|attr| &attr.name == name)
                    .map(|attr| attr.value.clone())
            })
    }

    /// Get an attribute value as an i64
    pub fn get_attribute_i64(&self, node: NodeId, name: &str) -> Option<i64> {
        let attr_name = XName::local(name);
        self.get_attribute_string(node, &attr_name)
            .and_then(|s| s.parse::<i64>().ok())
    }

    /// Get an attribute value as an i32
    pub fn get_attribute_i32(&self, node: NodeId, name: &str) -> Option<i32> {
        let attr_name = XName::local(name);
        self.get_attribute_string(node, &attr_name)
            .and_then(|s| s.parse::<i32>().ok())
    }

    /// Get an attribute value as a bool
    pub fn get_attribute_bool(&self, node: NodeId, name: &str) -> Option<bool> {
        let attr_name = XName::local(name);
        self.get_attribute_string(node, &attr_name)
            .map(|s| s == "1" || s == "true")
    }

    /// Get the text content of a text node
    pub fn text(&self, node: NodeId) -> Option<String> {
        self.get(node)
            .and_then(|data| data.text_content())
            .map(|s| s.to_string())
    }

    /// Convert a node and its descendants to an XML string
    pub fn to_xml_string(&self, node: NodeId) -> String {
        let mut result = String::new();
        self.node_to_xml(node, &mut result);
        result
    }

    fn node_to_xml(&self, node: NodeId, output: &mut String) {
        let Some(data) = self.get(node) else { return };

        match data {
            XmlNodeData::Element { name, attributes } => {
                // Opening tag
                output.push('<');
                if let Some(ns) = &name.namespace {
                    output.push_str(ns);
                    output.push(':');
                }
                output.push_str(&name.local_name);

                // Attributes
                for attr in attributes {
                    output.push(' ');
                    if let Some(ns) = &attr.name.namespace {
                        output.push_str(ns);
                        output.push(':');
                    }
                    output.push_str(&attr.name.local_name);
                    output.push_str("=\"");
                    Self::escape_attribute(&attr.value, output);
                    output.push('"');
                }

                // Check if element has children
                let has_children = self.children(node).next().is_some();
                
                if has_children {
                    output.push('>');
                    
                    // Process children
                    for child in self.children(node) {
                        self.node_to_xml(child, output);
                    }
                    
                    // Closing tag
                    output.push_str("</");
                    if let Some(ns) = &name.namespace {
                        output.push_str(ns);
                        output.push(':');
                    }
                    output.push_str(&name.local_name);
                    output.push('>');
                } else {
                    output.push_str("/>");
                }
            }
            XmlNodeData::Text(text) => {
                Self::escape_text(text, output);
            }
            XmlNodeData::CData(text) => {
                output.push_str("<![CDATA[");
                output.push_str(text);
                output.push_str("]]>");
            }
            XmlNodeData::Comment(text) => {
                output.push_str("<!--");
                output.push_str(text);
                output.push_str("-->");
            }
            XmlNodeData::ProcessingInstruction { target, data } => {
                output.push_str("<?");
                output.push_str(target);
                output.push(' ');
                output.push_str(data);
                output.push_str("?>");
            }
        }
    }

    fn escape_text(text: &str, output: &mut String) {
        for c in text.chars() {
            match c {
                '<' => output.push_str("&lt;"),
                '>' => output.push_str("&gt;"),
                '&' => output.push_str("&amp;"),
                _ => output.push(c),
            }
        }
    }

    fn escape_attribute(text: &str, output: &mut String) {
        for c in text.chars() {
            match c {
                '<' => output.push_str("&lt;"),
                '>' => output.push_str("&gt;"),
                '&' => output.push_str("&amp;"),
                '"' => output.push_str("&quot;"),
                _ => output.push(c),
            }
        }
    }
}

impl Default for XmlDocument {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn create_document_with_root() {
        let mut doc = XmlDocument::new();
        let root_name = XName::new("http://example.com", "root");
        let root_id = doc.add_root(XmlNodeData::element(root_name.clone()));
        
        assert!(doc.root().is_some());
        assert_eq!(doc.root(), Some(root_id));
        
        let data = doc.get(root_id).unwrap();
        assert_eq!(data.name(), Some(&root_name));
    }

    #[test]
    fn add_children_to_element() {
        let mut doc = XmlDocument::new();
        let root_id = doc.add_root(XmlNodeData::element(XName::local("root")));
        
        let child1 = doc.add_child(root_id, XmlNodeData::element(XName::local("child1")));
        let child2 = doc.add_child(root_id, XmlNodeData::element(XName::local("child2")));
        
        let children: Vec<_> = doc.children(root_id).collect();
        assert_eq!(children.len(), 2);
        assert!(children.contains(&child1));
        assert!(children.contains(&child2));
    }

    #[test]
    fn set_and_get_attribute() {
        let mut doc = XmlDocument::new();
        let root_id = doc.add_root(XmlNodeData::element(XName::local("root")));
        
        let attr_name = XName::local("id");
        doc.set_attribute(root_id, &attr_name, "test123");
        
        let data = doc.get(root_id).unwrap();
        let attrs = data.attributes().unwrap();
        assert_eq!(attrs.len(), 1);
        assert_eq!(attrs[0].name, attr_name);
        assert_eq!(attrs[0].value, "test123");
    }

    #[test]
    fn find_child_by_name() {
        let mut doc = XmlDocument::new();
        let root_id = doc.add_root(XmlNodeData::element(XName::local("root")));
        
        let child1 = doc.add_child(root_id, XmlNodeData::element(XName::local("child1")));
        let child2 = doc.add_child(root_id, XmlNodeData::element(XName::local("target")));
        let _child3 = doc.add_child(root_id, XmlNodeData::element(XName::local("child3")));
        
        let target_name = XName::local("target");
        let found = doc.find_child(root_id, &target_name);
        
        assert_eq!(found, Some(child2));
    }

    #[test]
    fn get_node_name() {
        let mut doc = XmlDocument::new();
        let root_name = XName::local("root");
        let root_id = doc.add_root(XmlNodeData::element(root_name.clone()));
        
        assert_eq!(doc.name(root_id), Some(&root_name));
        
        let text_id = doc.add_child(root_id, XmlNodeData::text("hello"));
        assert_eq!(doc.name(text_id), None);
    }

    #[test]
    fn get_attribute_string_method() {
        let mut doc = XmlDocument::new();
        let root_id = doc.add_root(XmlNodeData::element(XName::local("root")));
        
        let attr_name = XName::local("id");
        doc.set_attribute(root_id, &attr_name, "test123");
        
        assert_eq!(doc.get_attribute_string(root_id, &attr_name), Some("test123".to_string()));
        
        let missing_name = XName::local("missing");
        assert_eq!(doc.get_attribute_string(root_id, &missing_name), None);
    }

    #[test]
    fn get_attribute_i64_method() {
        let mut doc = XmlDocument::new();
        let root_id = doc.add_root(XmlNodeData::element(XName::local("root")));
        
        let attr_name = XName::local("width");
        doc.set_attribute(root_id, &attr_name, "12345");
        
        assert_eq!(doc.get_attribute_i64(root_id, "width"), Some(12345));
        assert_eq!(doc.get_attribute_i64(root_id, "height"), None);
    }

    #[test]
    fn get_text_content() {
        let mut doc = XmlDocument::new();
        let root_id = doc.add_root(XmlNodeData::element(XName::local("root")));
        
        let text_id = doc.add_child(root_id, XmlNodeData::text("Hello, World!"));
        
        assert_eq!(doc.text(text_id), Some("Hello, World!".to_string()));
        assert_eq!(doc.text(root_id), None);
    }

    #[test]
    fn to_xml_string_simple() {
        let mut doc = XmlDocument::new();
        let root_id = doc.add_root(XmlNodeData::element(XName::local("root")));
        
        let xml = doc.to_xml_string(root_id);
        assert_eq!(xml, "<root/>");
    }

    #[test]
    fn to_xml_string_with_attributes() {
        let mut doc = XmlDocument::new();
        let root_id = doc.add_root(XmlNodeData::element(XName::local("root")));
        doc.set_attribute(root_id, &XName::local("id"), "test");
        doc.set_attribute(root_id, &XName::local("name"), "example");
        
        let xml = doc.to_xml_string(root_id);
        // Note: attribute order might vary, so we just check it contains the parts
        assert!(xml.contains("<root"));
        assert!(xml.contains("id=\"test\""));
        assert!(xml.contains("name=\"example\""));
        assert!(xml.ends_with("/>"));
    }

    #[test]
    fn to_xml_string_with_children() {
        let mut doc = XmlDocument::new();
        let root_id = doc.add_root(XmlNodeData::element(XName::local("root")));
        let child_id = doc.add_child(root_id, XmlNodeData::element(XName::local("child")));
        doc.add_child(child_id, XmlNodeData::text("Hello"));
        
        let xml = doc.to_xml_string(root_id);
        assert_eq!(xml, "<root><child>Hello</child></root>");
    }

    #[test]
    fn to_xml_string_escapes_text() {
        let mut doc = XmlDocument::new();
        let root_id = doc.add_root(XmlNodeData::element(XName::local("root")));
        doc.add_child(root_id, XmlNodeData::text("<>&"));
        
        let xml = doc.to_xml_string(root_id);
        assert_eq!(xml, "<root>&lt;&gt;&amp;</root>");
    }

    #[test]
    fn to_xml_string_escapes_attributes() {
        let mut doc = XmlDocument::new();
        let root_id = doc.add_root(XmlNodeData::element(XName::local("root")));
        doc.set_attribute(root_id, &XName::local("val"), "<>&\"");
        
        let xml = doc.to_xml_string(root_id);
        assert!(xml.contains("val=\"&lt;&gt;&amp;&quot;\""));
    }
}
