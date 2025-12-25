use super::node::XmlNodeData;
use super::xname::{XAttribute, XName};
use crate::error::{RedlineError, Result};
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
}
