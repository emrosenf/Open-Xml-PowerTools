use crate::xml::arena::XmlDocument;
use crate::xml::node::XmlNodeData;
use indextree::NodeId;

pub fn descendants_trimmed<'a, F>(
    doc: &'a XmlDocument,
    node: NodeId,
    trim_predicate: F,
) -> impl Iterator<Item = NodeId> + 'a
where
    F: Fn(&XmlNodeData) -> bool + 'a,
{
    DescendantsTrimmedIter::new(doc, node, trim_predicate)
}

struct DescendantsTrimmedIter<'a, F>
where
    F: Fn(&XmlNodeData) -> bool,
{
    doc: &'a XmlDocument,
    stack: Vec<NodeId>,
    trim_predicate: F,
}

impl<'a, F> DescendantsTrimmedIter<'a, F>
where
    F: Fn(&XmlNodeData) -> bool,
{
    fn new(doc: &'a XmlDocument, start: NodeId, trim_predicate: F) -> Self {
        Self {
            doc,
            stack: vec![start],
            trim_predicate,
        }
    }
}

impl<'a, F> Iterator for DescendantsTrimmedIter<'a, F>
where
    F: Fn(&XmlNodeData) -> bool,
{
    type Item = NodeId;

    fn next(&mut self) -> Option<Self::Item> {
        while let Some(current) = self.stack.pop() {
            if let Some(data) = self.doc.get(current) {
                if (self.trim_predicate)(data) {
                    continue;
                }
                
                let children: Vec<_> = self.doc.children(current).collect();
                for child in children.into_iter().rev() {
                    self.stack.push(child);
                }
                
                return Some(current);
            }
        }
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::xname::XName;

    #[test]
    fn descendants_trimmed_skips_matching_nodes() {
        let mut doc = XmlDocument::new();
        let root = doc.add_root(XmlNodeData::element(XName::local("root")));
        let child1 = doc.add_child(root, XmlNodeData::element(XName::local("skip")));
        let _grandchild = doc.add_child(child1, XmlNodeData::element(XName::local("hidden")));
        let child2 = doc.add_child(root, XmlNodeData::element(XName::local("keep")));
        
        let result: Vec<_> = descendants_trimmed(&doc, root, |data| {
            data.name().map(|n| n.local_name == "skip").unwrap_or(false)
        }).collect();
        
        assert_eq!(result.len(), 2);
        assert!(result.contains(&root));
        assert!(result.contains(&child2));
    }
}
