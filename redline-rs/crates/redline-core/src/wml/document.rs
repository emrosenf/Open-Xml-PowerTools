use crate::error::Result;
use crate::package::OoxmlPackage;
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{MC, M, W};
use crate::xml::node::XmlNodeData;
use indextree::NodeId;

pub struct WmlDocument {
    package: OoxmlPackage,
}

impl WmlDocument {
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let package = OoxmlPackage::open(bytes)?;
        Ok(Self { package })
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.package.save()
    }

    pub fn package(&self) -> &OoxmlPackage {
        &self.package
    }

    pub fn package_mut(&mut self) -> &mut OoxmlPackage {
        &mut self.package
    }

    pub fn main_document(&self) -> Result<XmlDocument> {
        self.package.get_xml_part("word/document.xml")
    }

    pub fn footnotes(&self) -> Result<Option<XmlDocument>> {
        match self.package.get_part("word/footnotes.xml") {
            Some(_) => Ok(Some(self.package.get_xml_part("word/footnotes.xml")?)),
            None => Ok(None),
        }
    }

    pub fn endnotes(&self) -> Result<Option<XmlDocument>> {
        match self.package.get_part("word/endnotes.xml") {
            Some(_) => Ok(Some(self.package.get_xml_part("word/endnotes.xml")?)),
            None => Ok(None),
        }
    }

    pub fn styles(&self) -> Result<Option<XmlDocument>> {
        match self.package.get_part("word/styles.xml") {
            Some(_) => Ok(Some(self.package.get_xml_part("word/styles.xml")?)),
            None => Ok(None),
        }
    }
}

pub fn find_document_body(doc: &XmlDocument) -> Option<NodeId> {
    let root = doc.root()?;
    
    for child in doc.descendants(root) {
        if let Some(data) = doc.get(child) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "body" {
                    return Some(child);
                }
            }
        }
    }
    
    None
}

pub fn extract_paragraph_text(doc: &XmlDocument, para: NodeId) -> String {
    let mut texts = Vec::new();
    extract_text_recursive(doc, para, &mut texts, false);
    texts.join("")
}

fn extract_text_recursive(
    doc: &XmlDocument,
    node: NodeId,
    texts: &mut Vec<String>,
    accept_revisions: bool,
) {
    let Some(data) = doc.get(node) else { return };
    
    match data {
        XmlNodeData::Text(text) => {
            texts.push(text.clone());
        }
        XmlNodeData::Element { name, .. } => {
            let ns = name.namespace.as_deref();
            let local = name.local_name.as_str();
            
            if ns == Some(W::NS) && local == "del" && accept_revisions {
                return;
            }
            
            if ns == Some(W::NS) && local == "delText" {
                return;
            }
            
            if ns == Some(W::NS) && local == "t" {
                for child in doc.children(node) {
                    if let Some(XmlNodeData::Text(text)) = doc.get(child) {
                        texts.push(text.clone());
                    }
                }
                return;
            }
            
            if ns == Some(W::NS) && local == "br" {
                texts.push("\n".to_string());
                return;
            }
            
            if ns == Some(W::NS) && local == "tab" {
                texts.push("\t".to_string());
                return;
            }
            
            if ns == Some(W::NS) && local == "footnoteReference" {
                if let Some(attrs) = data.attributes() {
                    let id = attrs.iter()
                        .find(|a| a.name.local_name == "id" && a.name.namespace.as_deref() == Some(W::NS))
                        .map(|a| a.value.as_str())
                        .unwrap_or("unknown");
                    texts.push(format!(" FOOTNOTE_REF_{} ", id));
                }
                return;
            }
            
            if ns == Some(W::NS) && local == "endnoteReference" {
                if let Some(attrs) = data.attributes() {
                    let id = attrs.iter()
                        .find(|a| a.name.local_name == "id" && a.name.namespace.as_deref() == Some(W::NS))
                        .map(|a| a.value.as_str())
                        .unwrap_or("unknown");
                    texts.push(format!(" ENDNOTE_REF_{} ", id));
                }
                return;
            }
            
            if ns == Some(M::NS) && (local == "oMath" || local == "oMathPara") {
                let math_hash = compute_math_hash(doc, node);
                texts.push(format!(" MATH_{:x} ", math_hash));
                return;
            }
            
            if ns == Some(MC::NS) && local == "AlternateContent" {
                if let Some(fallback) = find_child_by_name(doc, node, MC::NS, "Fallback") {
                    for child in doc.children(fallback) {
                        extract_text_recursive(doc, child, texts, accept_revisions);
                    }
                    return;
                }
                if let Some(choice) = find_child_by_name(doc, node, MC::NS, "Choice") {
                    for child in doc.children(choice) {
                        extract_text_recursive(doc, child, texts, accept_revisions);
                    }
                    return;
                }
            }
            
            if ns == Some(W::NS) && local == "txbxContent" {
                texts.push(" TXBX_START ".to_string());
                for child in doc.children(node) {
                    extract_text_recursive(doc, child, texts, accept_revisions);
                }
                texts.push(" TXBX_END ".to_string());
                return;
            }
            
            if ns == Some(W::NS) && local == "drawing" {
                let drawing_info = get_drawing_info(doc, node);
                texts.push(drawing_info);
                return;
            }
            
            if ns == Some(W::NS) && local == "pict" {
                if has_textbox(doc, node) {
                    for child in doc.children(node) {
                        extract_text_recursive(doc, child, texts, accept_revisions);
                    }
                    return;
                }
                let embed_ref = find_embed_reference(doc, node);
                texts.push(format!("PICT_{}", embed_ref.unwrap_or("unknown".to_string())));
                return;
            }
            
            for child in doc.children(node) {
                extract_text_recursive(doc, child, texts, accept_revisions);
            }
        }
        _ => {}
    }
}

fn find_child_by_name(doc: &XmlDocument, parent: NodeId, ns: &str, local: &str) -> Option<NodeId> {
    for child in doc.children(parent) {
        if let Some(data) = doc.get(child) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(ns) && name.local_name == local {
                    return Some(child);
                }
            }
        }
    }
    None
}

fn compute_math_hash(doc: &XmlDocument, node: NodeId) -> u32 {
    let mut math_texts = Vec::new();
    extract_math_text(doc, node, &mut math_texts);
    let content = math_texts.join("");
    
    let mut hash: u32 = 0;
    for ch in content.chars() {
        hash = hash.wrapping_shl(5).wrapping_sub(hash).wrapping_add(ch as u32);
    }
    hash
}

fn extract_math_text(doc: &XmlDocument, node: NodeId, texts: &mut Vec<String>) {
    if let Some(data) = doc.get(node) {
        if let Some(name) = data.name() {
            if name.namespace.as_deref() == Some(M::NS) && name.local_name == "t" {
                for child in doc.children(node) {
                    if let Some(XmlNodeData::Text(text)) = doc.get(child) {
                        texts.push(text.clone());
                    }
                }
                return;
            }
        }
    }
    
    for child in doc.children(node) {
        extract_math_text(doc, child, texts);
    }
}

fn get_drawing_info(doc: &XmlDocument, node: NodeId) -> String {
    let mut parts = Vec::new();
    
    if let Some(extent) = find_descendant_by_local_name(doc, node, "extent") {
        if let Some(data) = doc.get(extent) {
            if let Some(attrs) = data.attributes() {
                for attr in attrs {
                    if attr.name.local_name == "cx" {
                        parts.push(format!("cx{}", attr.value));
                    }
                    if attr.name.local_name == "cy" {
                        parts.push(format!("cy{}", attr.value));
                    }
                }
            }
        }
    }
    
    if let Some(embed_ref) = find_embed_reference(doc, node) {
        parts.push(format!("e{}", embed_ref));
    }
    
    format!("DRAWING_{}", if parts.is_empty() { "unknown".to_string() } else { parts.join("_") })
}

fn find_descendant_by_local_name(doc: &XmlDocument, start: NodeId, local: &str) -> Option<NodeId> {
    for desc in doc.descendants(start) {
        if let Some(data) = doc.get(desc) {
            if let Some(name) = data.name() {
                if name.local_name == local {
                    return Some(desc);
                }
            }
        }
    }
    None
}

fn find_embed_reference(doc: &XmlDocument, node: NodeId) -> Option<String> {
    if let Some(data) = doc.get(node) {
        if let Some(attrs) = data.attributes() {
            for attr in attrs {
                let is_embed_attr = attr.name.local_name == "embed" 
                    || attr.name.local_name == "relid" 
                    || attr.name.local_name == "id";
                if is_embed_attr && attr.name.namespace.is_some() {
                    return Some(attr.value.clone());
                }
            }
        }
    }
    
    for child in doc.children(node) {
        if let Some(result) = find_embed_reference(doc, child) {
            return Some(result);
        }
    }
    
    None
}

fn has_textbox(doc: &XmlDocument, node: NodeId) -> bool {
    for desc in doc.descendants(node) {
        if let Some(data) = doc.get(desc) {
            if let Some(name) = data.name() {
                if name.local_name == "textbox" {
                    return true;
                }
            }
        }
    }
    false
}

pub fn find_paragraphs(doc: &XmlDocument, start: NodeId) -> Vec<NodeId> {
    let mut paragraphs = Vec::new();
    
    for node in doc.descendants(start) {
        if let Some(data) = doc.get(node) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "p" {
                    paragraphs.push(node);
                }
            }
        }
    }
    
    paragraphs
}

pub fn extract_all_text(doc: &XmlDocument, body: NodeId) -> String {
    let paragraphs = find_paragraphs(doc, body);
    let mut texts = Vec::new();
    
    for para in paragraphs {
        let para_text = extract_paragraph_text(doc, para);
        if !para_text.is_empty() {
            texts.push(para_text);
        }
    }
    
    texts.join("\n").trim().to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extract_simple_text() {
        let mut doc = XmlDocument::new();
        let body = doc.add_root(XmlNodeData::element(W::body()));
        let para = doc.add_child(body, XmlNodeData::element(W::p()));
        let run = doc.add_child(para, XmlNodeData::element(W::r()));
        let text_elem = doc.add_child(run, XmlNodeData::element(W::t()));
        doc.add_child(text_elem, XmlNodeData::Text("Hello World".to_string()));
        
        let result = extract_paragraph_text(&doc, para);
        assert_eq!(result, "Hello World");
    }
}
