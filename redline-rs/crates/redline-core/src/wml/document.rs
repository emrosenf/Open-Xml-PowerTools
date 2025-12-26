use crate::error::Result;
use crate::package::OoxmlPackage;
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{MC, M, W};
use crate::xml::node::XmlNodeData;
use indextree::NodeId;
use std::io::Write;

pub struct WmlDocument {
    package: OoxmlPackage,
}

impl WmlDocument {
    /// Create a minimal WML document package from main XML content (useful for testing)
    pub fn from_main_xml(main_xml: &[u8]) -> Result<Self> {
        let mut buffer = std::io::Cursor::new(Vec::new());
        {
            let mut zip = zip::ZipWriter::new(&mut buffer);
            
            // Add [Content_Types].xml
            zip.start_file("[Content_Types].xml", zip::write::SimpleFileOptions::default())?;
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#)?;

            // Add word/document.xml
            zip.start_file("word/document.xml", zip::write::SimpleFileOptions::default())?;
            zip.write_all(main_xml)?;
            
            zip.finish()?;
        }
        
        Self::from_bytes(&buffer.into_inner())
    }

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

    pub fn revisions(&self) -> crate::wml::revision::RevisionCounts {
        use crate::wml::revision::{count_revisions, RevisionCounts};
        let mut total_counts = RevisionCounts::default();

        if let Ok(doc) = self.main_document() {
            if let Some(root) = doc.root() {
                let counts = count_revisions(&doc, root);
                total_counts.insertions += counts.insertions;
                total_counts.deletions += counts.deletions;
                total_counts.format_changes += counts.format_changes;
            }
        }

        if let Ok(Some(doc)) = self.footnotes() {
            if let Some(root) = doc.root() {
                let counts = count_revisions(&doc, root);
                total_counts.insertions += counts.insertions;
                total_counts.deletions += counts.deletions;
                total_counts.format_changes += counts.format_changes;
            }
        }

        if let Ok(Some(doc)) = self.endnotes() {
            if let Some(root) = doc.root() {
                let counts = count_revisions(&doc, root);
                total_counts.insertions += counts.insertions;
                total_counts.deletions += counts.deletions;
                total_counts.format_changes += counts.format_changes;
            }
        }

        total_counts
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
                        .unwrap_or("0");
                    texts.push(format!(" FOOTNOTEREF{} ", id));
                }
                return;
            }
            
            if ns == Some(W::NS) && local == "endnoteReference" {
                if let Some(attrs) = data.attributes() {
                    let id = attrs.iter()
                        .find(|a| a.name.local_name == "id" && a.name.namespace.as_deref() == Some(W::NS))
                        .map(|a| a.value.as_str())
                        .unwrap_or("0");
                    texts.push(format!(" ENDNOTEREF{} ", id));
                }
                return;
            }
            
            if ns == Some(M::NS) && (local == "oMath" || local == "oMathPara") {
                let math_hash = compute_math_hash(doc, node);
                texts.push(format!(" MATH{:08x} ", math_hash));
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
                texts.push(" TXBXSTART ".to_string());
                for child in doc.children(node) {
                    extract_text_recursive(doc, child, texts, accept_revisions);
                }
                texts.push(" TXBXEND ".to_string());
                return;
            }
            
            if ns == Some(W::NS) && local == "drawing" {
                let drawing_info = get_drawing_info(doc, node);
                texts.push(format!(" {} ", drawing_info));
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
                let hash: u32 = embed_ref.as_deref().unwrap_or("unknown").bytes()
                    .fold(0u32, |acc, b| acc.wrapping_mul(31).wrapping_add(b as u32));
                texts.push(format!(" PICT{:08x} ", hash));
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
    let mut hash_input = String::new();
    
    if let Some(extent) = find_descendant_by_local_name(doc, node, "extent") {
        if let Some(data) = doc.get(extent) {
            if let Some(attrs) = data.attributes() {
                for attr in attrs {
                    if attr.name.local_name == "cx" || attr.name.local_name == "cy" {
                        hash_input.push_str(&attr.value);
                    }
                }
            }
        }
    }
    
    if let Some(embed_ref) = find_embed_reference(doc, node) {
        hash_input.push_str(&embed_ref);
    }
    
    if hash_input.is_empty() {
        "DRAWINGunknown".to_string()
    } else {
        let hash: u32 = hash_input.bytes().fold(0u32, |acc, b| acc.wrapping_mul(31).wrapping_add(b as u32));
        format!("DRAWING{:08x}", hash)
    }
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
    find_paragraphs_recursive(doc, start, &mut paragraphs, false);
    paragraphs
}

fn find_paragraphs_recursive(doc: &XmlDocument, node: NodeId, paragraphs: &mut Vec<NodeId>, in_textbox: bool) {
    if let Some(data) = doc.get(node) {
        if let Some(name) = data.name() {
            let ns = name.namespace.as_deref();
            let local = name.local_name.as_str();
            
            if ns == Some(W::NS) && local == "txbxContent" {
                for child in doc.children(node) {
                    find_paragraphs_recursive(doc, child, paragraphs, true);
                }
                return;
            }
            
            if ns == Some(W::NS) && local == "p" && !in_textbox {
                paragraphs.push(node);
            }
        }
    }
    
    for child in doc.children(node) {
        find_paragraphs_recursive(doc, child, paragraphs, in_textbox);
    }
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

pub fn find_footnotes_root(doc: &XmlDocument) -> Option<NodeId> {
    let root = doc.root()?;
    
    for child in doc.descendants(root) {
        if let Some(data) = doc.get(child) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "footnotes" {
                    return Some(child);
                }
            }
        }
    }
    
    None
}

pub fn find_endnotes_root(doc: &XmlDocument) -> Option<NodeId> {
    let root = doc.root()?;
    
    for child in doc.descendants(root) {
        if let Some(data) = doc.get(child) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "endnotes" {
                    return Some(child);
                }
            }
        }
    }
    
    None
}

pub fn find_note_paragraphs(doc: &XmlDocument, root: NodeId) -> Vec<NodeId> {
    let mut paragraphs = Vec::new();
    
    for node in doc.descendants(root) {
        if let Some(data) = doc.get(node) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) 
                    && (name.local_name == "footnote" || name.local_name == "endnote") 
                {
                    if let Some(attrs) = data.attributes() {
                        let id = attrs.iter()
                            .find(|a| a.name.local_name == "id" && a.name.namespace.as_deref() == Some(W::NS))
                            .map(|a| a.value.as_str());
                        if let Some(id_val) = id {
                            if id_val == "0" || id_val == "-1" {
                                continue;
                            }
                        }
                    }
                    for para in find_paragraphs(doc, node) {
                        paragraphs.push(para);
                    }
                }
            }
        }
    }
    
    paragraphs
}

pub fn find_note_by_id(doc: &XmlDocument, root: NodeId, id: &str) -> Option<NodeId> {
    for node in doc.descendants(root) {
        if let Some(data) = doc.get(node) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) 
                    && (name.local_name == "footnote" || name.local_name == "endnote") 
                {
                    if let Some(attrs) = data.attributes() {
                        let node_id = attrs.iter()
                            .find(|a| a.name.local_name == "id" && a.name.namespace.as_deref() == Some(W::NS))
                            .map(|a| a.value.as_str());
                        if let Some(id_val) = node_id {
                            if id_val == id {
                                return Some(node);
                            }
                        }
                    }
                }
            }
        }
    }
    None
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
