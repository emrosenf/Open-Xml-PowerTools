use crate::error::{RedlineError, Result};
use crate::xml::XmlDocument;
use std::collections::HashMap;
use std::io::{Cursor, Read, Write};
use zip::read::ZipArchive;
use zip::write::ZipWriter;
use zip::CompressionMethod;

use super::content_types::ContentTypes;
use super::relationships::Relationship;

pub struct OoxmlPackage {
    parts: HashMap<String, Vec<u8>>,
    content_types: ContentTypes,
    relationships: HashMap<String, Vec<Relationship>>,
}

impl OoxmlPackage {
    pub fn open(bytes: &[u8]) -> Result<Self> {
        let cursor = Cursor::new(bytes);
        let mut archive = ZipArchive::new(cursor)?;
        
        let mut parts = HashMap::new();
        
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = file.name().to_string();
            let mut content = Vec::new();
            file.read_to_end(&mut content)?;
            parts.insert(name, content);
        }
        
        let content_types = ContentTypes::default();
        let relationships = HashMap::new();
        
        Ok(Self {
            parts,
            content_types,
            relationships,
        })
    }

    pub fn save(&self) -> Result<Vec<u8>> {
        let mut buffer = Cursor::new(Vec::new());
        let mut writer = ZipWriter::new(&mut buffer);
        
        for (path, content) in &self.parts {
            let options: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default()
                .compression_method(CompressionMethod::Deflated);
            writer.start_file(path, options)?;
            writer.write_all(content)?;
        }
        
        writer.finish()?;
        Ok(buffer.into_inner())
    }

    pub fn get_part(&self, path: &str) -> Option<&[u8]> {
        self.parts.get(path).map(|v| v.as_slice())
    }

    pub fn get_xml_part(&self, path: &str) -> Result<XmlDocument> {
        let bytes = self.get_part(path).ok_or_else(|| RedlineError::MissingPart {
            part_path: path.to_string(),
            document_type: "OOXML".to_string(),
        })?;
        crate::xml::parser::parse_bytes(bytes)
    }

    pub fn set_part(&mut self, path: &str, content: Vec<u8>) {
        self.parts.insert(path.to_string(), content);
    }

    pub fn put_xml_part(&mut self, path: &str, doc: &XmlDocument) -> Result<()> {
        let bytes = crate::xml::builder::serialize_bytes(doc)?;
        self.set_part(path, bytes);
        Ok(())
    }

    pub fn delete_part(&mut self, path: &str) {
        self.parts.remove(path);
    }

    pub fn get_relationships(&self, source: &str) -> &[Relationship] {
        self.relationships
            .get(source)
            .map(|v| v.as_slice())
            .unwrap_or(&[])
    }

    pub fn add_relationship(&mut self, source: &str, rel: Relationship) {
        self.relationships
            .entry(source.to_string())
            .or_insert_with(Vec::new)
            .push(rel);
    }

    pub fn get_content_type(&self, path: &str) -> Option<&str> {
        self.content_types.get_content_type(path)
    }

    pub fn set_content_type(&mut self, path: &str, content_type: &str) {
        self.content_types.set_content_type(path, content_type);
    }

    pub fn part_names(&self) -> impl Iterator<Item = &String> {
        self.parts.keys()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn package_roundtrip() {
        let mut pkg = OoxmlPackage {
            parts: HashMap::new(),
            content_types: ContentTypes::default(),
            relationships: HashMap::new(),
        };
        
        pkg.set_part("test.xml", b"<root/>".to_vec());
        
        let saved = pkg.save().unwrap();
        let loaded = OoxmlPackage::open(&saved).unwrap();
        
        assert!(loaded.get_part("test.xml").is_some());
    }
}
