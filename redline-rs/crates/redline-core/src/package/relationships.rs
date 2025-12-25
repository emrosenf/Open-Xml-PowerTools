use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[derive(Default)]
pub enum TargetMode {
    #[default]
    Internal,
    External,
}


#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    #[serde(default)]
    pub target_mode: TargetMode,
}

impl Relationship {
    pub fn new(id: &str, rel_type: &str, target: &str) -> Self {
        Self {
            id: id.to_string(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: TargetMode::Internal,
        }
    }

    pub fn external(id: &str, rel_type: &str, target: &str) -> Self {
        Self {
            id: id.to_string(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: TargetMode::External,
        }
    }
}

pub mod relationship_types {
    pub const OFFICE_DOCUMENT: &str = 
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub const STYLES: &str = 
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const NUMBERING: &str = 
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub const FOOTNOTES: &str = 
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    pub const ENDNOTES: &str = 
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    pub const COMMENTS: &str = 
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const HYPERLINK: &str = 
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const IMAGE: &str = 
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
}
