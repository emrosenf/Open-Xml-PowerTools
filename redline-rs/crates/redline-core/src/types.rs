use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum DocumentType {
    Word,
    Excel,
    PowerPoint,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum RevisionType {
    Inserted,
    Deleted,
    Modified,
    Moved,
    FormattingChanged,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Revision {
    pub revision_type: RevisionType,
    pub text: Option<String>,
    pub author: Option<String>,
    pub date: Option<String>,
}
