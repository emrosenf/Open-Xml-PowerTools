//! ComparisonUnit hierarchy for WmlComparer
//!
//! This is a faithful port of the C# ComparisonUnit classes from WmlComparer.cs.
//!
//! The hierarchy is:
//! - ComparisonUnit (trait) - base for all comparison units
//!   - ComparisonUnitAtom - atomic content element (w:t chars, w:pPr, w:drawing, etc.)
//!   - ComparisonUnitWord - group of atoms forming a "word"
//!   - ComparisonUnitGroup - hierarchical grouping (Paragraph, Table, Row, Cell, Textbox)
//!
//! Key features:
//! - Each unit has a SHA1 hash for comparison (identity_hash)
//! - Atoms track ancestor elements with Unids for tree reconstruction
//! - Groups have CorrelatedSHA1Hash for efficient block-level matching
//!
//! ## Canonical Atom Model (RUST-1)
//!
//! The canonical atom model ensures consistent identity across all comparison stages:
//!
//! 1. **identity_hash()** - Stable SHA1 hash computed from:
//!    - Element local name (e.g., "t" for text, "pPr" for paragraph properties)
//!    - Content value (text characters, or hash for drawings/objects)
//!    - Settings-based normalization (case insensitivity, space conflation)
//!
//! 2. **content_type()** - Enum representing the atom's content category:
//!    - Text, Drawing, Field, ParagraphMark, etc.
//!
//! 3. **formatting_signature()** - Hash of normalized rPr for format change detection
//!
//! 4. **ancestor_unids()** - For tree reconstruction during result assembly
//!
//! The `PartialEq` implementation uses identity_hash for comparison, not structural equality.

use crate::util::lcs::Hashable;
use crate::wml::settings::WmlComparerSettings;
use indextree::NodeId;
use sha1::{Digest, Sha1};
use std::fmt;
use std::sync::Arc;

/// Correlation status for comparison units
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ComparisonCorrelationStatus {
    #[default]
    Nil,
    Normal,
    Unknown,
    Inserted,
    Deleted,
    Equal,
    FormatChanged,
    Group,
}

impl fmt::Display for ComparisonCorrelationStatus {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Nil => write!(f, "Nil"),
            Self::Normal => write!(f, "Normal"),
            Self::Unknown => write!(f, "Unknown"),
            Self::Inserted => write!(f, "Inserted"),
            Self::Deleted => write!(f, "Deleted"),
            Self::Equal => write!(f, "Equal"),
            Self::FormatChanged => write!(f, "FormatChanged"),
            Self::Group => write!(f, "Group"),
        }
    }
}

/// Type of comparison unit group
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComparisonUnitGroupType {
    Paragraph,
    Table,
    Row,
    Cell,
    Textbox,
}

impl fmt::Display for ComparisonUnitGroupType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Paragraph => write!(f, "Paragraph"),
            Self::Table => write!(f, "Table"),
            Self::Row => write!(f, "Row"),
            Self::Cell => write!(f, "Cell"),
            Self::Textbox => write!(f, "Textbox"),
        }
    }
}

/// Ancestor element info for tree reconstruction
#[derive(Debug, Clone)]
pub struct AncestorInfo {
    /// Node ID of the ancestor element
    pub node_id: NodeId,
    /// Namespace URI for the element (None for no namespace)
    pub namespace: Option<String>,
    /// Local name of the element (e.g., "p", "tbl", "tr", "tc")
    pub local_name: String,
    /// Unique ID (Unid) for this element - used for correlation
    pub unid: String,
    /// Attributes from the ancestor element (for reconstruction)
    pub attributes: Arc<Vec<crate::xml::xname::XAttribute>>,
    /// Whether this table cell has merged cell properties (vMerge or gridSpan)
    /// Used by DoLcsAlgorithmForTable to detect merged cells
    #[allow(dead_code)]
    pub has_merged_cells: bool,
}

/// Content type classification for atoms
/// Corresponds to the different types of content that can appear in a document
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ContentType {
    /// Text content (w:t, w:delText)
    Text,
    /// Paragraph properties marker (w:pPr)
    ParagraphMark,
    /// Drawing/image content (w:drawing)
    Drawing,
    /// Picture/VML content (w:pict)
    Picture,
    /// Math content (m:oMath, m:oMathPara)
    Math,
    /// Field content (w:fldChar, w:fldSimple)
    Field,
    /// Symbol content (w:sym)
    Symbol,
    /// Object content (w:object)
    Object,
    /// Footnote reference (w:footnoteReference)
    FootnoteReference,
    /// Endnote reference (w:endnoteReference)
    EndnoteReference,
    /// Break content (w:br, w:cr)
    Break,
    /// Tab content (w:tab, w:ptab)
    Tab,
    /// Textbox marker
    Textbox,
    /// Unknown/other content
    Unknown,
}

/// Content element types for atoms
#[derive(Debug, Clone, PartialEq)]
pub enum ContentElement {
    /// Text character - single character from w:t
    Text(char),
    /// Paragraph properties marker
    ParagraphProperties,
    /// Run properties marker
    RunProperties,
    /// Line break (w:br)
    Break,
    /// Tab (w:tab)
    Tab,
    /// Drawing/image with hash
    Drawing { hash: String },
    /// Picture (VML) with hash
    Picture { hash: String },
    /// Math equation with hash
    Math { hash: String },
    /// Footnote reference with ID
    FootnoteReference { id: String },
    /// Endnote reference with ID
    EndnoteReference { id: String },
    /// Textbox start marker
    TextboxStart,
    /// Textbox end marker
    TextboxEnd,
    /// Field begin (w:fldChar with fldCharType="begin")
    FieldBegin,
    /// Field separator (w:fldChar with fldCharType="separate")
    FieldSeparator,
    /// Field end (w:fldChar with fldCharType="end")
    FieldEnd,
    /// Simple field (w:fldSimple)
    SimpleField { instruction: String },
    /// Symbol (w:sym)
    Symbol { font: String, char_code: String },
    /// Object (w:object)
    Object { hash: String },
    /// Unknown element
    Unknown { name: String },
}

impl ContentElement {
    /// Get content type classification for this element
    pub fn content_type(&self) -> ContentType {
        match self {
            ContentElement::Text(_) => ContentType::Text,
            ContentElement::ParagraphProperties => ContentType::ParagraphMark,
            ContentElement::RunProperties => ContentType::Unknown, // rPr atoms are rare
            ContentElement::Break => ContentType::Break,
            ContentElement::Tab => ContentType::Tab,
            ContentElement::Drawing { .. } => ContentType::Drawing,
            ContentElement::Picture { .. } => ContentType::Picture,
            ContentElement::Math { .. } => ContentType::Math,
            ContentElement::FootnoteReference { .. } => ContentType::FootnoteReference,
            ContentElement::EndnoteReference { .. } => ContentType::EndnoteReference,
            ContentElement::TextboxStart | ContentElement::TextboxEnd => ContentType::Textbox,
            ContentElement::FieldBegin | ContentElement::FieldSeparator | 
            ContentElement::FieldEnd | ContentElement::SimpleField { .. } => ContentType::Field,
            ContentElement::Symbol { .. } => ContentType::Symbol,
            ContentElement::Object { .. } => ContentType::Object,
            ContentElement::Unknown { .. } => ContentType::Unknown,
        }
    }

    /// Get the element local name for hash computation
    /// Corresponds to C# contentElement.Name.LocalName
    pub fn local_name(&self) -> &'static str {
        match self {
            ContentElement::Text(_) => "t",
            ContentElement::ParagraphProperties => "pPr",
            ContentElement::RunProperties => "rPr",
            ContentElement::Break => "br",
            ContentElement::Tab => "tab",
            ContentElement::Drawing { .. } => "drawing",
            ContentElement::Picture { .. } => "pict",
            ContentElement::Math { .. } => "oMath",
            ContentElement::FootnoteReference { .. } => "footnoteReference",
            ContentElement::EndnoteReference { .. } => "endnoteReference",
            ContentElement::TextboxStart => "txbxContent", // start marker
            ContentElement::TextboxEnd => "txbxContent",   // end marker
            ContentElement::FieldBegin => "fldChar",
            ContentElement::FieldSeparator => "fldChar",
            ContentElement::FieldEnd => "fldChar",
            ContentElement::SimpleField { .. } => "fldSimple",
            ContentElement::Symbol { .. } => "sym",
            ContentElement::Object { .. } => "object",
            ContentElement::Unknown { name } => {
                // Return a static str for unknown - we'll use "unknown"
                // since we can't return a &'static str from a dynamic String
                let _ = name; // suppress unused warning
                "unknown"
            },
        }
    }

    /// Get the text value for hash computation
    /// Corresponds to C# contentElement.Value
    /// For text elements, this is the character; for drawings, it's the content hash
    pub fn text_value(&self) -> String {
        match self {
            ContentElement::Text(ch) => ch.to_string(),
            ContentElement::ParagraphProperties => String::new(), // pPr has no text value
            ContentElement::RunProperties => String::new(),
            ContentElement::Break => String::new(),
            ContentElement::Tab => String::new(),
            ContentElement::Drawing { hash } => hash.clone(),
            ContentElement::Picture { hash } => hash.clone(),
            ContentElement::Math { hash } => hash.clone(),
            ContentElement::FootnoteReference { .. } => String::new(),
            ContentElement::EndnoteReference { .. } => String::new(),
            ContentElement::TextboxStart => String::new(),
            ContentElement::TextboxEnd => String::new(),
            ContentElement::FieldBegin => "begin".to_string(),
            ContentElement::FieldSeparator => "separate".to_string(),
            ContentElement::FieldEnd => "end".to_string(),
            ContentElement::SimpleField { instruction } => instruction.clone(),
            ContentElement::Symbol { font, char_code } => format!("{}:{}", font, char_code),
            ContentElement::Object { hash } => hash.clone(),
            ContentElement::Unknown { name } => name.clone(),
        }
    }

    /// Get hash string for this content element (used for SHA1 computation)
    /// 
    /// This is a faithful port of C# GetSha1HashStringForElement (WmlComparer.cs:8468-8477)
    /// The hash string is: localName + textValue
    /// 
    /// For text: "t" + character
    /// For drawings: "drawing" + contentHash
    /// For pPr: "pPr" (empty text value)
    pub fn hash_string(&self) -> String {
        format!("{}{}", self.local_name(), self.text_value())
    }

    /// Get hash string with settings-based normalization
    /// 
    /// This is the full C# GetSha1HashStringForElement implementation:
    /// - Applies case transformation if CaseInsensitive
    /// - Applies space normalization if ConflateBreakingAndNonbreakingSpaces
    pub fn hash_string_with_settings(&self, settings: &WmlComparerSettings) -> String {
        let mut text = self.text_value();
        
        if settings.case_insensitive {
            // C#: text = text.ToUpper(settings.CultureInfo)
            // We use simple to_uppercase since we don't have full CultureInfo support
            text = text.to_uppercase();
        }
        
        if settings.conflate_breaking_and_nonbreaking_spaces {
            // C#: text = text.Replace(' ', '\x00a0')
            // Replace regular space with non-breaking space
            text = text.replace(' ', "\u{00a0}");
        }
        
        format!("{}{}", self.local_name(), text)
    }

    /// Get display value for this content element
    pub fn display_value(&self) -> String {
        match self {
            ContentElement::Text(ch) => ch.to_string(),
            ContentElement::ParagraphProperties => "¶".to_string(),
            ContentElement::Break => "⏎".to_string(),
            ContentElement::Tab => "→".to_string(),
            _ => "".to_string(),
        }
    }
}

/// Atomic comparison unit - the smallest unit of comparison
/// Corresponds to C# ComparisonUnitAtom (WmlComparer.cs:8280)
///
/// ## Canonical Identity Model
///
/// The atom's identity is defined by its `sha1_hash` (identity_hash), which is computed from:
/// - Element local name (e.g., "t" for text)
/// - Content value (text character or content hash)
/// - Settings-based normalization
///
/// Two atoms are considered equal if they have the same identity hash.
/// This enables consistent LCS matching across all comparison stages.
#[derive(Debug, Clone)]
pub struct ComparisonUnitAtom {
    /// The content element this atom represents
    pub content_element: ContentElement,
    /// SHA1 hash of the content - THE CANONICAL IDENTITY
    /// Computed using GetSha1HashStringForElement logic:
    /// SHA1(localName + textValue [with case/space normalization])
    pub sha1_hash: String,
    /// Ancestor elements from body to this element (body → leaf order)
    pub ancestor_elements: Vec<AncestorInfo>,
    /// Correlation status
    pub correlation_status: ComparisonCorrelationStatus,
    /// Formatting signature (for TrackFormattingChanges)
    /// This is the serialized normalized rPr element
    pub formatting_signature: Option<String>,
    /// Normalized run properties (for format change detection)
    pub normalized_rpr: Option<String>,
    /// Part name this atom belongs to (main, footnotes, endnotes)
    pub part_name: String,
    
    // Fields for "before" document tracking (Equal/FormatChanged atoms)
    /// Content element from "before" document
    pub content_element_before: Option<ContentElement>,
    /// Formatting signature from "before" document
    pub formatting_signature_before: Option<String>,
    /// Reference to comparison unit atom from "before" document
    pub comparison_unit_atom_before: Option<Box<ComparisonUnitAtom>>,
    /// Ancestor elements from "before" document
    pub ancestor_elements_before: Option<Vec<AncestorInfo>>,
    /// Part name from "before" document
    pub part_before: Option<String>,
    /// Revision tracking element (w:ins or w:del)
    pub rev_track_element: Option<String>,
    /// Formatting change rPr from "before" document
    pub formatting_change_rpr_before: Option<String>,
    
    // Fields populated by AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly
    /// Ancestor Unids array (from C# AncestorUnids property)
    /// This is populated after correlation and is used for tree reconstruction
    pub ancestor_unids: Vec<String>,
    /// Formatting change rPr "before" signature (for grouping)
    pub formatting_change_rpr_before_signature: Option<String>,
}

/// Implement PartialEq based on identity hash, not structural equality
/// This matches C# behavior where atoms are compared by SHA1Hash
impl PartialEq for ComparisonUnitAtom {
    fn eq(&self, other: &Self) -> bool {
        self.sha1_hash == other.sha1_hash
    }
}

impl Eq for ComparisonUnitAtom {}

impl ComparisonUnitAtom {
    /// Create a new atom with the given content element and ancestors
    /// 
    /// This is a faithful port of C# ComparisonUnitAtom constructor (WmlComparer.cs:8347-8378)
    /// 
    /// The identity hash is computed from:
    /// - Element local name + text value (with settings normalization)
    /// - Uses pre-computed SHA1Hash attribute if present (from preprocessing)
    pub fn new(
        content_element: ContentElement,
        ancestor_elements: Vec<AncestorInfo>,
        part_name: &str,
        settings: &WmlComparerSettings,
    ) -> Self {
        // Compute identity hash using C# GetSha1HashStringForElement logic
        let hash_string = content_element.hash_string_with_settings(settings);
        let sha1_hash = compute_sha1(&hash_string);

        // Find revision tracking element from ancestors (C# lines 8352-8363)
        // Search from leaf to body (reverse order)
        let mut correlation_status = ComparisonCorrelationStatus::Equal;
        let mut rev_track_element = None;
        
        // C#: revTrackElement = ancestors.FirstOrDefault(a => a.Name == W.del || a.Name == W.ins);
        for ancestor in ancestor_elements.iter().rev() {
            if ancestor.local_name == "ins" {
                correlation_status = ComparisonCorrelationStatus::Inserted;
                rev_track_element = Some("ins".to_string());
                break;
            } else if ancestor.local_name == "del" {
                correlation_status = ComparisonCorrelationStatus::Deleted;
                rev_track_element = Some("del".to_string());
                break;
            }
        }

        Self {
            content_element,
            sha1_hash,
            ancestor_elements,
            correlation_status,
            formatting_signature: None, // Populated separately via set_formatting_signature
            normalized_rpr: None,
            part_name: part_name.to_string(),
            content_element_before: None,
            formatting_signature_before: None,
            comparison_unit_atom_before: None,
            ancestor_elements_before: None,
            part_before: None,
            rev_track_element,
            formatting_change_rpr_before: None,
            ancestor_unids: Vec::new(),
            formatting_change_rpr_before_signature: None,
        }
    }
    
    /// Create a new atom with a pre-computed SHA1 hash
    /// 
    /// This is used when the hash has been pre-computed during preprocessing
    /// (C# lines 8364-8368: checks for PtOpenXml.SHA1Hash attribute)
    pub fn new_with_hash(
        content_element: ContentElement,
        ancestor_elements: Vec<AncestorInfo>,
        part_name: &str,
        sha1_hash: String,
    ) -> Self {
        // Find revision tracking element from ancestors
        let mut correlation_status = ComparisonCorrelationStatus::Equal;
        let mut rev_track_element = None;
        
        for ancestor in ancestor_elements.iter().rev() {
            if ancestor.local_name == "ins" {
                correlation_status = ComparisonCorrelationStatus::Inserted;
                rev_track_element = Some("ins".to_string());
                break;
            } else if ancestor.local_name == "del" {
                correlation_status = ComparisonCorrelationStatus::Deleted;
                rev_track_element = Some("del".to_string());
                break;
            }
        }

        Self {
            content_element,
            sha1_hash,
            ancestor_elements,
            correlation_status,
            formatting_signature: None,
            normalized_rpr: None,
            part_name: part_name.to_string(),
            content_element_before: None,
            formatting_signature_before: None,
            comparison_unit_atom_before: None,
            ancestor_elements_before: None,
            part_before: None,
            rev_track_element,
            formatting_change_rpr_before: None,
            ancestor_unids: Vec::new(),
            formatting_change_rpr_before_signature: None,
        }
    }
    
    /// Get the canonical identity hash for this atom
    /// This is the primary identifier used for LCS comparison
    pub fn identity_hash(&self) -> &str {
        &self.sha1_hash
    }
    
    /// Get the content type classification
    pub fn content_type(&self) -> ContentType {
        self.content_element.content_type()
    }
    
    /// Get the formatting signature (hash of normalized rPr)
    /// Returns None if formatting tracking is disabled or no rPr exists
    pub fn formatting_signature(&self) -> Option<&str> {
        self.formatting_signature.as_deref()
    }
    
    /// Set the formatting signature for this atom
    pub fn set_formatting_signature(&mut self, signature: Option<String>) {
        self.formatting_signature = signature;
    }
    
    /// Get ancestor Unids for tree reconstruction
    /// These are populated by AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly
    pub fn ancestor_unids(&self) -> &[String] {
        &self.ancestor_unids
    }
    
    /// Set ancestor Unids for tree reconstruction
    pub fn set_ancestor_unids(&mut self, unids: Vec<String>) {
        self.ancestor_unids = unids;
    }

    /// Get the Unid of the nth ancestor (0 = closest to body)
    pub fn ancestor_unid(&self, index: usize) -> Option<&str> {
        self.ancestor_elements.get(index).map(|a| a.unid.as_str())
    }

    /// Get the local name of the nth ancestor
    pub fn ancestor_name(&self, index: usize) -> Option<&str> {
        self.ancestor_elements.get(index).map(|a| a.local_name.as_str())
    }

    /// Check if this atom is inside a table
    pub fn is_in_table(&self) -> bool {
        self.ancestor_elements.iter().any(|a| a.local_name == "tbl")
    }

    /// Check if this atom is inside a table cell
    pub fn is_in_cell(&self) -> bool {
        self.ancestor_elements.iter().any(|a| a.local_name == "tc")
    }

    /// Get the paragraph Unid for this atom
    pub fn paragraph_unid(&self) -> Option<&str> {
        self.ancestor_elements
            .iter()
            .rev()
            .find(|a| a.local_name == "p")
            .map(|a| a.unid.as_str())
    }

    /// Get the table row Unid for this atom
    pub fn row_unid(&self) -> Option<&str> {
        self.ancestor_elements
            .iter()
            .rev()
            .find(|a| a.local_name == "tr")
            .map(|a| a.unid.as_str())
    }

    /// Get the table cell Unid for this atom
    pub fn cell_unid(&self) -> Option<&str> {
        self.ancestor_elements
            .iter()
            .rev()
            .find(|a| a.local_name == "tc")
            .map(|a| a.unid.as_str())
    }

    /// Get the table Unid for this atom
    pub fn table_unid(&self) -> Option<&str> {
        self.ancestor_elements
            .iter()
            .rev()
            .find(|a| a.local_name == "tbl")
            .map(|a| a.unid.as_str())
    }

    /// Format as string with indentation
    /// Corresponds to C# ToString(int indent) (WmlComparer.cs:8496)
    pub fn to_string_with_indent(&self, indent: usize) -> String {
        const XNAME_PAD: usize = 16;
        let indent_str = " ".repeat(indent);
        let hash_short = if self.sha1_hash.len() >= 8 {
            &self.sha1_hash[..8]
        } else {
            &self.sha1_hash
        };
        
        let correlation_status_str = if self.correlation_status != ComparisonCorrelationStatus::Nil {
            format!("[{:8}] ", format!("{}", self.correlation_status))
        } else {
            String::new()
        };

        let element_name = match &self.content_element {
            ContentElement::Text(_) => "t",
            ContentElement::ParagraphProperties => "pPr",
            ContentElement::RunProperties => "rPr",
            ContentElement::Break => "br",
            ContentElement::Tab => "tab",
            ContentElement::Drawing { .. } => "drawing",
            ContentElement::Picture { .. } => "pict",
            ContentElement::Math { .. } => "oMath",
            ContentElement::FootnoteReference { .. } => "footnoteRef",
            ContentElement::EndnoteReference { .. } => "endnoteRef",
            ContentElement::TextboxStart => "txbxStart",
            ContentElement::TextboxEnd => "txbxEnd",
            ContentElement::FieldBegin => "fldBegin",
            ContentElement::FieldSeparator => "fldSep",
            ContentElement::FieldEnd => "fldEnd",
            ContentElement::SimpleField { .. } => "fldSimple",
            ContentElement::Symbol { .. } => "sym",
            ContentElement::Object { .. } => "object",
            ContentElement::Unknown { name } => name.as_str(),
        };

        let padded_name = format!("{} ", element_name).chars()
            .chain(std::iter::repeat('-'))
            .take(XNAME_PAD)
            .collect::<String>();

        let display_val = self.content_element.display_value();
        let value_part = if !display_val.is_empty() {
            format!(": {} ", display_val)
        } else {
            ":   ".to_string()
        };

        let ancestors_str = self.ancestor_elements
            .iter()
            .map(|a| {
                let unid_short = if a.unid.len() >= 8 {
                    &a.unid[..8]
                } else {
                    &a.unid
                };
                format!("{}[{}]/", a.local_name, unid_short)
            })
            .collect::<String>()
            .trim_end_matches('/')
            .to_string();

        format!(
            "{}Atom {} {} {}SHA1:{} Ancestors:{}",
            indent_str, padded_name, value_part, correlation_status_str, hash_short, ancestors_str
        )
    }
}

impl Hashable for ComparisonUnitAtom {
    fn hash(&self) -> &str {
        &self.sha1_hash
    }
}

impl ComparisonUnitAtom {
    /// Format a list of comparison unit atoms as a string
    /// Corresponds to C# ComparisonUnitAtomListToString (WmlComparer.cs:8582)
    pub fn comparison_unit_atom_list_to_string(atoms: &[ComparisonUnitAtom], indent: usize) -> String {
        let mut result = String::new();
        for (i, atom) in atoms.iter().enumerate() {
            let indent_str = " ".repeat(indent);
            result.push_str(&format!("{}[{:06}] {}\n", indent_str, i + 1, atom.to_string_with_indent(0)));
        }
        result
    }
}

/// Word-level comparison unit - groups atoms into words
/// Corresponds to C# ComparisonUnitWord (WmlComparer.cs:8212)
#[derive(Debug, Clone)]
pub struct ComparisonUnitWord {
    /// Atoms that make up this word
    pub atoms: Arc<Vec<ComparisonUnitAtom>>,
    /// SHA1 hash of all atom hashes concatenated
    pub sha1_hash: String,
    /// Correlation status
    pub correlation_status: ComparisonCorrelationStatus,
}

// Static sets for relationship tracking - corresponds to C# FrozenSets (lines 8224-8268)
// In Rust, we use lazy_static or const arrays. For O(1) lookup, we'd use HashSet at runtime.
// For now, keeping as const arrays since the C# uses FrozenSet for immutable lookup.
const ELEMENTS_WITH_RELATIONSHIP_IDS: &[&str] = &[
    "blip",          // A.blip
    "hlinkClick",    // A.hlinkClick
    "relIds",        // A.relIds, DGM.relIds
    "chart",         // C.chart
    "externalData",  // C.externalData
    "userShapes",    // C.userShapes
    "OLEObject",     // O.OLEObject
    "fill",          // VML.fill
    "imagedata",     // VML.imagedata
    "stroke",        // VML.stroke
    "altChunk",      // W.altChunk
    "attachedTemplate", // W.attachedTemplate
    "control",       // W.control
    "dataSource",    // W.dataSource
    "embedBold",     // W.embedBold
    "embedBoldItalic", // W.embedBoldItalic
    "embedItalic",   // W.embedItalic
    "embedRegular",  // W.embedRegular
    "footerReference", // W.footerReference
    "headerReference", // W.headerReference
    "headerSource",  // W.headerSource
    "hyperlink",     // W.hyperlink
    "printerSettings", // W.printerSettings
    "recipientData", // W.recipientData
    "saveThroughXslt", // W.saveThroughXslt
    "sourceFileName", // W.sourceFileName
    "src",           // W.src
    "subDoc",        // W.subDoc
    "toolbarData",   // WNE.toolbarData
];

const RELATIONSHIP_ATTRIBUTE_NAMES: &[&str] = &[
    "embed",  // R.embed
    "link",   // R.link
    "id",     // R.id
    "cs",     // R.cs
    "dm",     // R.dm
    "lo",     // R.lo
    "qs",     // R.qs
    "href",   // R.href
    "pict",   // R.pict
];

impl ComparisonUnitWord {
    /// Create a new word from a list of atoms
    pub fn new(atoms: Vec<ComparisonUnitAtom>) -> Self {
        let sha1_hash = compute_sha1_concat(atoms.iter().map(|a| a.sha1_hash.as_str()));

        Self {
            atoms: Arc::new(atoms),
            sha1_hash,
            correlation_status: ComparisonCorrelationStatus::Nil,
        }
    }

    /// Get the first atom in this word
    pub fn first_atom(&self) -> Option<&ComparisonUnitAtom> {
        self.atoms.first()
    }

    /// Get the text content of this word
    pub fn text(&self) -> String {
        self.atoms
            .iter()
            .map(|a| a.content_element.display_value())
            .collect()
    }

    /// Check if this word is just a paragraph mark
    pub fn is_paragraph_mark(&self) -> bool {
        self.atoms.len() == 1
            && matches!(
                self.atoms[0].content_element,
                ContentElement::ParagraphProperties
            )
    }

    /// Format as string with indentation
    /// Corresponds to C# ToString(int indent) (WmlComparer.cs:8270)
    pub fn to_string_with_indent(&self, indent: usize) -> String {
        let indent_str = " ".repeat(indent);
        let hash_short = if self.sha1_hash.len() >= 8 {
            &self.sha1_hash[..8]
        } else {
            &self.sha1_hash
        };
        
        let mut result = format!("{}Word SHA1:{}\n", indent_str, hash_short);
        for atom in self.atoms.iter() {
            result.push_str(&atom.to_string_with_indent(indent + 2));
            result.push('\n');
        }
        result
    }
}

impl Hashable for ComparisonUnitWord {
    fn hash(&self) -> &str {
        &self.sha1_hash
    }
}

/// Group-level comparison unit - hierarchical grouping
/// Corresponds to C# ComparisonUnitGroup (WmlComparer.cs:8608)
#[derive(Debug, Clone)]
pub struct ComparisonUnitGroup {
    /// Type of this group
    pub group_type: ComparisonUnitGroupType,
    /// Contents - can be words or nested groups
    pub contents: ComparisonUnitGroupContents,
    /// SHA1 hash (from first atom's ancestor)
    pub sha1_hash: String,
    /// Correlated SHA1 hash (pre-computed for block-level matching)
    pub correlated_sha1_hash: Option<String>,
    /// Structure SHA1 hash
    pub structure_sha1_hash: Option<String>,
    /// Correlation status
    pub correlation_status: ComparisonCorrelationStatus,
    /// Level in the hierarchy (0 = outermost)
    pub level: usize,
    /// Cached total count of atoms in this group (recursive)
    pub atom_count: usize,
}

/// Contents of a comparison unit group
#[derive(Debug, Clone)]
pub enum ComparisonUnitGroupContents {
    Words(Vec<ComparisonUnitWord>),
    Groups(Vec<ComparisonUnitGroup>),
}

impl ComparisonUnitGroup {
    /// Create a new group from words
    pub fn from_words(
        words: Vec<ComparisonUnitWord>,
        group_type: ComparisonUnitGroupType,
        level: usize,
    ) -> Self {
        let sha1_hash = compute_sha1_concat(words.iter().map(|w| w.sha1_hash.as_str()));
        
        let atom_count: usize = words.iter().map(|w| w.atoms.len()).sum();

        // Get correlated SHA1 hash from the first atom's ancestors
        let correlated_sha1_hash = words.first()
            .and_then(|w| w.atoms.first())
            .and_then(|atom| {
                let ancestor_name = match group_type {
                    ComparisonUnitGroupType::Table => "tbl",
                    ComparisonUnitGroupType::Row => "tr",
                    ComparisonUnitGroupType::Cell => "tc",
                    ComparisonUnitGroupType::Paragraph => "p",
                    ComparisonUnitGroupType::Textbox => "txbxContent",
                };

                atom.ancestor_elements.iter()
                    .find(|a| a.local_name == ancestor_name)
                    .and_then(|a| {
                        a.attributes.iter()
                            .find(|attr| attr.name.local_name == "CorrelatedSHA1Hash")
                            .map(|attr| attr.value.clone())
                    })
            });

        Self {
            group_type,
            contents: ComparisonUnitGroupContents::Words(words),
            sha1_hash,
            correlated_sha1_hash,
            structure_sha1_hash: None,
            correlation_status: ComparisonCorrelationStatus::Nil,
            level,
            atom_count,
        }
    }

    /// Create a new group from nested groups
    pub fn from_groups(
        groups: Vec<ComparisonUnitGroup>,
        group_type: ComparisonUnitGroupType,
        level: usize,
    ) -> Self {
        let sha1_hash = compute_sha1_concat(groups.iter().map(|g| g.sha1_hash.as_str()));
        
        let atom_count: usize = groups.iter().map(|g| g.atom_count).sum();

        // Get correlated SHA1 hash from the first group's first atom's ancestors
        let correlated_sha1_hash = groups.first()
            .and_then(|g| g.first_atom())
            .and_then(|atom| {
                let ancestor_name = match group_type {
                    ComparisonUnitGroupType::Table => "tbl",
                    ComparisonUnitGroupType::Row => "tr",
                    ComparisonUnitGroupType::Cell => "tc",
                    ComparisonUnitGroupType::Paragraph => "p",
                    ComparisonUnitGroupType::Textbox => "txbxContent",
                };

                atom.ancestor_elements.iter()
                    .find(|a| a.local_name == ancestor_name)
                    .and_then(|a| {
                        a.attributes.iter()
                            .find(|attr| attr.name.local_name == "CorrelatedSHA1Hash")
                            .map(|attr| attr.value.clone())
                    })
            });

        Self {
            group_type,
            contents: ComparisonUnitGroupContents::Groups(groups),
            sha1_hash,
            correlated_sha1_hash,
            structure_sha1_hash: None,
            correlation_status: ComparisonCorrelationStatus::Nil,
            level,
            atom_count,
        }
    }

    /// Get all descendant atoms
    pub fn descendant_atoms(&self) -> Vec<&ComparisonUnitAtom> {
        let mut atoms = Vec::with_capacity(self.atom_count);
        self.collect_atoms(&mut atoms);
        atoms
    }

    fn collect_atoms<'a>(&'a self, atoms: &mut Vec<&'a ComparisonUnitAtom>) {
        match &self.contents {
            ComparisonUnitGroupContents::Words(words) => {
                for word in words {
                    for atom in word.atoms.iter() {
                        atoms.push(atom);
                    }
                }
            }
            ComparisonUnitGroupContents::Groups(groups) => {
                for group in groups {
                    group.collect_atoms(atoms);
                }
            }
        }
    }

    /// Get the count of descendant content atoms
    pub fn descendant_atom_count(&self) -> usize {
        self.atom_count
    }

    /// Get contents as a vector of ComparisonUnit
    /// Used for cell content flattening in P0-3b
    pub fn contents_as_units(&self) -> Vec<ComparisonUnit> {
        match &self.contents {
            ComparisonUnitGroupContents::Words(words) => {
                words.iter().map(|w| ComparisonUnit::Word(w.clone())).collect()
            }
            ComparisonUnitGroupContents::Groups(groups) => {
                groups.iter().map(|g| ComparisonUnit::Group(g.clone())).collect()
            }
        }
    }

    /// Get the first atom in this group
    pub fn first_atom(&self) -> Option<&ComparisonUnitAtom> {
        match &self.contents {
            ComparisonUnitGroupContents::Words(words) => {
                words.first().and_then(|w| w.first_atom())
            }
            ComparisonUnitGroupContents::Groups(groups) => {
                groups.first().and_then(|g| g.first_atom())
            }
        }
    }

    /// Get the first comparison unit atom from a comparison unit (group or word)
    /// Corresponds to C# GetFirstComparisonUnitAtomOfGroup (WmlComparer.cs:8642)
    pub fn get_first_comparison_unit_atom_of_group(unit: &ComparisonUnit) -> Option<ComparisonUnitAtom> {
        match unit {
            ComparisonUnit::Word(w) => w.atoms.first().cloned(),
            ComparisonUnit::Group(g) => {
                match &g.contents {
                    ComparisonUnitGroupContents::Words(words) => {
                        words.first().and_then(|w| w.atoms.first().cloned())
                    }
                    ComparisonUnitGroupContents::Groups(groups) => {
                        groups.first().and_then(|g| {
                            Self::get_first_comparison_unit_atom_of_group(&ComparisonUnit::Group(g.clone()))
                        })
                    }
                }
            }
        }
    }

    /// Format as string with indentation
    /// Corresponds to C# ToString(int indent) (WmlComparer.cs:8661)
    pub fn to_string_with_indent(&self, indent: usize) -> String {
        let indent_str = " ".repeat(indent);
        let mut result = format!(
            "{}Group Type: {} SHA1:{}\n",
            indent_str, self.group_type, self.sha1_hash
        );
        
        match &self.contents {
            ComparisonUnitGroupContents::Words(words) => {
                for word in words {
                    result.push_str(&word.to_string_with_indent(indent + 2));
                }
            }
            ComparisonUnitGroupContents::Groups(groups) => {
                for group in groups {
                    result.push_str(&group.to_string_with_indent(indent + 2));
                }
            }
        }
        
        result
    }
}

impl Hashable for ComparisonUnitGroup {
    fn hash(&self) -> &str {
        // Use correlated hash if available (for block-level matching)
        self.correlated_sha1_hash.as_deref().unwrap_or(&self.sha1_hash)
    }
}

/// Compute SHA1 hash of a string
fn compute_sha1(content: &str) -> String {
    let mut hasher = Sha1::new();
    hasher.update(content.as_bytes());
    let result = hasher.finalize();
    format!("{:x}", result)
}

fn compute_sha1_concat<'a>(parts: impl IntoIterator<Item = &'a str>) -> String {
    let mut hasher = Sha1::new();
    for part in parts {
        hasher.update(part.as_bytes());
    }
    let result = hasher.finalize();
    format!("{:x}", result)
}

pub fn generate_unid() -> String {
    uuid::Uuid::new_v4().as_simple().to_string()
}

pub struct WordSeparatorSettings {
    pub word_separators: Vec<char>,
}

impl Default for WordSeparatorSettings {
    fn default() -> Self {
        Self {
            word_separators: vec![
                ' ', '-', ')', '(', ';', ',', 
                '（', '）', '，', '、', '；', '。', '：', '的',
            ],
        }
    }
}

impl WordSeparatorSettings {
    /// Check if a character is a word separator
    pub fn is_word_separator(&self, ch: char) -> bool {
        ch.is_whitespace() || self.word_separators.contains(&ch)
    }
}

/// Word break element names (elements that cause word breaks)
const WORD_BREAK_ELEMENTS: &[&str] = &[
    "pPr", "tab", "br", "continuationSeparator", "cr", "dayLong", "dayShort",
    "drawing", "pict", "endnoteRef", "footnoteRef", "monthLong", "monthShort",
    "noBreakHyphen", "object", "ptab", "separator", "sym", "yearLong", "yearShort",
    "oMathPara", "oMath", "footnoteReference", "endnoteReference",
];

/// Comparison grouping element names
const COMPARISON_GROUPING_ELEMENTS: &[&str] = &["p", "tbl", "tr", "tc", "txbxContent"];

/// Check if an element is a word break element
fn is_word_break_element(local_name: &str) -> bool {
    WORD_BREAK_ELEMENTS.contains(&local_name)
}

/// Check if an element is a comparison grouping element
fn is_comparison_grouping_element(local_name: &str) -> bool {
    COMPARISON_GROUPING_ELEMENTS.contains(&local_name)
}

/// Intermediate structure for word grouping
struct AtomWithGroupingKey {
    key: usize,
    atom: ComparisonUnitAtom,
}

/// Word with hierarchical grouping key
struct WordWithHierarchy {
    word: ComparisonUnitWord,
    hierarchy: Vec<String>, // e.g., ["p:abc123", "tc:def456"]
}

/// Get the comparison unit list from atoms
/// 
/// This is a faithful port of C# GetComparisonUnitList (WmlComparer.cs:7292)
/// 
/// Steps:
/// 1. Assign grouping keys to atoms based on word separators
/// 2. Group adjacent atoms with same key into words
/// 3. Extract hierarchical grouping from first atom's ancestors
/// 4. Recursively group into ComparisonUnitGroup hierarchy
pub fn get_comparison_unit_list(
    atoms: Vec<ComparisonUnitAtom>,
    settings: &WordSeparatorSettings,
) -> Vec<ComparisonUnit> {
    if atoms.is_empty() {
        return Vec::new();
    }

    // Step 1: Assign grouping keys using Rollup logic
    let grouping_keys = assign_grouping_keys(&atoms, settings);
    let atoms_with_keys: Vec<_> = atoms
        .into_iter()
        .zip(grouping_keys)
        .map(|(atom, key)| AtomWithGroupingKey { key, atom })
        .collect();

    // Step 2: Group adjacent atoms with same key into words
    let words_with_hierarchy = group_into_words(atoms_with_keys);

    // Step 3 & 4: Build hierarchical structure
    get_hierarchical_comparison_units(words_with_hierarchy, 0)
}

/// Assign grouping keys to atoms (Rollup logic from C#)
fn assign_grouping_keys(
    atoms: &[ComparisonUnitAtom],
    settings: &WordSeparatorSettings,
) -> Vec<usize> {
    let mut result = Vec::with_capacity(atoms.len());
    let mut next_index = 0usize;

    for (i, atom) in atoms.iter().enumerate() {
        let key = match &atom.content_element {
            ContentElement::Text(ch) => {
                if *ch == '.' || *ch == ',' {
                    // Special case: . and , in numbers stay in same word
                    let before_is_digit = if i > 0 {
                        matches!(&atoms[i - 1].content_element, ContentElement::Text(c) if c.is_ascii_digit())
                    } else {
                        false
                    };
                    let after_is_digit = if i + 1 < atoms.len() {
                        matches!(&atoms[i + 1].content_element, ContentElement::Text(c) if c.is_ascii_digit())
                    } else {
                        false
                    };

                    if before_is_digit || after_is_digit {
                        next_index // Keep in same word
                    } else {
                        // Punctuation is its own word
                        next_index += 1;
                        let key = next_index;
                        next_index += 1;
                        key
                    }
                } else if is_chinese_character(*ch) || settings.is_word_separator(*ch) {
                    // Chinese characters and word separators are their own words
                    next_index += 1;
                    let key = next_index;
                    next_index += 1;
                    key
                } else {
                    // Regular character stays in current word
                    next_index
                }
            }
            ContentElement::ParagraphProperties => {
                // pPr is a word break element
                next_index += 1;
                let key = next_index;
                next_index += 1;
                key
            }
            _ => {
                // Check if it's a word break element based on content type
                let is_word_break = matches!(
                    &atom.content_element,
                    ContentElement::Break
                        | ContentElement::Tab
                        | ContentElement::Drawing { .. }
                        | ContentElement::Picture { .. }
                        | ContentElement::Math { .. }
                        | ContentElement::FootnoteReference { .. }
                        | ContentElement::EndnoteReference { .. }
                        | ContentElement::Symbol { .. }
                        | ContentElement::Object { .. }
                        | ContentElement::FieldBegin
                        | ContentElement::FieldEnd
                );

                if is_word_break {
                    next_index += 1;
                    let key = next_index;
                    next_index += 1;
                    key
                } else {
                    next_index
                }
            }
        };

        result.push(key);
    }

    result
}

/// Check if character is in Chinese character range (CJK Unified Ideographs)
fn is_chinese_character(ch: char) -> bool {
    let code = ch as u32;
    (0x4e00..=0x9fff).contains(&code)
}

/// Group adjacent atoms with same key into words, adding hierarchy info
fn group_into_words(atoms_with_keys: Vec<AtomWithGroupingKey>) -> Vec<WordWithHierarchy> {
    if atoms_with_keys.is_empty() {
        return Vec::new();
    }

    let mut result = Vec::new();
    let mut current_key = atoms_with_keys[0].key;
    let mut current_atoms = Vec::new();

    for atom_with_key in atoms_with_keys {
        if atom_with_key.key != current_key {
            // Flush current word
            if !current_atoms.is_empty() {
                let hierarchy = extract_hierarchy(&current_atoms[0]);
                let word = ComparisonUnitWord::new(current_atoms);
                result.push(WordWithHierarchy { word, hierarchy });
                current_atoms = Vec::new();
            }
            current_key = atom_with_key.key;
        }
        current_atoms.push(atom_with_key.atom);
    }

    // Flush final word
    if !current_atoms.is_empty() {
        let hierarchy = extract_hierarchy(&current_atoms[0]);
        let word = ComparisonUnitWord::new(current_atoms);
        result.push(WordWithHierarchy { word, hierarchy });
    }

    result
}

/// Extract hierarchical grouping array from atom's ancestors
fn extract_hierarchy(atom: &ComparisonUnitAtom) -> Vec<String> {
    atom.ancestor_elements
        .iter()
        .filter(|a| is_comparison_grouping_element(&a.local_name))
        .map(|a| format!("{}:{}", a.local_name, a.unid))
        .collect()
}

/// Recursively build hierarchical comparison units
fn get_hierarchical_comparison_units(
    words: Vec<WordWithHierarchy>,
    level: usize,
) -> Vec<ComparisonUnit> {
    if words.is_empty() {
        return Vec::new();
    }

    let mut result = Vec::new();
    let mut iter = words.into_iter().peekable();

    while let Some(first) = iter.next() {
        let key = get_hierarchy_key(&first.hierarchy, level).map(|value| value.to_string());
        let mut group = Vec::new();
        group.push(first);

        while let Some(next) = iter.peek() {
            if get_hierarchy_key(&next.hierarchy, level) == key.as_deref() {
                group.push(iter.next().expect("peeked item should exist"));
            } else {
                break;
            }
        }

        match key {
            None => {
                for word in group {
                    result.push(ComparisonUnit::Word(word.word));
                }
            }
            Some(key) => {
                let group_type = parse_group_type(&key);
                let child_units = get_hierarchical_comparison_units(group, level + 1);
                let group_unit = ComparisonUnitGroup::from_comparison_units(child_units, group_type, level);
                result.push(ComparisonUnit::Group(group_unit));
            }
        }
    }

    result
}

/// Get hierarchy key at a specific level, or None if beyond hierarchy depth
fn get_hierarchy_key(hierarchy: &[String], level: usize) -> Option<&str> {
    hierarchy.get(level).map(|value| value.as_str())
}

/// Parse group type from hierarchy key (e.g., "p:abc123" -> Paragraph)
fn parse_group_type(key: &str) -> ComparisonUnitGroupType {
    let element_name = key.split(':').next().unwrap_or("");
    match element_name {
        "p" => ComparisonUnitGroupType::Paragraph,
        "tbl" => ComparisonUnitGroupType::Table,
        "tr" => ComparisonUnitGroupType::Row,
        "tc" => ComparisonUnitGroupType::Cell,
        "txbxContent" => ComparisonUnitGroupType::Textbox,
        _ => ComparisonUnitGroupType::Paragraph, // Default fallback
    }
}

/// A comparison unit that can be either a word or a group
#[derive(Debug, Clone)]
pub enum ComparisonUnit {
    Word(ComparisonUnitWord),
    Group(ComparisonUnitGroup),
}

impl ComparisonUnit {
    /// Get the SHA1 hash of this unit
    pub fn hash(&self) -> &str {
        match self {
            ComparisonUnit::Word(w) => &w.sha1_hash,
            ComparisonUnit::Group(g) => g.hash(),
        }
    }

    /// Get correlation status
    pub fn correlation_status(&self) -> ComparisonCorrelationStatus {
        match self {
            ComparisonUnit::Word(w) => w.correlation_status,
            ComparisonUnit::Group(g) => g.correlation_status,
        }
    }

    /// Set correlation status
    pub fn set_correlation_status(&mut self, status: ComparisonCorrelationStatus) {
        match self {
            ComparisonUnit::Word(w) => w.correlation_status = status,
            ComparisonUnit::Group(g) => g.correlation_status = status,
        }
    }

    /// Get all descendants (including groups and words)
    /// Corresponds to C# Descendants() (WmlComparer.cs:8163)
    pub fn descendants(&self) -> Vec<ComparisonUnit> {
        let mut result = Vec::new();
        self.descendants_internal(&mut result);
        result
    }

    fn descendants_internal(&self, result: &mut Vec<ComparisonUnit>) {
        match self {
            ComparisonUnit::Word(_) => {
            }
            ComparisonUnit::Group(g) => {
                match &g.contents {
                    ComparisonUnitGroupContents::Words(words) => {
                        for word in words {
                            result.push(ComparisonUnit::Word(word.clone()));
                        }
                    }
                    ComparisonUnitGroupContents::Groups(groups) => {
                        for group in groups {
                            result.push(ComparisonUnit::Group(group.clone()));
                            ComparisonUnit::Group(group.clone()).descendants_internal(result);
                        }
                    }
                }
            }
        }
    }

    /// Get all descendant content atoms
    /// Corresponds to C# DescendantContentAtoms() (WmlComparer.cs:8170)
    pub fn descendant_content_atoms(&self) -> Vec<ComparisonUnitAtom> {
        let mut result = Vec::new();
        for unit in self.descendants() {
            if let ComparisonUnit::Word(word) = unit {
                result.extend(word.atoms.iter().cloned());
            }
        }
        result
    }

    /// Get all descendant atoms (as references)
    pub fn descendant_atoms(&self) -> Vec<&ComparisonUnitAtom> {
        match self {
            ComparisonUnit::Word(w) => w.atoms.iter().collect(),
            ComparisonUnit::Group(g) => g.descendant_atoms(),
        }
    }

    /// Get count of descendant content atoms
    /// Corresponds to C# DescendantContentAtomsCount (WmlComparer.cs:8177)
    pub fn descendant_content_atoms_count(&self) -> usize {
        match self {
            ComparisonUnit::Word(w) => w.atoms.len(),
            ComparisonUnit::Group(g) => g.descendant_atom_count(),
        }
    }

    /// Format as string with indentation
    /// Corresponds to C# ToString(int indent) (WmlComparer.cs:8198)
    pub fn to_string_with_indent(&self, indent: usize) -> String {
        match self {
            ComparisonUnit::Word(w) => w.to_string_with_indent(indent),
            ComparisonUnit::Group(g) => g.to_string_with_indent(indent),
        }
    }

    /// Format a list of comparison units as a string
    /// Corresponds to C# ComparisonUnitListToString (WmlComparer.cs:8200)
    pub fn comparison_unit_list_to_string(units: &[ComparisonUnit]) -> String {
        let mut result = String::from("Dump Comparision Unit List To String\n");
        for unit in units {
            result.push_str(&unit.to_string_with_indent(2));
            result.push('\n');
        }
        result
    }

    /// Check if this is a group
    pub fn is_group(&self) -> bool {
        matches!(self, ComparisonUnit::Group(_))
    }

    /// Get as group if it is one
    pub fn as_group(&self) -> Option<&ComparisonUnitGroup> {
        match self {
            ComparisonUnit::Group(g) => Some(g),
            _ => None,
        }
    }

    /// Get as word if it is one
    pub fn as_word(&self) -> Option<&ComparisonUnitWord> {
        match self {
            ComparisonUnit::Word(w) => Some(w),
            _ => None,
        }
    }
}

impl Hashable for ComparisonUnit {
    fn hash(&self) -> &str {
        ComparisonUnit::hash(self)
    }
}

impl ComparisonUnitGroup {
    /// Create a group from a list of comparison units
    pub fn from_comparison_units(
        units: Vec<ComparisonUnit>,
        group_type: ComparisonUnitGroupType,
        level: usize,
    ) -> Self {
        // Separate words and groups
        let mut words = Vec::new();
        let mut groups = Vec::new();

        for unit in units {
            match unit {
                ComparisonUnit::Word(w) => words.push(w),
                ComparisonUnit::Group(g) => groups.push(g),
            }
        }

        // If we have groups, create a group-containing group
        // If we only have words, create a word-containing group
        if !groups.is_empty() {
            // If we have both, we need to wrap words in paragraph groups
            if !words.is_empty() {
                // This is an edge case - wrap loose words in a pseudo-group
                let word_group = ComparisonUnitGroup::from_words(
                    words,
                    ComparisonUnitGroupType::Paragraph,
                    level + 1,
                );
                groups.insert(0, word_group);
            }
            ComparisonUnitGroup::from_groups(groups, group_type, level)
        } else {
            ComparisonUnitGroup::from_words(words, group_type, level)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_content_element_hash() {
        let text = ContentElement::Text('a');
        assert_eq!(text.hash_string(), "ta");

        let ppr = ContentElement::ParagraphProperties;
        assert_eq!(ppr.hash_string(), "pPr");

        let drawing = ContentElement::Drawing {
            hash: "abc123".to_string(),
        };
        assert_eq!(drawing.hash_string(), "drawingabc123");
    }

    #[test]
    fn test_atom_creation() {
        use crate::xml::arena::XmlDocument;
        use crate::xml::node::XmlNodeData;
        use crate::wml::settings::WmlComparerSettings;
        
        let mut doc = XmlDocument::new();
        let node = doc.add_root(XmlNodeData::Text("test".to_string()));
        
        let settings = WmlComparerSettings::default();
        let atom = ComparisonUnitAtom::new(
            ContentElement::Text('H'),
            vec![AncestorInfo {
                node_id: node,
                namespace: None,
                local_name: "p".to_string(),
                unid: "abc123".to_string(),
                attributes: Arc::new(Vec::new()),
                has_merged_cells: false,
            }],
            "main",
            &settings,
        );

        assert!(!atom.sha1_hash.is_empty());
        assert_eq!(atom.paragraph_unid(), Some("abc123"));
    }

    #[test]
    fn test_word_creation() {
        use crate::wml::settings::WmlComparerSettings;
        let settings = WmlComparerSettings::default();
        let atoms = vec![
            ComparisonUnitAtom::new(ContentElement::Text('H'), vec![], "main", &settings),
            ComparisonUnitAtom::new(ContentElement::Text('i'), vec![], "main", &settings),
        ];

        let word = ComparisonUnitWord::new(atoms);
        assert_eq!(word.text(), "Hi");
        assert!(!word.sha1_hash.is_empty());
    }

    #[test]
    fn test_group_creation() {
        use crate::wml::settings::WmlComparerSettings;
        let settings = WmlComparerSettings::default();
        let atoms = vec![
            ComparisonUnitAtom::new(ContentElement::Text('A'), vec![], "main", &settings),
        ];
        let word = ComparisonUnitWord::new(atoms);
        let group = ComparisonUnitGroup::from_words(
            vec![word],
            ComparisonUnitGroupType::Paragraph,
            0,
        );

        assert_eq!(group.group_type, ComparisonUnitGroupType::Paragraph);
        assert_eq!(group.descendant_atom_count(), 1);
    }

    #[test]
    fn test_generate_unid() {
        let unid1 = generate_unid();
        let unid2 = generate_unid();
        
        assert_eq!(unid1.len(), 32);
        assert_eq!(unid2.len(), 32);
        assert_ne!(unid1, unid2);
    }

    // ===== Canonical Atom Model Tests (RUST-1) =====

    #[test]
    fn test_content_type_classification() {
        // Text
        let text = ContentElement::Text('a');
        assert_eq!(text.content_type(), ContentType::Text);
        
        // Paragraph mark
        let ppr = ContentElement::ParagraphProperties;
        assert_eq!(ppr.content_type(), ContentType::ParagraphMark);
        
        // Drawing
        let drawing = ContentElement::Drawing { hash: "abc".to_string() };
        assert_eq!(drawing.content_type(), ContentType::Drawing);
        
        // Field
        let field = ContentElement::FieldBegin;
        assert_eq!(field.content_type(), ContentType::Field);
        
        // Picture
        let pict = ContentElement::Picture { hash: "xyz".to_string() };
        assert_eq!(pict.content_type(), ContentType::Picture);
    }

    #[test]
    fn test_content_element_local_name() {
        assert_eq!(ContentElement::Text('a').local_name(), "t");
        assert_eq!(ContentElement::ParagraphProperties.local_name(), "pPr");
        assert_eq!(ContentElement::Drawing { hash: "x".to_string() }.local_name(), "drawing");
        assert_eq!(ContentElement::Picture { hash: "x".to_string() }.local_name(), "pict");
        assert_eq!(ContentElement::Break.local_name(), "br");
        assert_eq!(ContentElement::Tab.local_name(), "tab");
        assert_eq!(ContentElement::FieldBegin.local_name(), "fldChar");
        assert_eq!(ContentElement::Symbol { font: "f".to_string(), char_code: "c".to_string() }.local_name(), "sym");
    }

    #[test]
    fn test_content_element_text_value() {
        // Text element has single character value
        assert_eq!(ContentElement::Text('H').text_value(), "H");
        assert_eq!(ContentElement::Text(' ').text_value(), " ");
        
        // ParagraphProperties has empty text value
        assert_eq!(ContentElement::ParagraphProperties.text_value(), "");
        
        // Drawing has hash as text value
        assert_eq!(ContentElement::Drawing { hash: "abc123".to_string() }.text_value(), "abc123");
        
        // Field characters have type as value
        assert_eq!(ContentElement::FieldBegin.text_value(), "begin");
        assert_eq!(ContentElement::FieldEnd.text_value(), "end");
    }

    #[test]
    fn test_hash_string_matches_c_sharp_format() {
        // C# format: localName + textValue
        
        // Text: "t" + char
        let text = ContentElement::Text('H');
        assert_eq!(text.hash_string(), "tH");
        
        // ParagraphProperties: "pPr" + "" = "pPr"
        let ppr = ContentElement::ParagraphProperties;
        assert_eq!(ppr.hash_string(), "pPr");
        
        // Drawing: "drawing" + hash
        let drawing = ContentElement::Drawing { hash: "abc123".to_string() };
        assert_eq!(drawing.hash_string(), "drawingabc123");
        
        // Symbol: "sym" + font:char
        let sym = ContentElement::Symbol { font: "Symbol".to_string(), char_code: "F020".to_string() };
        assert_eq!(sym.hash_string(), "symSymbol:F020");
    }

    #[test]
    fn test_hash_string_with_case_insensitive() {
        let text_lower = ContentElement::Text('a');
        let text_upper = ContentElement::Text('A');
        
        // Without case insensitivity, different hashes
        let settings_default = WmlComparerSettings::default().with_case_insensitive(false);
        assert_ne!(
            text_lower.hash_string_with_settings(&settings_default),
            text_upper.hash_string_with_settings(&settings_default)
        );
        
        // With case insensitivity, same hashes (both uppercase)
        let settings_ci = WmlComparerSettings::default().with_case_insensitive(true);
        assert_eq!(
            text_lower.hash_string_with_settings(&settings_ci),
            text_upper.hash_string_with_settings(&settings_ci)
        );
        assert_eq!(text_lower.hash_string_with_settings(&settings_ci), "tA");
    }

    #[test]
    fn test_hash_string_with_space_conflation() {
        let text_space = ContentElement::Text(' ');
        let text_nbsp = ContentElement::Text('\u{00a0}'); // Non-breaking space
        
        // Default settings have conflation enabled
        let settings = WmlComparerSettings::default();
        
        // Regular space gets converted to NBSP
        let space_hash = text_space.hash_string_with_settings(&settings);
        let nbsp_hash = text_nbsp.hash_string_with_settings(&settings);
        
        // After conflation, both should have NBSP character
        assert_eq!(space_hash, "t\u{00a0}");
        assert_eq!(space_hash, nbsp_hash);
    }

    #[test]
    fn test_atom_identity_hash() {
        let settings = WmlComparerSettings::default();
        
        let atom = ComparisonUnitAtom::new(
            ContentElement::Text('H'),
            vec![],
            "main",
            &settings,
        );
        
        // identity_hash() returns the sha1_hash
        assert_eq!(atom.identity_hash(), atom.sha1_hash.as_str());
        assert!(!atom.identity_hash().is_empty());
    }

    #[test]
    fn test_atom_equality_based_on_identity_hash() {
        let settings = WmlComparerSettings::default();
        
        // Same content element → same identity hash → equal
        let atom1 = ComparisonUnitAtom::new(
            ContentElement::Text('H'),
            vec![],
            "main",
            &settings,
        );
        let atom2 = ComparisonUnitAtom::new(
            ContentElement::Text('H'),
            vec![],
            "main",
            &settings,
        );
        
        assert_eq!(atom1, atom2); // PartialEq based on sha1_hash
        
        // Different content element → different identity hash → not equal
        let atom3 = ComparisonUnitAtom::new(
            ContentElement::Text('X'),
            vec![],
            "main",
            &settings,
        );
        
        assert_ne!(atom1, atom3);
    }

    #[test]
    fn test_atom_with_precomputed_hash() {
        let atom = ComparisonUnitAtom::new_with_hash(
            ContentElement::Text('H'),
            vec![],
            "main",
            "precomputed_hash_abc123".to_string(),
        );
        
        assert_eq!(atom.identity_hash(), "precomputed_hash_abc123");
    }

    #[test]
    fn test_drawings_in_same_scheme_as_text() {
        let settings = WmlComparerSettings::default();
        
        // Both text and drawings use the same hash computation scheme:
        // SHA1(localName + textValue)
        
        let text_atom = ComparisonUnitAtom::new(
            ContentElement::Text('H'),
            vec![],
            "main",
            &settings,
        );
        
        let drawing_atom = ComparisonUnitAtom::new(
            ContentElement::Drawing { hash: "image_hash_123".to_string() },
            vec![],
            "main",
            &settings,
        );
        
        // Both have non-empty identity hashes
        assert!(!text_atom.identity_hash().is_empty());
        assert!(!drawing_atom.identity_hash().is_empty());
        
        // Different content → different hashes
        assert_ne!(text_atom.identity_hash(), drawing_atom.identity_hash());
        
        // Same drawing hash → same identity hash
        let drawing_atom2 = ComparisonUnitAtom::new(
            ContentElement::Drawing { hash: "image_hash_123".to_string() },
            vec![],
            "main",
            &settings,
        );
        
        assert_eq!(drawing_atom, drawing_atom2);
    }

    #[test]
    fn test_formatting_signature_methods() {
        let settings = WmlComparerSettings::default();
        
        let mut atom = ComparisonUnitAtom::new(
            ContentElement::Text('H'),
            vec![],
            "main",
            &settings,
        );
        
        // Initially None
        assert!(atom.formatting_signature().is_none());
        
        // Set formatting signature
        atom.set_formatting_signature(Some("<w:rPr><w:b/></w:rPr>".to_string()));
        assert_eq!(atom.formatting_signature(), Some("<w:rPr><w:b/></w:rPr>"));
    }

    #[test]
    fn test_ancestor_unids_methods() {
        let settings = WmlComparerSettings::default();
        
        let mut atom = ComparisonUnitAtom::new(
            ContentElement::Text('H'),
            vec![],
            "main",
            &settings,
        );
        
        // Initially empty
        assert!(atom.ancestor_unids().is_empty());
        
        // Set ancestor unids
        atom.set_ancestor_unids(vec!["unid1".to_string(), "unid2".to_string()]);
        assert_eq!(atom.ancestor_unids(), &["unid1".to_string(), "unid2".to_string()]);
    }
}
