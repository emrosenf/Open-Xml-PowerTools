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
//! - Each unit has a SHA1 hash for comparison
//! - Atoms track ancestor elements with Unids for tree reconstruction
//! - Groups have CorrelatedSHA1Hash for efficient block-level matching

use crate::util::lcs::Hashable;
use indextree::NodeId;
use sha1::{Digest, Sha1};
use std::fmt;

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
    /// Local name of the element (e.g., "p", "tbl", "tr", "tc")
    pub local_name: String,
    /// Unique ID (Unid) for this element - used for correlation
    pub unid: String,
    /// Attributes from the ancestor element (for reconstruction)
    pub attributes: Vec<crate::xml::xname::XAttribute>,
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
    /// Get hash string for this content element
    pub fn hash_string(&self) -> String {
        match self {
            ContentElement::Text(ch) => format!("t{}", ch),
            ContentElement::ParagraphProperties => "pPr".to_string(),
            ContentElement::RunProperties => "rPr".to_string(),
            ContentElement::Break => "br".to_string(),
            ContentElement::Tab => "tab".to_string(),
            ContentElement::Drawing { hash } => format!("drawing{}", hash),
            ContentElement::Picture { hash } => format!("pict{}", hash),
            ContentElement::Math { hash } => format!("math{}", hash),
            ContentElement::FootnoteReference { id } => format!("footnoteRef{}", id),
            ContentElement::EndnoteReference { id } => format!("endnoteRef{}", id),
            ContentElement::TextboxStart => "txbxStart".to_string(),
            ContentElement::TextboxEnd => "txbxEnd".to_string(),
            ContentElement::FieldBegin => "fldBegin".to_string(),
            ContentElement::FieldSeparator => "fldSep".to_string(),
            ContentElement::FieldEnd => "fldEnd".to_string(),
            ContentElement::SimpleField { instruction } => format!("fldSimple{}", instruction),
            ContentElement::Symbol { font, char_code } => format!("sym{}:{}", font, char_code),
            ContentElement::Object { hash } => format!("object{}", hash),
            ContentElement::Unknown { name } => format!("unknown{}", name),
        }
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
#[derive(Debug, Clone)]
pub struct ComparisonUnitAtom {
    /// The content element this atom represents
    pub content_element: ContentElement,
    /// SHA1 hash of the content
    pub sha1_hash: String,
    /// Ancestor elements from body to this element (body → leaf order)
    pub ancestor_elements: Vec<AncestorInfo>,
    /// Correlation status
    pub correlation_status: ComparisonCorrelationStatus,
    /// Formatting signature (for TrackFormattingChanges)
    pub formatting_signature: Option<String>,
    /// Normalized run properties (for format change detection)
    pub normalized_rpr: Option<String>,
    /// Part name this atom belongs to (main, footnotes, endnotes)
    pub part_name: String,
    
    // Fields for "before" document tracking (Equal/FormatChanged atoms)
    /// Content element from "before" document
    pub content_element_before: Option<ContentElement>,
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

impl ComparisonUnitAtom {
    /// Create a new atom with the given content element and ancestors
    pub fn new(
        content_element: ContentElement,
        ancestor_elements: Vec<AncestorInfo>,
        part_name: &str,
    ) -> Self {
        let hash_string = content_element.hash_string();
        let sha1_hash = compute_sha1(&hash_string);

        Self {
            content_element,
            sha1_hash,
            ancestor_elements,
            correlation_status: ComparisonCorrelationStatus::Nil,
            formatting_signature: None,
            normalized_rpr: None,
            part_name: part_name.to_string(),
            content_element_before: None,
            comparison_unit_atom_before: None,
            ancestor_elements_before: None,
            part_before: None,
            rev_track_element: None,
            formatting_change_rpr_before: None,
            ancestor_unids: Vec::new(),
            formatting_change_rpr_before_signature: None,
        }
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
    pub atoms: Vec<ComparisonUnitAtom>,
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
        let hash_input: String = atoms.iter().map(|a| a.sha1_hash.as_str()).collect();
        let sha1_hash = compute_sha1(&hash_input);

        Self {
            atoms,
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
        for atom in &self.atoms {
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
        let sha1_hash = if let Some(first_word) = words.first() {
            first_word.sha1_hash.clone()
        } else {
            compute_sha1("")
        };

        Self {
            group_type,
            contents: ComparisonUnitGroupContents::Words(words),
            sha1_hash,
            correlated_sha1_hash: None,
            structure_sha1_hash: None,
            correlation_status: ComparisonCorrelationStatus::Nil,
            level,
        }
    }

    /// Create a new group from nested groups
    pub fn from_groups(
        groups: Vec<ComparisonUnitGroup>,
        group_type: ComparisonUnitGroupType,
        level: usize,
    ) -> Self {
        let sha1_hash = if let Some(first_group) = groups.first() {
            first_group.sha1_hash.clone()
        } else {
            compute_sha1("")
        };

        Self {
            group_type,
            contents: ComparisonUnitGroupContents::Groups(groups),
            sha1_hash,
            correlated_sha1_hash: None,
            structure_sha1_hash: None,
            correlation_status: ComparisonCorrelationStatus::Nil,
            level,
        }
    }

    /// Get all descendant atoms
    pub fn descendant_atoms(&self) -> Vec<&ComparisonUnitAtom> {
        let mut atoms = Vec::new();
        self.collect_atoms(&mut atoms);
        atoms
    }

    fn collect_atoms<'a>(&'a self, atoms: &mut Vec<&'a ComparisonUnitAtom>) {
        match &self.contents {
            ComparisonUnitGroupContents::Words(words) => {
                for word in words {
                    for atom in &word.atoms {
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
        match &self.contents {
            ComparisonUnitGroupContents::Words(words) => {
                words.iter().map(|w| w.atoms.len()).sum()
            }
            ComparisonUnitGroupContents::Groups(groups) => {
                groups.iter().map(|g| g.descendant_atom_count()).sum()
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
    let atoms_with_keys = assign_grouping_keys(&atoms, settings);

    // Step 2: Group adjacent atoms with same key into words
    let words_with_hierarchy = group_into_words(atoms_with_keys);

    // Step 3 & 4: Build hierarchical structure
    get_hierarchical_comparison_units(&words_with_hierarchy, 0)
}

/// Assign grouping keys to atoms (Rollup logic from C#)
fn assign_grouping_keys(
    atoms: &[ComparisonUnitAtom],
    settings: &WordSeparatorSettings,
) -> Vec<AtomWithGroupingKey> {
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

        result.push(AtomWithGroupingKey {
            key,
            atom: atom.clone(),
        });
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
    words: &[WordWithHierarchy],
    level: usize,
) -> Vec<ComparisonUnit> {
    if words.is_empty() {
        return Vec::new();
    }

    // Group by hierarchy key at current level
    let mut result = Vec::new();
    let mut current_key = get_hierarchy_key(&words[0].hierarchy, level);
    let mut current_group: Vec<&WordWithHierarchy> = Vec::new();

    for word in words {
        let key = get_hierarchy_key(&word.hierarchy, level);
        if key != current_key {
            // Process current group
            result.extend(process_hierarchy_group(&current_group, level, &current_key));
            current_group.clear();
            current_key = key;
        }
        current_group.push(word);
    }

    // Process final group
    result.extend(process_hierarchy_group(&current_group, level, &current_key));

    result
}

/// Get hierarchy key at a specific level, or empty string if beyond hierarchy depth
fn get_hierarchy_key(hierarchy: &[String], level: usize) -> String {
    if level < hierarchy.len() {
        hierarchy[level].clone()
    } else {
        String::new()
    }
}

/// Process a group of words at a hierarchy level
fn process_hierarchy_group(
    words: &[&WordWithHierarchy],
    level: usize,
    key: &str,
) -> Vec<ComparisonUnit> {
    if words.is_empty() {
        return Vec::new();
    }

    if key.is_empty() {
        // No more hierarchy - return words directly
        words
            .iter()
            .map(|w| ComparisonUnit::Word(w.word.clone()))
            .collect()
    } else {
        // Create a group and recurse
        let group_type = parse_group_type(key);
        
        // Collect owned copies for recursion
        let owned_words: Vec<WordWithHierarchy> = words
            .iter()
            .map(|w| WordWithHierarchy {
                word: w.word.clone(),
                hierarchy: w.hierarchy.clone(),
            })
            .collect();
        
        let child_units = get_hierarchical_comparison_units(&owned_words, level + 1);
        let group = ComparisonUnitGroup::from_comparison_units(child_units, group_type, level);
        
        vec![ComparisonUnit::Group(group)]
    }
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
        self.descendants()
            .into_iter()
            .filter_map(|cu| match cu {
                ComparisonUnit::Word(w) => Some(w.atoms),
                _ => None,
            })
            .flatten()
            .collect()
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
        
        let mut doc = XmlDocument::new();
        let node = doc.add_root(XmlNodeData::Text("test".to_string()));
        
        let atom = ComparisonUnitAtom::new(
            ContentElement::Text('H'),
            vec![AncestorInfo {
                node_id: node,
                local_name: "p".to_string(),
                unid: "abc123".to_string(),
                attributes: vec![],
            }],
            "main",
        );

        assert!(!atom.sha1_hash.is_empty());
        assert_eq!(atom.paragraph_unid(), Some("abc123"));
    }

    #[test]
    fn test_word_creation() {
        let atoms = vec![
            ComparisonUnitAtom::new(ContentElement::Text('H'), vec![], "main"),
            ComparisonUnitAtom::new(ContentElement::Text('i'), vec![], "main"),
        ];

        let word = ComparisonUnitWord::new(atoms);
        assert_eq!(word.text(), "Hi");
        assert!(!word.sha1_hash.is_empty());
    }

    #[test]
    fn test_group_creation() {
        let atoms = vec![
            ComparisonUnitAtom::new(ContentElement::Text('A'), vec![], "main"),
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
}
