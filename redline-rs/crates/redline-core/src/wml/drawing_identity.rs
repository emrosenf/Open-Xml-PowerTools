//! Drawing Canonical Identity - SHA1 Hash Pipeline
//!
//! This module implements stable canonical identity for drawings/images via SHA1 hash
//! of their content, rather than using relationship IDs which can change across saves.
//!
//! ## Problem Solved
//! Relationship IDs (rId5, rId6) change across saves even if the image is identical.
//! The current XML-only hashing approach causes false positives (same image detected as different).
//!
//! ## C# Reference (WmlComparer.cs lines 4340-4420)
//! The algorithm for drawings:
//! 1. For `w:drawing` and `w:pict` elements, check for textbox content first
//! 2. If has `w:txbxContent`, recurse into textbox (special handling)
//! 3. Otherwise, for image content:
//!    - Find `a:blip` element with `r:embed` attribute (DrawingML)
//!    - Or find `v:imagedata` with `r:id` attribute (VML)
//!    - Resolve relationship ID to ImagePart
//!    - Hash the binary stream (SHA1)
//!    - Use hash as canonical identity
//!
//! ## Implementation Notes
//! - This module provides functions to compute stable drawing identity
//! - Textboxes are handled specially - their content is hashed, not structure
//! - Missing relationships result in a sentinel value rather than errors

use crate::package::OoxmlPackage;
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{A, PT, R, V, W, WP14};
use crate::xml::node::XmlNodeData;
use indextree::NodeId;
use sha1::{Digest, Sha1};

/// Sentinel value for drawings with missing or unresolvable relationships
const MISSING_RELATIONSHIP_SENTINEL: &str = "MISSING_RELATIONSHIP";

/// Sentinel value for textbox drawings (content-based identity)
const TEXTBOX_DRAWING_PREFIX: &str = "TEXTBOX:";

/// Compute the canonical identity hash for a drawing element.
///
/// This is the main entry point for drawing identity computation.
///
/// # Arguments
/// * `doc` - The XML document containing the drawing
/// * `drawing_node` - The NodeId of the w:drawing or w:pict element
/// * `package` - The OOXML package for resolving relationships (optional)
/// * `source_part` - The path to the source part (e.g., "word/document.xml")
///
/// # Returns
/// A stable SHA1 hash string representing the drawing's canonical identity.
///
/// # Algorithm
/// 1. Check for textbox content (w:txbxContent)
/// 2. If textbox, hash the textbox content
/// 3. Otherwise, find image reference (a:blip or v:imagedata)
/// 4. Resolve relationship and hash image binary content
/// 5. Fall back to XML structure hash if relationship cannot be resolved
pub fn compute_drawing_identity(
    doc: &XmlDocument,
    drawing_node: NodeId,
    package: Option<&OoxmlPackage>,
    source_part: &str,
) -> String {
    // Step 1: Check for textbox content
    let txbx_contents = find_textbox_contents(doc, drawing_node);
    if !txbx_contents.is_empty() {
        return compute_textbox_identity(doc, &txbx_contents);
    }

    // Step 2: Try to find image reference and resolve it
    if let Some(pkg) = package {
        // Try DrawingML (a:blip with r:embed)
        if let Some(embed_id) = find_blip_embed(doc, drawing_node) {
            if let Some(hash) = resolve_and_hash_image(pkg, source_part, &embed_id) {
                return hash;
            }
        }

        // Try VML (v:imagedata with r:id or o:relid)
        if let Some(rel_id) = find_vml_imagedata_relid(doc, drawing_node) {
            if let Some(hash) = resolve_and_hash_image(pkg, source_part, &rel_id) {
                return hash;
            }
        }
    }

    // Step 3: Fall back to XML structure hash (existing behavior)
    compute_xml_structure_hash(doc, drawing_node)
}

/// Find all w:txbxContent descendants within a drawing element.
fn find_textbox_contents(doc: &XmlDocument, node: NodeId) -> Vec<NodeId> {
    let mut txbx_nodes = Vec::new();

    for desc in doc.descendants(node) {
        if let Some(data) = doc.get(desc) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "txbxContent" {
                    txbx_nodes.push(desc);
                }
            }
        }
    }

    txbx_nodes
}

/// Compute identity hash for textbox content.
///
/// Textboxes are special - we hash their text content rather than their
/// XML structure to handle the case where the container (VML vs DrawingML)
/// differs but the content is the same.
fn compute_textbox_identity(doc: &XmlDocument, txbx_contents: &[NodeId]) -> String {
    let mut hasher = Sha1::new();
    hasher.update(TEXTBOX_DRAWING_PREFIX.as_bytes());

    for &txbx_node in txbx_contents {
        hash_textbox_content_recursive(doc, txbx_node, &mut hasher);
    }

    let result = hasher.finalize();
    format!("{:x}", result)
}

/// Recursively hash textbox content, extracting text and structure.
fn hash_textbox_content_recursive(doc: &XmlDocument, node: NodeId, hasher: &mut Sha1) {
    let Some(data) = doc.get(node) else { return };

    match data {
        XmlNodeData::Element { name, .. } => {
            // Include element name in hash for structural identity
            hasher.update(name.local_name.as_bytes());

            // Recurse into children
            for child in doc.children(node) {
                hash_textbox_content_recursive(doc, child, hasher);
            }
        }
        XmlNodeData::Text(text) => {
            hasher.update(text.as_bytes());
        }
        _ => {}
    }
}

/// Find a:blip element with r:embed attribute (DrawingML image reference).
fn find_blip_embed(doc: &XmlDocument, node: NodeId) -> Option<String> {
    for desc in doc.descendants(node) {
        if let Some(data) = doc.get(desc) {
            if let Some(name) = data.name() {
                // Look for a:blip element
                if name.namespace.as_deref() == Some(A::NS) && name.local_name == "blip" {
                    if let Some(attrs) = data.attributes() {
                        // Look for r:embed attribute
                        for attr in attrs {
                            if attr.name.local_name == "embed"
                                && attr.name.namespace.as_deref() == Some(R::NS)
                            {
                                return Some(attr.value.clone());
                            }
                        }
                    }
                }
            }
        }
    }
    None
}

/// Find v:imagedata element with r:id or o:relid attribute (VML image reference).
fn find_vml_imagedata_relid(doc: &XmlDocument, node: NodeId) -> Option<String> {
    for desc in doc.descendants(node) {
        if let Some(data) = doc.get(desc) {
            if let Some(name) = data.name() {
                // Look for v:imagedata element
                if name.namespace.as_deref() == Some(V::NS) && name.local_name == "imagedata" {
                    if let Some(attrs) = data.attributes() {
                        // Look for r:id attribute first
                        for attr in attrs {
                            if attr.name.local_name == "id"
                                && attr.name.namespace.as_deref() == Some(R::NS)
                            {
                                return Some(attr.value.clone());
                            }
                        }
                        // Fallback to o:relid
                        for attr in attrs {
                            if attr.name.local_name == "relid" {
                                return Some(attr.value.clone());
                            }
                        }
                    }
                }
            }
        }
    }
    None
}

/// Resolve a relationship ID to an image part and hash its binary content.
fn resolve_and_hash_image(
    package: &OoxmlPackage,
    source_part: &str,
    rel_id: &str,
) -> Option<String> {
    // Get the relationships file path for the source part
    let rels_path = get_relationships_path(source_part);

    // Try to parse the relationships and find the target
    let rels_content = package.get_part(&rels_path)?;
    let target = parse_relationship_target(rels_content, rel_id)?;

    // Resolve relative path to absolute part path
    let image_path = resolve_part_path(source_part, &target);

    // Get the image bytes and hash them
    let image_bytes = package.get_part(&image_path)?;
    Some(sha1_hash_bytes(image_bytes))
}

/// Get the path to the .rels file for a given part.
fn get_relationships_path(part_path: &str) -> String {
    // For "word/document.xml" -> "word/_rels/document.xml.rels"
    if let Some(last_slash) = part_path.rfind('/') {
        let dir = &part_path[..last_slash];
        let file = &part_path[last_slash + 1..];
        format!("{}/_rels/{}.rels", dir, file)
    } else {
        format!("_rels/{}.rels", part_path)
    }
}

/// Parse a relationships XML file and find the target for a given ID.
fn parse_relationship_target(rels_content: &[u8], rel_id: &str) -> Option<String> {
    // Simple XML parsing for relationship files
    // Format: <Relationship Id="rIdX" Target="target/path.ext" Type="..." />
    let content = std::str::from_utf8(rels_content).ok()?;

    // Find the relationship with matching Id
    for line in content.lines() {
        if line.contains(&format!("Id=\"{}\"", rel_id)) {
            // Extract Target attribute
            if let Some(target_start) = line.find("Target=\"") {
                let start = target_start + 8;
                let rest = &line[start..];
                if let Some(end) = rest.find('"') {
                    return Some(rest[..end].to_string());
                }
            }
        }
    }

    None
}

/// Resolve a relative path to an absolute part path.
fn resolve_part_path(source_part: &str, relative_target: &str) -> String {
    // Handle absolute paths (starting with /)
    if relative_target.starts_with('/') {
        return relative_target.trim_start_matches('/').to_string();
    }

    // Get directory of source part
    let source_dir = if let Some(last_slash) = source_part.rfind('/') {
        &source_part[..last_slash]
    } else {
        ""
    };

    // Combine and normalize
    if source_dir.is_empty() {
        relative_target.to_string()
    } else {
        format!("{}/{}", source_dir, relative_target)
    }
}

/// Compute SHA1 hash of binary content.
fn sha1_hash_bytes(bytes: &[u8]) -> String {
    let mut hasher = Sha1::new();
    hasher.update(bytes);
    let result = hasher.finalize();
    format!("{:x}", result)
}

/// Compute XML structure hash as fallback.
/// This is the existing behavior - hash the XML structure excluding PT attributes.
fn compute_xml_structure_hash(doc: &XmlDocument, node: NodeId) -> String {
    use crate::xml::namespaces::PT;

    let mut hasher = Sha1::new();
    hash_xml_element_recursive(doc, node, &mut hasher, &PT::Unid(), &PT::SHA1Hash());
    let result = hasher.finalize();
    format!("{:x}", result)
}

/// Recursively hash XML element structure.
///
/// Skips:
/// - PT namespace attributes (Unid, SHA1Hash)
/// - Relationship attributes (r:embed, r:id, r:link) which vary between saves
/// - WP14 namespace attributes (anchorId, editId) which are edit-tracking IDs
fn hash_xml_element_recursive(
    doc: &XmlDocument,
    node: NodeId,
    hasher: &mut Sha1,
    pt_unid: &crate::xml::xname::XName,
    pt_sha1: &crate::xml::xname::XName,
) {
    let Some(data) = doc.get(node) else { return };

    match data {
        XmlNodeData::Element { name, attributes } => {
            hasher.update(name.local_name.as_bytes());
            let mut filtered_attrs: Vec<_> = attributes
                .iter()
                .filter(|attr| {
                    if attr.name.namespace.as_deref() == Some("http://www.w3.org/2000/xmlns/") {
                        return false;
                    }
                    if &attr.name == pt_unid
                        || &attr.name == pt_sha1
                        || attr.name.namespace.as_deref() == Some(PT::NS)
                    {
                        return false;
                    }
                    if attr.name.namespace.as_deref() == Some(R::NS) {
                        return false;
                    }
                    if attr.name.namespace.as_deref() == Some(WP14::NS) {
                        return false;
                    }
                    if attr.name.namespace.is_none()
                        && matches!(
                            attr.name.local_name.as_str(),
                            "ObjectID" | "ShapeID" | "id" | "type"
                        )
                    {
                        return false;
                    }
                    true
                })
                .collect();

            filtered_attrs.sort_by(|a, b| {
                let a_ns = a.name.namespace.as_deref().unwrap_or("");
                let b_ns = b.name.namespace.as_deref().unwrap_or("");
                (a_ns, a.name.local_name.as_str(), a.value.as_str()).cmp(&(
                    b_ns,
                    b.name.local_name.as_str(),
                    b.value.as_str(),
                ))
            });

            for attr in filtered_attrs {
                // Skip PT namespace attributes
                hasher.update(attr.name.local_name.as_bytes());
                hasher.update(attr.value.as_bytes());
            }
            for child in doc.children(node) {
                hash_xml_element_recursive(doc, child, hasher, pt_unid, pt_sha1);
            }
        }
        XmlNodeData::Text(text) => {
            hasher.update(text.as_bytes());
        }
        _ => {}
    }
}

/// Check if a drawing element contains textbox content.
pub fn has_textbox_content(doc: &XmlDocument, drawing_node: NodeId) -> bool {
    !find_textbox_contents(doc, drawing_node).is_empty()
}

/// Get information about a drawing element for debugging.
pub fn get_drawing_info(doc: &XmlDocument, drawing_node: NodeId) -> DrawingInfo {
    let has_textbox = has_textbox_content(doc, drawing_node);
    let blip_embed = find_blip_embed(doc, drawing_node);
    let vml_relid = find_vml_imagedata_relid(doc, drawing_node);

    DrawingInfo {
        has_textbox,
        blip_embed,
        vml_relid,
    }
}

/// Information about a drawing element.
#[derive(Debug, Clone)]
pub struct DrawingInfo {
    /// Whether the drawing contains textbox content
    pub has_textbox: bool,
    /// The r:embed attribute from a:blip (DrawingML)
    pub blip_embed: Option<String>,
    /// The r:id or o:relid attribute from v:imagedata (VML)
    pub vml_relid: Option<String>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::arena::XmlDocument;
    use crate::xml::node::XmlNodeData;
    use crate::xml::xname::XName;

    #[test]
    fn test_sha1_hash_bytes() {
        let content = b"test content";
        let hash = sha1_hash_bytes(content);
        // SHA1 produces 40 hex characters
        assert_eq!(hash.len(), 40);
        // Verify it's consistent
        assert_eq!(sha1_hash_bytes(content), sha1_hash_bytes(content));
        // Different content produces different hash
        assert_ne!(
            sha1_hash_bytes(b"test content"),
            sha1_hash_bytes(b"other content")
        );
    }

    #[test]
    fn test_get_relationships_path() {
        assert_eq!(
            get_relationships_path("word/document.xml"),
            "word/_rels/document.xml.rels"
        );
        assert_eq!(
            get_relationships_path("word/media/image1.png"),
            "word/media/_rels/image1.png.rels"
        );
    }

    #[test]
    fn test_resolve_part_path() {
        // Relative path
        assert_eq!(
            resolve_part_path("word/document.xml", "media/image1.png"),
            "word/media/image1.png"
        );
        // Absolute path
        assert_eq!(
            resolve_part_path("word/document.xml", "/word/media/image1.png"),
            "word/media/image1.png"
        );
    }

    #[test]
    fn test_find_textbox_contents() {
        let mut doc = XmlDocument::new();
        let drawing = doc.add_root(XmlNodeData::element(W::drawing()));
        let shape = doc.add_child(drawing, XmlNodeData::element(XName::new(V::NS, "shape")));
        let textbox = doc.add_child(shape, XmlNodeData::element(XName::new(V::NS, "textbox")));
        let txbx = doc.add_child(textbox, XmlNodeData::element(W::txbxContent()));
        doc.add_child(txbx, XmlNodeData::Text("Hello".to_string()));

        let txbx_contents = find_textbox_contents(&doc, drawing);
        assert_eq!(txbx_contents.len(), 1);
        assert_eq!(txbx_contents[0], txbx);
    }

    #[test]
    fn test_find_blip_embed() {
        let mut doc = XmlDocument::new();
        let drawing = doc.add_root(XmlNodeData::element(W::drawing()));
        let blip = doc.add_child(
            drawing,
            XmlNodeData::Element {
                name: XName::new(A::NS, "blip"),
                attributes: vec![crate::xml::xname::XAttribute {
                    name: XName::new(R::NS, "embed"),
                    value: "rId5".to_string(),
                }],
            },
        );

        let embed = find_blip_embed(&doc, drawing);
        assert_eq!(embed, Some("rId5".to_string()));
    }

    #[test]
    fn test_find_vml_imagedata_relid() {
        let mut doc = XmlDocument::new();
        let pict = doc.add_root(XmlNodeData::element(W::pict()));
        let imagedata = doc.add_child(
            pict,
            XmlNodeData::Element {
                name: XName::new(V::NS, "imagedata"),
                attributes: vec![crate::xml::xname::XAttribute {
                    name: XName::new(R::NS, "id"),
                    value: "rId7".to_string(),
                }],
            },
        );

        let relid = find_vml_imagedata_relid(&doc, pict);
        assert_eq!(relid, Some("rId7".to_string()));
    }

    #[test]
    fn test_parse_relationship_target() {
        let rels_content = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.jpg"/>
</Relationships>"#;

        assert_eq!(
            parse_relationship_target(rels_content, "rId1"),
            Some("media/image1.png".to_string())
        );
        assert_eq!(
            parse_relationship_target(rels_content, "rId2"),
            Some("media/image2.jpg".to_string())
        );
        assert_eq!(parse_relationship_target(rels_content, "rId99"), None);
    }

    #[test]
    fn test_textbox_identity_consistent() {
        let mut doc = XmlDocument::new();
        let drawing = doc.add_root(XmlNodeData::element(W::drawing()));
        let txbx = doc.add_child(drawing, XmlNodeData::element(W::txbxContent()));
        let para = doc.add_child(txbx, XmlNodeData::element(W::p()));
        doc.add_child(para, XmlNodeData::Text("Hello World".to_string()));

        let txbx_contents = find_textbox_contents(&doc, drawing);
        let hash1 = compute_textbox_identity(&doc, &txbx_contents);
        let hash2 = compute_textbox_identity(&doc, &txbx_contents);

        assert_eq!(hash1, hash2);
        assert!(!hash1.is_empty());
    }

    #[test]
    fn test_compute_drawing_identity_textbox() {
        let mut doc = XmlDocument::new();
        let drawing = doc.add_root(XmlNodeData::element(W::drawing()));
        let txbx = doc.add_child(drawing, XmlNodeData::element(W::txbxContent()));
        let para = doc.add_child(txbx, XmlNodeData::element(W::p()));
        doc.add_child(para, XmlNodeData::Text("Hello World".to_string()));

        // Without package, should compute textbox identity
        let hash = compute_drawing_identity(&doc, drawing, None, "word/document.xml");

        assert!(!hash.is_empty());
        assert_eq!(hash.len(), 40); // SHA1 hex length
    }

    #[test]
    fn test_compute_drawing_identity_fallback_to_xml() {
        let mut doc = XmlDocument::new();
        let drawing = doc.add_root(XmlNodeData::element(W::drawing()));
        let blip = doc.add_child(
            drawing,
            XmlNodeData::Element {
                name: XName::new(A::NS, "blip"),
                attributes: vec![crate::xml::xname::XAttribute {
                    name: XName::new(R::NS, "embed"),
                    value: "rId5".to_string(),
                }],
            },
        );

        // Without package, should fall back to XML structure hash
        let hash = compute_drawing_identity(&doc, drawing, None, "word/document.xml");

        assert!(!hash.is_empty());
        assert_eq!(hash.len(), 40); // SHA1 hex length
    }

    #[test]
    fn test_has_textbox_content() {
        let mut doc1 = XmlDocument::new();
        let drawing1 = doc1.add_root(XmlNodeData::element(W::drawing()));
        let txbx = doc1.add_child(drawing1, XmlNodeData::element(W::txbxContent()));

        assert!(has_textbox_content(&doc1, drawing1));

        let mut doc2 = XmlDocument::new();
        let drawing2 = doc2.add_root(XmlNodeData::element(W::drawing()));
        let blip = doc2.add_child(
            drawing2,
            XmlNodeData::Element {
                name: XName::new(A::NS, "blip"),
                attributes: vec![],
            },
        );

        assert!(!has_textbox_content(&doc2, drawing2));
    }

    #[test]
    fn test_get_drawing_info() {
        let mut doc = XmlDocument::new();
        let drawing = doc.add_root(XmlNodeData::element(W::drawing()));
        let blip = doc.add_child(
            drawing,
            XmlNodeData::Element {
                name: XName::new(A::NS, "blip"),
                attributes: vec![crate::xml::xname::XAttribute {
                    name: XName::new(R::NS, "embed"),
                    value: "rId5".to_string(),
                }],
            },
        );

        let info = get_drawing_info(&doc, drawing);
        assert!(!info.has_textbox);
        assert_eq!(info.blip_embed, Some("rId5".to_string()));
        assert_eq!(info.vml_relid, None);
    }
}
