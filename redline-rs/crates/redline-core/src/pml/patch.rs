//! Patching PML documents with changes.
//!
//! This module provides functionality to apply or revert changes to PowerPoint presentations.

use crate::error::{RedlineError, Result};
use crate::pml::{PmlChange, PmlChangeType, PmlDocument};
use crate::xml::namespaces::P;
use crate::xml::{XAttribute, XName, XmlDocument, XmlNodeData};
use indextree::NodeId;

/// Apply a set of changes to a base PowerPoint document.
///
/// Returns the modified document as bytes.
pub fn apply_pml_changes(base_doc: &[u8], changes: &[PmlChange]) -> Result<Vec<u8>> {
    let mut doc = PmlDocument::from_bytes(base_doc)?;

    // Group changes by slide index to minimize I/O
    // Note: Slide insertions/deletions change indices, so we must be careful with order.
    // For now, we assume changes are sorted or we process them sequentially.
    // Actually, PmlChange contains `slide_index`. If we insert a slide, subsequent indices shift.
    // The diff engine produces changes based on the *comparison result*, which usually maps
    // new indices.
    // If we are applying changes to the *old* document to make it look like the *new* one,
    // we should process them in order.

    // However, PmlChange structure is flattened.
    // Let's implement a basic applicator that handles slide content changes.
    // Slide insertion/deletion/move requires updating presentation.xml and relationships,
    // which is complex.
    // For the MVP "Accept/Reject" in a viewer, we often just want to update text/shapes on existing slides.

    // For this implementation, we focus on content changes within slides (Text, Shapes).
    // Structural changes (Slide Insert/Delete) are harder to "patch" individually without
    // re-implementing the entire presentation structure logic.

    for change in changes {
        match change.change_type {
            PmlChangeType::TextChanged
            | PmlChangeType::TextFormattingChanged
            | PmlChangeType::ShapeMoved
            | PmlChangeType::ShapeResized
            | PmlChangeType::ShapeRotated => {
                apply_slide_content_change(&mut doc, change)?;
            }
            _ => {
                // Other changes not yet supported for patching
            }
        }
    }

    doc.to_bytes()
}

/// Revert a set of changes from a modified PowerPoint document.
pub fn revert_pml_changes(result_doc: &[u8], changes: &[PmlChange]) -> Result<Vec<u8>> {
    let inverse_changes: Vec<PmlChange> = changes.iter().map(|c| invert_change(c)).collect();
    apply_pml_changes(result_doc, &inverse_changes)
}

fn invert_change(change: &PmlChange) -> PmlChange {
    let mut inverse = change.clone();

    // Swap old/new values
    inverse.old_value = change.new_value.clone();
    inverse.new_value = change.old_value.clone();

    inverse.old_x = change.new_x;
    inverse.new_x = change.old_x;

    inverse.old_y = change.new_y;
    inverse.new_y = change.old_y;

    inverse.old_cx = change.new_cx;
    inverse.new_cx = change.old_cx;

    inverse.old_cy = change.new_cy;
    inverse.new_cy = change.old_cy;

    // Invert text changes if present
    if let Some(text_changes) = &change.text_changes {
        let mut inverted_text_changes = Vec::new();
        for tc in text_changes {
            let mut inverted_tc = tc.clone();
            inverted_tc.old_text = tc.new_text.clone();
            inverted_tc.new_text = tc.old_text.clone();
            inverted_text_changes.push(inverted_tc);
        }
        inverse.text_changes = Some(inverted_text_changes);
    }

    inverse
}

fn apply_slide_content_change(doc: &mut PmlDocument, change: &PmlChange) -> Result<()> {
    let slide_index = change
        .slide_index
        .ok_or_else(|| RedlineError::InvalidPackage {
            message: "Missing slide index".to_string(),
        })?;

    // 0-based index to 1-based ID assumption or lookup?
    // We need to map slide index to part name.
    // PmlDocument doesn't expose a helper for this yet.
    // We can assume standard PPTX structure or implement a helper.
    // Standard: ppt/slides/slide{index+1}.xml (usually).
    // But better to look up via presentation.xml relationships.

    // For now, let's implement a helper in this module to find slide part by index.
    let slide_part = get_slide_part_by_index(doc, slide_index)?;

    let mut slide_xml = doc.package().get_xml_part(&slide_part)?;
    let root = slide_xml
        .root()
        .ok_or_else(|| RedlineError::InvalidPackage {
            message: "Slide has no root".to_string(),
        })?;

    // Find the shape
    let shape_id = change
        .shape_id
        .as_deref()
        .ok_or_else(|| RedlineError::InvalidPackage {
            message: "Missing shape ID".to_string(),
        })?;

    let shape_node = find_shape_by_id(&slide_xml, root, shape_id);

    if let Some(shape_node) = shape_node {
        match change.change_type {
            PmlChangeType::TextChanged | PmlChangeType::TextFormattingChanged => {
                // For simple text replacement:
                // If text_changes detail is present, use it?
                // Or just use new_value?
                // Using new_value is coarser but easier.
                if let Some(new_text) = &change.new_value {
                    update_shape_text(&mut slide_xml, shape_node, new_text);
                }
            }
            PmlChangeType::ShapeMoved | PmlChangeType::ShapeResized => {
                if let (Some(x), Some(y), Some(cx), Some(cy)) =
                    (change.new_x, change.new_y, change.new_cx, change.new_cy)
                {
                    update_shape_transform(
                        &mut slide_xml,
                        shape_node,
                        x as i64,
                        y as i64,
                        cx as i64,
                        cy as i64,
                    );
                }
            }
            _ => {}
        }
    }

    doc.package_mut().put_xml_part(&slide_part, &slide_xml)?;

    Ok(())
}

fn get_slide_part_by_index(doc: &PmlDocument, index: usize) -> Result<String> {
    // Basic implementation: assume "ppt/slides/slide{index+1}.xml"
    // A robust implementation would parse presentation.xml and relationships.
    // Given scope, we try the standard path.
    let path = format!("ppt/slides/slide{}.xml", index + 1);
    if doc.package().get_part(&path).is_some() {
        Ok(path)
    } else {
        // Fallback: try finding it via relationships?
        // Let's just return error if not found at standard location for now.
        Err(RedlineError::InvalidPackage {
            message: format!("Could not find slide part for index {}", index),
        })
    }
}

fn find_shape_by_id(doc: &XmlDocument, root: NodeId, shape_id: &str) -> Option<NodeId> {
    // DFS to find shape with ID
    // Shape IDs in PPTX are usually in cNvPr element: <p:cNvPr id="4" name="Title 1"/>

    for node in doc.descendants(root) {
        if let Some(data) = doc.get(node) {
            if let Some(name) = data.name() {
                if name.local_name == "cNvPr" {
                    if let Some(attrs) = data.attributes() {
                        if attrs
                            .iter()
                            .any(|a| a.name.local_name == "id" && a.value == shape_id)
                        {
                            // Found the non-visual props. The shape is the grandparent/parent container.
                            // usually: sp -> nvSpPr -> cNvPr
                            // We return the container (sp)
                            if let Some(nv_pr) = doc.parent(node) {
                                if let Some(shape) = doc.parent(nv_pr) {
                                    return Some(shape);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    None
}

fn update_shape_text(doc: &mut XmlDocument, shape_node: NodeId, new_text: &str) {
    // Locate txBody
    let tx_body = doc.descendants(shape_node).find(|&n| {
        doc.get(n).map_or(false, |d| {
            d.name().map_or(false, |name| name.local_name == "txBody")
        })
    });

    if let Some(tx_body) = tx_body {
        // Simple update: remove all existing paragraphs and add a new one with the text
        let children: Vec<_> = doc.children(tx_body).collect();
        for child in children {
            // Keep bodyPr and lstStyle?
            // Usually we want to replace <p> elements.
            if let Some(data) = doc.get(child) {
                if let Some(name) = data.name() {
                    if name.local_name == "p" {
                        doc.detach(child);
                    }
                }
            }
        }

        // Add new paragraph
        use crate::xml::namespaces::A;
        let p = doc.add_child(tx_body, XmlNodeData::element(XName::new(A::NS, "p")));
        let r = doc.add_child(p, XmlNodeData::element(XName::new(A::NS, "r")));
        let t = doc.add_child(r, XmlNodeData::element(XName::new(A::NS, "t")));
        doc.add_child(t, XmlNodeData::text(new_text));
    }
}

fn update_shape_transform(
    doc: &mut XmlDocument,
    shape_node: NodeId,
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
) {
    // Locate spPr -> xfrm
    let sp_pr = doc.children(shape_node).find(|&n| {
        doc.get(n).map_or(false, |d| {
            d.name().map_or(false, |name| name.local_name == "spPr")
        })
    });

    if let Some(sp_pr) = sp_pr {
        use crate::xml::namespaces::A;
        let xfrm_name = XName::new(A::NS, "xfrm");
        let xfrm = doc.children(sp_pr).find(|&n| {
            doc.get(n)
                .map_or(false, |d| d.name().map_or(false, |name| name == &xfrm_name))
        });

        if let Some(xfrm) = xfrm {
            // Update off (offset) and ext (extent)
            let off_name = XName::new(A::NS, "off");
            let ext_name = XName::new(A::NS, "ext");

            // Collect IDs first to avoid borrow issues
            let off_id = doc
                .children(xfrm)
                .find(|&n| doc.get(n).map_or(false, |d| d.name() == Some(&off_name)));
            let ext_id = doc
                .children(xfrm)
                .find(|&n| doc.get(n).map_or(false, |d| d.name() == Some(&ext_name)));

            if let Some(off) = off_id {
                doc.set_attribute(off, &XName::local("x"), &x.to_string());
                doc.set_attribute(off, &XName::local("y"), &y.to_string());
            }

            if let Some(ext) = ext_id {
                doc.set_attribute(ext, &XName::local("cx"), &cx.to_string());
                doc.set_attribute(ext, &XName::local("cy"), &cy.to_string());
            }
        }
    }
}
