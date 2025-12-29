//! Element Ordering Module
//!
//! This module provides functionality to order XML elements according to the OOXML standard.
//! The order is critical for document validity and proper rendering by Office applications.
//!
//! **CRITICAL C# Translation**
//! This ports `WmlOrderElementsPerStandard` from PtOpenXmlUtil.cs line 1361.
//!
//! Order dictionaries define the canonical sequence of child elements within:
//! - pPr (paragraph properties)
//! - rPr (run properties)
//! - tblPr (table properties)
//! - tcPr (table cell properties)
//! - tblBorders, tcBorders, pBdr (border properties)

use crate::xml::namespaces::W;
use crate::xml::xname::XName;
use once_cell::sync::Lazy;
use std::collections::HashMap;

const W14_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordml";

fn w(local: &str) -> XName {
    XName::new(W::NS, local)
}

fn w14(local: &str) -> XName {
    XName::new(W14_NS, local)
}

/// Paragraph Properties Element Order
/// Maps element names to their canonical order within w:pPr
/// From C#: Order_pPr dictionary (PtOpenXmlUtil.cs lines 1194-1232)
pub static ORDER_P_PR: Lazy<HashMap<XName, i32>> = Lazy::new(|| {
    let mut m = HashMap::new();
    m.insert(w("pStyle"), 10);
    m.insert(w("keepNext"), 20);
    m.insert(w("keepLines"), 30);
    m.insert(w("pageBreakBefore"), 40);
    m.insert(w("framePr"), 50);
    m.insert(w("widowControl"), 60);
    m.insert(w("numPr"), 70);
    m.insert(w("suppressLineNumbers"), 80);
    m.insert(w("pBdr"), 90);
    m.insert(w("shd"), 100);
    m.insert(w("tabs"), 120);
    m.insert(w("suppressAutoHyphens"), 130);
    m.insert(w("kinsoku"), 140);
    m.insert(w("wordWrap"), 150);
    m.insert(w("overflowPunct"), 160);
    m.insert(w("topLinePunct"), 170);
    m.insert(w("autoSpaceDE"), 180);
    m.insert(w("autoSpaceDN"), 190);
    m.insert(w("bidi"), 200);
    m.insert(w("adjustRightInd"), 210);
    m.insert(w("snapToGrid"), 220);
    m.insert(w("spacing"), 230);
    m.insert(w("ind"), 240);
    m.insert(w("contextualSpacing"), 250);
    m.insert(w("mirrorIndents"), 260);
    m.insert(w("suppressOverlap"), 270);
    m.insert(w("jc"), 280);
    m.insert(w("textDirection"), 290);
    m.insert(w("textAlignment"), 300);
    m.insert(w("textboxTightWrap"), 310);
    m.insert(w("outlineLvl"), 320);
    m.insert(w("divId"), 330);
    m.insert(w("cnfStyle"), 340);
    m.insert(w("rPr"), 350);
    m.insert(w("sectPr"), 360);
    m.insert(w("pPrChange"), 370);
    m
});

/// Run Properties Element Order
/// Maps element names to their canonical order within w:rPr
/// From C#: Order_rPr dictionary (PtOpenXmlUtil.cs lines 1234-1284)
pub static ORDER_R_PR: Lazy<HashMap<XName, i32>> = Lazy::new(|| {
    let mut m = HashMap::new();
    m.insert(w("moveFrom"), 5);
    m.insert(w("moveTo"), 7);
    m.insert(w("ins"), 10);
    m.insert(w("del"), 20);
    m.insert(w("rStyle"), 30);
    m.insert(w("rFonts"), 40);
    m.insert(w("b"), 50);
    m.insert(w("bCs"), 60);
    m.insert(w("i"), 70);
    m.insert(w("iCs"), 80);
    m.insert(w("caps"), 90);
    m.insert(w("smallCaps"), 100);
    m.insert(w("strike"), 110);
    m.insert(w("dstrike"), 120);
    m.insert(w("outline"), 130);
    m.insert(w("shadow"), 140);
    m.insert(w("emboss"), 150);
    m.insert(w("imprint"), 160);
    m.insert(w("noProof"), 170);
    m.insert(w("snapToGrid"), 180);
    m.insert(w("vanish"), 190);
    m.insert(w("webHidden"), 200);
    m.insert(w("color"), 210);
    m.insert(w("spacing"), 220);
    m.insert(w("w"), 230);
    m.insert(w("kern"), 240);
    m.insert(w("position"), 250);
    m.insert(w("sz"), 260);
    m.insert(w14("shadow"), 270);
    m.insert(w14("textOutline"), 280);
    m.insert(w14("textFill"), 290);
    m.insert(w14("scene3d"), 300);
    m.insert(w14("props3d"), 310);
    m.insert(w("szCs"), 320);
    m.insert(w("highlight"), 330);
    m.insert(w("u"), 340);
    m.insert(w("effect"), 350);
    m.insert(w("bdr"), 360);
    m.insert(w("shd"), 370);
    m.insert(w("fitText"), 380);
    m.insert(w("vertAlign"), 390);
    m.insert(w("rtl"), 400);
    m.insert(w("cs"), 410);
    m.insert(w("em"), 420);
    m.insert(w("lang"), 430);
    m.insert(w("eastAsianLayout"), 440);
    m.insert(w("specVanish"), 450);
    m.insert(w("oMath"), 460);
    m
});

/// Table Properties Element Order
/// From C#: Order_tblPr dictionary (PtOpenXmlUtil.cs lines 1286-1305)
pub static ORDER_TBL_PR: Lazy<HashMap<XName, i32>> = Lazy::new(|| {
    let mut m = HashMap::new();
    m.insert(w("tblStyle"), 10);
    m.insert(w("tblpPr"), 20);
    m.insert(w("tblOverlap"), 30);
    m.insert(w("bidiVisual"), 40);
    m.insert(w("tblStyleRowBandSize"), 50);
    m.insert(w("tblStyleColBandSize"), 60);
    m.insert(w("tblW"), 70);
    m.insert(w("jc"), 80);
    m.insert(w("tblCellSpacing"), 90);
    m.insert(w("tblInd"), 100);
    m.insert(w("tblBorders"), 110);
    m.insert(w("shd"), 120);
    m.insert(w("tblLayout"), 130);
    m.insert(w("tblCellMar"), 140);
    m.insert(w("tblLook"), 150);
    m.insert(w("tblCaption"), 160);
    m.insert(w("tblDescription"), 170);
    m
});

/// Table Cell Properties Element Order
/// From C#: Order_tcPr dictionary (PtOpenXmlUtil.cs lines 1319-1335)
pub static ORDER_TC_PR: Lazy<HashMap<XName, i32>> = Lazy::new(|| {
    let mut m = HashMap::new();
    m.insert(w("cnfStyle"), 10);
    m.insert(w("tcW"), 20);
    m.insert(w("gridSpan"), 30);
    m.insert(w("hMerge"), 40);
    m.insert(w("vMerge"), 50);
    m.insert(w("tcBorders"), 60);
    m.insert(w("shd"), 70);
    m.insert(w("noWrap"), 80);
    m.insert(w("tcMar"), 90);
    m.insert(w("textDirection"), 100);
    m.insert(w("tcFitText"), 110);
    m.insert(w("vAlign"), 120);
    m.insert(w("hideMark"), 130);
    m.insert(w("headers"), 140);
    m
});

/// Table Borders Element Order
/// From C#: Order_tblBorders dictionary (PtOpenXmlUtil.cs lines 1307-1317)
pub static ORDER_TBL_BORDERS: Lazy<HashMap<XName, i32>> = Lazy::new(|| {
    let mut m = HashMap::new();
    m.insert(w("top"), 10);
    m.insert(w("left"), 20);
    m.insert(w("start"), 30);
    m.insert(w("bottom"), 40);
    m.insert(w("right"), 50);
    m.insert(w("end"), 60);
    m.insert(w("insideH"), 70);
    m.insert(w("insideV"), 80);
    m
});

/// Table Cell Borders Element Order
/// From C#: Order_tcBorders dictionary (PtOpenXmlUtil.cs lines 1337-1349)
pub static ORDER_TC_BORDERS: Lazy<HashMap<XName, i32>> = Lazy::new(|| {
    let mut m = HashMap::new();
    m.insert(w("top"), 10);
    m.insert(w("start"), 20);
    m.insert(w("left"), 30);
    m.insert(w("bottom"), 40);
    m.insert(w("right"), 50);
    m.insert(w("end"), 60);
    m.insert(w("insideH"), 70);
    m.insert(w("insideV"), 80);
    m.insert(w("tl2br"), 90);
    m.insert(w("tr2bl"), 100);
    m
});

/// Paragraph Borders Element Order
/// From C#: Order_pBdr dictionary (PtOpenXmlUtil.cs lines 1351-1359)
pub static ORDER_P_BDR: Lazy<HashMap<XName, i32>> = Lazy::new(|| {
    let mut m = HashMap::new();
    m.insert(w("top"), 10);
    m.insert(w("left"), 20);
    m.insert(w("bottom"), 30);
    m.insert(w("right"), 40);
    m.insert(w("between"), 50);
    m.insert(w("bar"), 60);
    m
});

/// Returns the order priority for a given element name in the specified context
///
/// # Arguments
/// * `element_name` - The XName of the element to look up
/// * `order_map` - The ordering map for the specific context (pPr, rPr, etc.)
///
/// # Returns
/// The ordering priority (lower numbers come first), or 999 if not found
pub fn get_element_order(element_name: &XName, order_map: &HashMap<XName, i32>) -> i32 {
    *order_map.get(element_name).unwrap_or(&999)
}

use crate::xml::arena::XmlDocument;
use indextree::NodeId;

/// Order XML elements per OOXML standard
/// Port of C# WmlOrderElementsPerStandard (PtOpenXmlUtil.cs line 1361)
///
/// This function recursively reorders child elements within:
/// - w:pPr (paragraph properties)
/// - w:rPr (run properties)
/// - w:tblPr (table properties)
/// - w:tcPr (table cell properties)
/// - w:tblBorders, w:tcBorders, w:pBdr (border properties)
/// - w:p (paragraphs - ensures pPr comes first)
/// - w:r (runs - ensures rPr comes first)
pub fn order_elements_per_standard(doc: &mut XmlDocument) {
    if let Some(root) = doc.root() {
        order_element_recursive(doc, root);
    }
}

fn order_element_recursive(doc: &mut XmlDocument, node_id: NodeId) {
    // First, get the element name to determine what reordering is needed
    let element_name = doc.get(node_id).and_then(|d| d.name().cloned());

    // Recursively process all children first
    let children: Vec<NodeId> = doc.children(node_id).collect();
    for child in &children {
        order_element_recursive(doc, *child);
    }

    // Now reorder this node's children based on its type
    if let Some(name) = element_name {
        let local = name.local_name.as_str();
        let ns = name.namespace.as_deref();

        // Only process elements in the W namespace
        if ns != Some(W::NS) {
            return;
        }

        match local {
            "pPr" => reorder_children(doc, node_id, &ORDER_P_PR),
            "rPr" => reorder_children(doc, node_id, &ORDER_R_PR),
            "tblPr" => reorder_children(doc, node_id, &ORDER_TBL_PR),
            "tcPr" => reorder_children(doc, node_id, &ORDER_TC_PR),
            "tblBorders" => reorder_children(doc, node_id, &ORDER_TBL_BORDERS),
            "tcBorders" => reorder_children(doc, node_id, &ORDER_TC_BORDERS),
            "pBdr" => reorder_children(doc, node_id, &ORDER_P_BDR),
            "p" => {
                // For paragraphs, ensure pPr comes first
                ensure_first_child(doc, node_id, "pPr");
            }
            "r" => {
                // For runs, ensure rPr comes first
                ensure_first_child(doc, node_id, "rPr");
            }
            _ => {}
        }
    }
}

/// Reorder children of a node according to the given order map
fn reorder_children(doc: &mut XmlDocument, parent: NodeId, order_map: &HashMap<XName, i32>) {
    // Collect children with their sort keys
    let mut children_with_keys: Vec<(NodeId, i32)> = doc.children(parent)
        .map(|child| {
            let key = doc.get(child)
                .and_then(|d| d.name())
                .map(|name| get_element_order(name, order_map))
                .unwrap_or(999);
            (child, key)
        })
        .collect();

    // Check if already in order
    let mut is_sorted = true;
    for i in 1..children_with_keys.len() {
        if children_with_keys[i].1 < children_with_keys[i-1].1 {
            is_sorted = false;
            break;
        }
    }

    if is_sorted {
        return; // Already in correct order
    }

    // Sort by order key (stable sort to preserve relative order of same-priority elements)
    children_with_keys.sort_by_key(|(_, key)| *key);

    // Detach all children
    for (child, _) in &children_with_keys {
        doc.detach(*child);
    }

    // Re-attach in sorted order
    for (child, _) in children_with_keys {
        doc.reparent(parent, child);
    }
}

/// Ensure a specific element type comes first among children
fn ensure_first_child(doc: &mut XmlDocument, parent: NodeId, first_local_name: &str) {
    let children: Vec<NodeId> = doc.children(parent).collect();

    // Find the element that should be first
    let mut first_element = None;
    let mut first_element_idx = None;

    for (idx, child) in children.iter().enumerate() {
        if let Some(data) = doc.get(*child) {
            if let Some(name) = data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == first_local_name {
                    first_element = Some(*child);
                    first_element_idx = Some(idx);
                    break;
                }
            }
        }
    }

    // If the element exists and is not already first, move it
    if let (Some(elem), Some(idx)) = (first_element, first_element_idx) {
        if idx > 0 {
            // Detach and re-insert at beginning
            doc.detach(elem);

            // We need to insert before the first child
            // Since XmlDocument doesn't have insert_before, we detach all and re-add
            let remaining: Vec<NodeId> = doc.children(parent).collect();
            for child in &remaining {
                doc.detach(*child);
            }

            // Add the first element
            doc.reparent(parent, elem);

            // Add the rest
            for child in remaining {
                doc.reparent(parent, child);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_order_p_pr() {
        assert_eq!(get_element_order(&w("pStyle"), &ORDER_P_PR), 10);
        assert_eq!(get_element_order(&w("rPr"), &ORDER_P_PR), 350);
        assert_eq!(get_element_order(&w("pPrChange"), &ORDER_P_PR), 370);
    }

    #[test]
    fn test_order_r_pr() {
        assert_eq!(get_element_order(&w("rStyle"), &ORDER_R_PR), 30);
        assert_eq!(get_element_order(&w("b"), &ORDER_R_PR), 50);
        assert_eq!(get_element_order(&w("i"), &ORDER_R_PR), 70);
    }

    #[test]
    fn test_unknown_element() {
        let unknown = XName::new("http://unknown.ns", "unknownElement");
        assert_eq!(get_element_order(&unknown, &ORDER_P_PR), 999);
    }
}
