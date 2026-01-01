// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

//! PmlMarkupRenderer - Produces marked presentations with visual change overlays
//!
//! This module is responsible for:
//! 1. Taking a newer presentation + comparison result
//! 2. Cloning the newer presentation
//! 3. Adding visual overlays for each change (labels, boxes, callouts)
//! 4. Optionally adding speaker notes annotations
//! 5. Optionally adding a summary slide at the end
//!
//! Visual Overlay Types:
//! - ShapeInserted: "NEW" label (Green)
//! - ShapeDeleted: "DELETED" label (Red)
//! - ShapeMoved: "MOVED" label (Blue)
//! - ShapeResized: "RESIZED" label (Orange)
//! - TextChanged: "TEXT CHANGED" label (Orange)
//! - ImageReplaced: "IMAGE REPLACED" label (Orange)
//!
//! C# Parity: Faithful port of PmlComparer.cs lines 2176-2677

use super::result::PmlComparisonResult;
use super::types::{PmlChange, PmlChangeType};
use super::{PmlComparerSettings, PmlDocument};
use crate::error::Result;
use crate::xml::namespaces::XMLNS;
use crate::xml::{XAttribute, XName, XmlDocument, XmlNodeData, A, P, R};
use std::collections::HashMap;

/// Main entry point for rendering a marked presentation.
///
/// Takes the newer presentation as a base and adds visual annotations for each detected change.
/// Returns the original document unchanged if no changes detected.
pub fn render_marked_presentation(
    newer_doc: &PmlDocument,
    result: &PmlComparisonResult,
    settings: &PmlComparerSettings,
) -> Result<PmlDocument> {
    // Early return if no changes
    if result.total_changes == 0 {
        let bytes = newer_doc.to_bytes()?;
        return PmlDocument::from_bytes(&bytes);
    }

    // Clone the document
    let bytes = newer_doc.to_bytes()?;
    let mut doc = PmlDocument::from_bytes(&bytes)?;
    let pkg = doc.package_mut();

    // Get slide paths from presentation.xml
    let slide_paths = get_slide_paths(pkg)?;

    // Group changes by slide index
    let changes_by_slide = group_changes_by_slide(&result.changes);

    // Process each slide with changes
    for (slide_index, changes) in changes_by_slide {
        if let Some(slide_path) =
            slide_paths.get(&slide_index.expect("Slide index must be present"))
        {
            if let Ok(mut slide_doc) = pkg.get_xml_part(slide_path) {
                add_change_overlays(&mut slide_doc, &changes, settings);
                pkg.put_xml_part(slide_path, &slide_doc)?;
            }
        }
    }

    // Add summary slide if enabled
    if settings.add_summary_slide {
        add_summary_slide(pkg, result, settings)?;
    }

    Ok(doc)
}

/// Get mapping of slide indices to their XML paths
fn get_slide_paths(pkg: &crate::package::OoxmlPackage) -> Result<HashMap<usize, String>> {
    let mut paths = HashMap::new();

    // Read presentation.xml
    let pres_path = "ppt/presentation.xml";
    let pres_doc = pkg.get_xml_part(pres_path)?;

    // Read relationships
    let rels_path = "ppt/_rels/presentation.xml.rels";
    let rels_doc = pkg.get_xml_part(rels_path)?;

    // Build rId to target path mapping
    let mut rid_to_path: HashMap<String, String> = HashMap::new();
    if let Some(rels_root) = rels_doc.root() {
        for rel_id in rels_doc.children(rels_root) {
            if let Some(data) = rels_doc.get(rel_id) {
                if let Some(attrs) = data.attributes() {
                    let mut id = None;
                    let mut target = None;
                    let mut type_attr = None;
                    for attr in attrs {
                        if attr.name.local_name == "Id" {
                            id = Some(attr.value.clone());
                        } else if attr.name.local_name == "Target" {
                            target = Some(attr.value.clone());
                        } else if attr.name.local_name == "Type" {
                            type_attr = Some(attr.value.clone());
                        }
                    }
                    // Only include slide relationships
                    if let (Some(id), Some(target), Some(type_)) = (id, target, type_attr) {
                        if type_.contains("slide")
                            && !type_.contains("slideLayout")
                            && !type_.contains("slideMaster")
                        {
                            let full_path = if target.starts_with('/') {
                                target[1..].to_string()
                            } else {
                                format!("ppt/{}", target)
                            };
                            rid_to_path.insert(id, full_path);
                        }
                    }
                }
            }
        }
    }

    // Find sldIdLst in presentation
    if let Some(pres_root) = pres_doc.root() {
        for node_id in pres_doc.descendants(pres_root) {
            if let Some(data) = pres_doc.get(node_id) {
                if let Some(name) = data.name() {
                    if name.local_name == "sldId" && name.namespace.as_deref() == Some(P::NS) {
                        if let Some(attrs) = data.attributes() {
                            let mut r_id = None;
                            for attr in attrs {
                                if attr.name.local_name == "id"
                                    && attr.name.namespace.as_deref() == Some(R::NS)
                                {
                                    r_id = Some(attr.value.clone());
                                }
                            }
                            if let Some(rid) = r_id {
                                if let Some(path) = rid_to_path.get(&rid) {
                                    // Extract slide number from path (e.g., slide1.xml -> 0)
                                    let slide_num = extract_slide_number(path);
                                    paths.insert(slide_num, path.clone());
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    Ok(paths)
}

/// Extract slide number from path (e.g., "ppt/slides/slide1.xml" -> 0)
fn extract_slide_number(path: &str) -> usize {
    let filename = path.rsplit('/').next().unwrap_or(path);
    let num_str: String = filename.chars().filter(|c| c.is_ascii_digit()).collect();
    num_str.parse::<usize>().unwrap_or(1).saturating_sub(1)
}

fn group_changes_by_slide(changes: &[PmlChange]) -> HashMap<Option<usize>, Vec<&PmlChange>> {
    let mut by_slide: HashMap<Option<usize>, Vec<&PmlChange>> = HashMap::new();
    for change in changes {
        by_slide.entry(change.slide_index).or_default().push(change);
    }
    by_slide
}

/// Add visual change overlays to a slide
fn add_change_overlays(
    slide_doc: &mut XmlDocument,
    changes: &[&PmlChange],
    settings: &PmlComparerSettings,
) {
    let Some(root) = slide_doc.root() else {
        return;
    };

    // Find spTree (shape tree)
    let sp_tree_name = P::sp_tree();
    let sp_tree_id = find_sp_tree(slide_doc, root);

    let Some(sp_tree_id) = sp_tree_id else {
        return;
    };

    // Get next available shape ID
    let mut next_id = get_next_shape_id(slide_doc, sp_tree_id);

    // Add overlays for each change
    for change in changes {
        let (label_text, color) = get_label_for_change(change, settings);
        if !label_text.is_empty() {
            add_change_label(
                slide_doc,
                sp_tree_id,
                change,
                &label_text,
                &color,
                &mut next_id,
            );
        }
    }
}

/// Find the spTree element in a slide
fn find_sp_tree(doc: &XmlDocument, root: indextree::NodeId) -> Option<indextree::NodeId> {
    // Look for cSld/spTree
    let c_sld_name = P::c_sld();
    let sp_tree_name = P::sp_tree();

    for node_id in doc.descendants(root) {
        if let Some(data) = doc.get(node_id) {
            if let Some(name) = data.name() {
                if name.local_name == "spTree" && name.namespace.as_deref() == Some(P::NS) {
                    return Some(node_id);
                }
            }
        }
    }
    None
}

/// Get the next available shape ID from a shape tree
fn get_next_shape_id(doc: &XmlDocument, sp_tree_id: indextree::NodeId) -> u32 {
    let mut max_id: u32 = 0;

    for node_id in doc.descendants(sp_tree_id) {
        if let Some(data) = doc.get(node_id) {
            if let Some(name) = data.name() {
                if name.local_name == "cNvPr" {
                    if let Some(attrs) = data.attributes() {
                        for attr in attrs {
                            if attr.name.local_name == "id" {
                                if let Ok(id) = attr.value.parse::<u32>() {
                                    max_id = max_id.max(id);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    max_id + 1
}

/// Get label text and color for a change type
fn get_label_for_change(change: &PmlChange, settings: &PmlComparerSettings) -> (String, String) {
    match change.change_type {
        PmlChangeType::ShapeInserted => ("NEW".to_string(), settings.inserted_color.clone()),
        PmlChangeType::ShapeDeleted => ("DELETED".to_string(), settings.deleted_color.clone()),
        PmlChangeType::ShapeMoved => ("MOVED".to_string(), settings.moved_color.clone()),
        PmlChangeType::ShapeResized => ("RESIZED".to_string(), settings.modified_color.clone()),
        PmlChangeType::TextChanged => ("TEXT CHANGED".to_string(), settings.modified_color.clone()),
        PmlChangeType::TextFormattingChanged => {
            ("FORMATTING".to_string(), settings.modified_color.clone())
        }
        PmlChangeType::ImageReplaced => (
            "IMAGE REPLACED".to_string(),
            settings.modified_color.clone(),
        ),
        PmlChangeType::TableContentChanged => {
            ("TABLE CHANGED".to_string(), settings.modified_color.clone())
        }
        PmlChangeType::ChartDataChanged => {
            ("CHART CHANGED".to_string(), settings.modified_color.clone())
        }
        _ => (String::new(), String::new()),
    }
}

/// Add a change label shape to the slide
fn add_change_label(
    doc: &mut XmlDocument,
    sp_tree_id: indextree::NodeId,
    change: &PmlChange,
    label_text: &str,
    color: &str,
    next_id: &mut u32,
) {
    // Position: use change coordinates or default
    let x = change.new_x.unwrap_or(914400.0) as i64; // Default 1 inch from left
    let y = change
        .new_y
        .map(|y: f64| (y - 300000.0).max(0.0) as i64)
        .unwrap_or(457200); // Above shape
    let cx = 1828800_i64; // 2 inches width
    let cy = 274320_i64; // 0.3 inches height

    // Create sp (shape) element
    let sp = doc.add_child(sp_tree_id, XmlNodeData::element(P::sp()));

    // nvSpPr (non-visual shape properties)
    let nv_sp_pr = doc.add_child(sp, XmlNodeData::element(P::nv_sp_pr()));

    // cNvPr (common non-visual properties)
    let shape_name = format!("ChangeLabel_{}", *next_id);
    doc.add_child(
        nv_sp_pr,
        XmlNodeData::element_with_attrs(
            P::c_nv_pr(),
            vec![
                XAttribute::new(XName::local("id"), &next_id.to_string()),
                XAttribute::new(XName::local("name"), &shape_name),
            ],
        ),
    );

    // cNvSpPr
    let c_nv_sp_pr = doc.add_child(nv_sp_pr, XmlNodeData::element(P::c_nv_sp_pr()));
    doc.add_child(
        c_nv_sp_pr,
        XmlNodeData::element_with_attrs(
            P::sp_locks(),
            vec![XAttribute::new(XName::local("noGrp"), "1")],
        ),
    );

    // nvPr
    doc.add_child(nv_sp_pr, XmlNodeData::element(P::nv_pr()));

    // spPr (shape properties)
    let sp_pr = doc.add_child(sp, XmlNodeData::element(P::sp_pr()));

    // xfrm (transform)
    let xfrm = doc.add_child(sp_pr, XmlNodeData::element(A::xfrm()));
    doc.add_child(
        xfrm,
        XmlNodeData::element_with_attrs(
            A::off(),
            vec![
                XAttribute::new(XName::local("x"), &x.to_string()),
                XAttribute::new(XName::local("y"), &y.to_string()),
            ],
        ),
    );
    doc.add_child(
        xfrm,
        XmlNodeData::element_with_attrs(
            A::ext(),
            vec![
                XAttribute::new(XName::local("cx"), &cx.to_string()),
                XAttribute::new(XName::local("cy"), &cy.to_string()),
            ],
        ),
    );

    // prstGeom (preset geometry - rectangle)
    let prst_geom = doc.add_child(
        sp_pr,
        XmlNodeData::element_with_attrs(
            A::prst_geom(),
            vec![XAttribute::new(XName::local("prst"), "rect")],
        ),
    );
    doc.add_child(prst_geom, XmlNodeData::element(A::avLst()));

    // solidFill (background color)
    let solid_fill = doc.add_child(sp_pr, XmlNodeData::element(A::solid_fill()));
    doc.add_child(
        solid_fill,
        XmlNodeData::element_with_attrs(
            A::srgb_clr(),
            vec![XAttribute::new(XName::local("val"), color)],
        ),
    );

    // ln (outline)
    let ln = doc.add_child(
        sp_pr,
        XmlNodeData::element_with_attrs(
            A::ln(),
            vec![XAttribute::new(XName::local("w"), "9525")], // 0.75pt
        ),
    );
    doc.add_child(ln, XmlNodeData::element(A::no_fill()));

    // txBody (text body)
    let tx_body = doc.add_child(sp, XmlNodeData::element(P::tx_body()));

    // bodyPr
    doc.add_child(
        tx_body,
        XmlNodeData::element_with_attrs(
            A::body_pr(),
            vec![
                XAttribute::new(XName::local("wrap"), "square"),
                XAttribute::new(XName::local("rtlCol"), "0"),
            ],
        ),
    );

    // lstStyle
    doc.add_child(tx_body, XmlNodeData::element(A::lst_style()));

    // a:p (paragraph)
    let para = doc.add_child(tx_body, XmlNodeData::element(A::p()));

    // a:r (run)
    let run = doc.add_child(para, XmlNodeData::element(A::r()));

    // rPr (run properties) - white text, bold
    doc.add_child(
        run,
        XmlNodeData::element_with_attrs(
            A::r_pr(),
            vec![
                XAttribute::new(XName::local("lang"), "en-US"),
                XAttribute::new(XName::local("sz"), "1100"), // 11pt
                XAttribute::new(XName::local("b"), "1"),     // bold
            ],
        ),
    );

    // a:t (text)
    let t = doc.add_child(run, XmlNodeData::element(A::t()));
    doc.add_child(t, XmlNodeData::text(label_text));

    *next_id += 1;
}

/// Add a summary slide at the end of the presentation
fn add_summary_slide(
    pkg: &mut crate::package::OoxmlPackage,
    result: &PmlComparisonResult,
    _settings: &PmlComparerSettings,
) -> Result<()> {
    // Create slide content
    let mut slide_doc = XmlDocument::new();

    let root = slide_doc.add_root(XmlNodeData::element_with_attrs(
        P::sld(),
        vec![
            XAttribute::new(XName::new(XMLNS::NS, "a"), A::NS),
            XAttribute::new(XName::new(XMLNS::NS, "r"), R::NS),
            XAttribute::new(XName::new(XMLNS::NS, "p"), P::NS),
        ],
    ));

    // cSld
    let c_sld = slide_doc.add_child(root, XmlNodeData::element(P::c_sld()));

    // spTree
    let sp_tree = slide_doc.add_child(c_sld, XmlNodeData::element(P::sp_tree()));

    // Group shape properties (required)
    let nv_grp_sp_pr = slide_doc.add_child(sp_tree, XmlNodeData::element(P::nv_grp_sp_pr()));
    slide_doc.add_child(
        nv_grp_sp_pr,
        XmlNodeData::element_with_attrs(
            P::c_nv_pr(),
            vec![
                XAttribute::new(XName::local("id"), "1"),
                XAttribute::new(XName::local("name"), ""),
            ],
        ),
    );
    slide_doc.add_child(
        nv_grp_sp_pr,
        XmlNodeData::element(XName::new(P::NS, "cNvGrpSpPr")),
    );
    slide_doc.add_child(nv_grp_sp_pr, XmlNodeData::element(P::nv_pr()));

    let grp_sp_pr = slide_doc.add_child(sp_tree, XmlNodeData::element(P::grp_sp_pr()));
    let xfrm = slide_doc.add_child(grp_sp_pr, XmlNodeData::element(A::xfrm()));
    slide_doc.add_child(
        xfrm,
        XmlNodeData::element_with_attrs(
            A::off(),
            vec![
                XAttribute::new(XName::local("x"), "0"),
                XAttribute::new(XName::local("y"), "0"),
            ],
        ),
    );
    slide_doc.add_child(
        xfrm,
        XmlNodeData::element_with_attrs(
            A::ext(),
            vec![
                XAttribute::new(XName::local("cx"), "0"),
                XAttribute::new(XName::local("cy"), "0"),
            ],
        ),
    );
    slide_doc.add_child(
        xfrm,
        XmlNodeData::element_with_attrs(
            XName::new(A::NS, "chOff"),
            vec![
                XAttribute::new(XName::local("x"), "0"),
                XAttribute::new(XName::local("y"), "0"),
            ],
        ),
    );
    slide_doc.add_child(
        xfrm,
        XmlNodeData::element_with_attrs(
            XName::new(A::NS, "chExt"),
            vec![
                XAttribute::new(XName::local("cx"), "0"),
                XAttribute::new(XName::local("cy"), "0"),
            ],
        ),
    );

    // Add title shape
    add_text_shape(
        &mut slide_doc,
        sp_tree,
        2,
        "Comparison Summary",
        457200,
        274638,
        8229600,
        1143000,
        4400,
        true,
    );

    // Build summary text
    let mut summary_lines = vec![
        format!("Total Changes: {}", result.total_changes),
        format!("Slides Inserted: {}", result.slides_inserted),
        format!("Slides Deleted: {}", result.slides_deleted),
        format!("Shapes Inserted: {}", result.shapes_inserted),
        format!("Shapes Deleted: {}", result.shapes_deleted),
        format!("Shapes Moved: {}", result.shapes_moved),
        format!("Shapes Resized: {}", result.shapes_resized),
        format!("Text Changes: {}", result.text_changes),
    ];

    // Add individual change descriptions (limit to first 20)
    summary_lines.push(String::new());
    summary_lines.push("Changes:".to_string());
    for (i, change) in result.changes.iter().take(20).enumerate() {
        let desc = change.get_description();
        summary_lines.push(format!("{}. {}", i + 1, desc));
    }
    if result.changes.len() > 20 {
        summary_lines.push(format!("... and {} more", result.changes.len() - 20));
    }

    let summary_text = summary_lines.join("\n");
    add_text_shape(
        &mut slide_doc,
        sp_tree,
        3,
        &summary_text,
        457200,
        1600200,
        8229600,
        4525963,
        1800,
        false,
    );

    // Determine slide path
    let slide_num = get_next_slide_number(pkg)?;
    let slide_path = format!("ppt/slides/slide{}.xml", slide_num);

    // Save slide
    pkg.put_xml_part(&slide_path, &slide_doc)?;

    // Update presentation.xml to include new slide
    update_presentation_for_new_slide(pkg, slide_num)?;

    Ok(())
}

/// Add a text shape to the slide
fn add_text_shape(
    doc: &mut XmlDocument,
    sp_tree: indextree::NodeId,
    id: u32,
    text: &str,
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    font_size: i32,
    is_title: bool,
) {
    let sp = doc.add_child(sp_tree, XmlNodeData::element(P::sp()));

    // nvSpPr
    let nv_sp_pr = doc.add_child(sp, XmlNodeData::element(P::nv_sp_pr()));
    let name = if is_title { "Title" } else { "Content" };
    doc.add_child(
        nv_sp_pr,
        XmlNodeData::element_with_attrs(
            P::c_nv_pr(),
            vec![
                XAttribute::new(XName::local("id"), &id.to_string()),
                XAttribute::new(XName::local("name"), name),
            ],
        ),
    );
    doc.add_child(nv_sp_pr, XmlNodeData::element(P::c_nv_sp_pr()));
    doc.add_child(nv_sp_pr, XmlNodeData::element(P::nv_pr()));

    // spPr
    let sp_pr = doc.add_child(sp, XmlNodeData::element(P::sp_pr()));
    let xfrm = doc.add_child(sp_pr, XmlNodeData::element(A::xfrm()));
    doc.add_child(
        xfrm,
        XmlNodeData::element_with_attrs(
            A::off(),
            vec![
                XAttribute::new(XName::local("x"), &x.to_string()),
                XAttribute::new(XName::local("y"), &y.to_string()),
            ],
        ),
    );
    doc.add_child(
        xfrm,
        XmlNodeData::element_with_attrs(
            A::ext(),
            vec![
                XAttribute::new(XName::local("cx"), &cx.to_string()),
                XAttribute::new(XName::local("cy"), &cy.to_string()),
            ],
        ),
    );
    let prst_geom = doc.add_child(
        sp_pr,
        XmlNodeData::element_with_attrs(
            A::prst_geom(),
            vec![XAttribute::new(XName::local("prst"), "rect")],
        ),
    );
    doc.add_child(prst_geom, XmlNodeData::element(A::avLst()));

    // txBody
    let tx_body = doc.add_child(sp, XmlNodeData::element(P::tx_body()));
    doc.add_child(tx_body, XmlNodeData::element(A::body_pr()));
    doc.add_child(tx_body, XmlNodeData::element(A::lst_style()));

    // Split text into paragraphs
    for line in text.lines() {
        let para = doc.add_child(tx_body, XmlNodeData::element(A::p()));
        let run = doc.add_child(para, XmlNodeData::element(A::r()));

        let mut attrs = vec![
            XAttribute::new(XName::local("lang"), "en-US"),
            XAttribute::new(XName::local("sz"), &(font_size * 100).to_string()),
        ];
        if is_title {
            attrs.push(XAttribute::new(XName::local("b"), "1"));
        }
        doc.add_child(run, XmlNodeData::element_with_attrs(A::r_pr(), attrs));

        let t = doc.add_child(run, XmlNodeData::element(A::t()));
        doc.add_child(t, XmlNodeData::text(line));
    }
}

/// Get the next slide number
fn get_next_slide_number(pkg: &crate::package::OoxmlPackage) -> Result<usize> {
    let mut max_num = 0;

    // Check existing slides
    for name in pkg.part_names() {
        if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
            let num_str: String = name
                .trim_start_matches("ppt/slides/slide")
                .trim_end_matches(".xml")
                .to_string();
            if let Ok(num) = num_str.parse::<usize>() {
                max_num = max_num.max(num);
            }
        }
    }

    Ok(max_num + 1)
}

/// Update presentation.xml to include a new slide
fn update_presentation_for_new_slide(
    pkg: &mut crate::package::OoxmlPackage,
    slide_num: usize,
) -> Result<()> {
    let pres_path = "ppt/presentation.xml";
    let mut pres_doc = pkg.get_xml_part(pres_path)?;

    let Some(pres_root) = pres_doc.root() else {
        return Ok(());
    };

    // Find sldIdLst
    let sld_id_lst_name = P::sld_id_lst();
    let sld_id_lst = pres_doc
        .find_child(pres_root, &sld_id_lst_name)
        .unwrap_or_else(|| {
            pres_doc.add_child(pres_root, XmlNodeData::element(sld_id_lst_name.clone()))
        });

    // Find max sldId
    let sld_id_name = P::sld_id();
    let mut max_id: u32 = 255;
    for child_id in pres_doc.children(sld_id_lst).collect::<Vec<_>>() {
        if let Some(data) = pres_doc.get(child_id) {
            if let Some(attrs) = data.attributes() {
                for attr in attrs {
                    if attr.name.local_name == "id" && attr.name.namespace.is_none() {
                        if let Ok(id) = attr.value.parse::<u32>() {
                            max_id = max_id.max(id);
                        }
                    }
                }
            }
        }
    }

    // Add new sldId
    let new_id = max_id + 1;
    let rel_id = format!("rIdSummary{}", slide_num);
    pres_doc.add_child(
        sld_id_lst,
        XmlNodeData::element_with_attrs(
            sld_id_name,
            vec![
                XAttribute::new(XName::local("id"), &new_id.to_string()),
                XAttribute::new(XName::new(R::NS, "id"), &rel_id),
            ],
        ),
    );

    pkg.put_xml_part(pres_path, &pres_doc)?;

    // Add relationship
    let rels_path = "ppt/_rels/presentation.xml.rels";
    if let Ok(mut rels_doc) = pkg.get_xml_part(rels_path) {
        if let Some(rels_root) = rels_doc.root() {
            rels_doc.add_child(rels_root, XmlNodeData::element_with_attrs(
                XName::new("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship"),
                vec![
                    XAttribute::new(XName::local("Id"), &rel_id),
                    XAttribute::new(XName::local("Type"), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"),
                    XAttribute::new(XName::local("Target"), &format!("slides/slide{}.xml", slide_num)),
                ],
            ));
            pkg.put_xml_part(rels_path, &rels_doc)?;
        }
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn group_changes_by_slide_works() {
        let changes = vec![
            PmlChange {
                slide_index: 1,
                shape_id: Some("shape1".to_string()),
                shape_name: Some("Title".to_string()),
                change_type: PmlChangeType::TextChanged,
                description: Some("Text changed".to_string()),
                new_x: Some(100),
                new_y: Some(200),
                old_x: None,
                old_y: None,
            },
            PmlChange {
                slide_index: 1,
                shape_id: Some("shape2".to_string()),
                shape_name: Some("Content".to_string()),
                change_type: PmlChangeType::ShapeInserted,
                description: Some("Shape inserted".to_string()),
                new_x: Some(300),
                new_y: Some(400),
                old_x: None,
                old_y: None,
            },
            PmlChange {
                slide_index: 2,
                shape_id: Some("shape3".to_string()),
                shape_name: Some("Image".to_string()),
                change_type: PmlChangeType::ImageReplaced,
                description: Some("Image replaced".to_string()),
                new_x: Some(500),
                new_y: Some(600),
                old_x: None,
                old_y: None,
            },
        ];

        let grouped = group_changes_by_slide(&changes);

        assert_eq!(grouped.len(), 2);
        assert_eq!(grouped.get(&1).unwrap().len(), 2);
        assert_eq!(grouped.get(&2).unwrap().len(), 1);
    }

    #[test]
    fn extract_slide_number_works() {
        assert_eq!(extract_slide_number("ppt/slides/slide1.xml"), 0);
        assert_eq!(extract_slide_number("ppt/slides/slide5.xml"), 4);
        assert_eq!(extract_slide_number("ppt/slides/slide10.xml"), 9);
    }

    #[test]
    fn get_label_for_change_returns_correct_labels() {
        let settings = PmlComparerSettings::default();

        let change = PmlChange {
            slide_index: 0,
            shape_id: None,
            shape_name: None,
            change_type: PmlChangeType::ShapeInserted,
            description: None,
            new_x: None,
            new_y: None,
            old_x: None,
            old_y: None,
        };
        let (label, _color) = get_label_for_change(&change, &settings);
        assert_eq!(label, "NEW");

        let change2 = PmlChange {
            change_type: PmlChangeType::ShapeDeleted,
            ..change.clone()
        };
        let (label2, _) = get_label_for_change(&change2, &settings);
        assert_eq!(label2, "DELETED");
    }
}
