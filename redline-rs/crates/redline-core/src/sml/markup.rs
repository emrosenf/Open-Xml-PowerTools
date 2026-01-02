// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

//! SmlMarkupRenderer - Renders a marked workbook showing differences.
//!
//! This module is responsible for taking a SmlComparisonResult and producing
//! a marked Excel workbook with:
//! - Highlight fills for changed cells
//! - Comments explaining changes
//! - A _DiffSummary sheet with change statistics
//!
//! ## C# Parity
//!
//! This is a faithful port of C# SmlMarkupRenderer from SmlComparer.cs.

use crate::error::Result;
use crate::sml::types::{SmlChange, SmlChangeType};
use crate::sml::{SmlComparerSettings, SmlComparisonResult, SmlDocument};
use crate::xml::namespaces::XMLNS;
use crate::xml::{XAttribute, XName, XmlDocument, XmlNodeData, R, S};
use std::collections::HashMap;

/// Internal structure holding style IDs for highlight fills.
#[derive(Debug, Clone)]
struct HighlightStyles {
    added_fill_id: usize,
    modified_value_fill_id: usize,
    modified_formula_fill_id: usize,
    modified_format_fill_id: usize,

    added_style_id: usize,
    modified_value_style_id: usize,
    modified_formula_style_id: usize,
    modified_format_style_id: usize,
}

/// Renders a marked workbook highlighting all differences between two workbooks.
///
/// The output is based on the source workbook (typically the newer version) with:
/// - Cell highlights using fill colors from settings
/// - Cell comments describing the changes
/// - A `_DiffSummary` sheet with statistics and detailed change list
pub(crate) fn render_marked_workbook(
    source: &SmlDocument,
    result: &SmlComparisonResult,
    settings: &SmlComparerSettings,
) -> Result<SmlDocument> {
    // Clone source document to work with
    let bytes = source.to_bytes()?;
    let mut doc = SmlDocument::from_bytes(&bytes)?;
    let pkg = doc.package_mut();

    // Add highlight styles to styles.xml
    let styles_path = "xl/styles.xml";
    let highlight_styles = if let Ok(mut styles_doc) = pkg.get_xml_part(styles_path) {
        let hs = add_highlight_styles(&mut styles_doc, settings);
        pkg.put_xml_part(styles_path, &styles_doc)?;
        hs
    } else {
        // No styles part - create minimal one
        let mut styles_doc = create_minimal_styles();
        let hs = add_highlight_styles(&mut styles_doc, settings);
        pkg.put_xml_part(styles_path, &styles_doc)?;
        hs
    };

    // Group changes by sheet
    let mut changes_by_sheet: HashMap<String, Vec<&SmlChange>> = HashMap::new();
    for change in &result.changes {
        if change.cell_address.is_some() {
            if let Some(sheet_name) = &change.sheet_name {
                changes_by_sheet
                    .entry(sheet_name.clone())
                    .or_default()
                    .push(change);
            }
        }
    }

    // Get sheet name to path mapping
    let sheet_paths = get_sheet_paths(pkg)?;

    // Process each sheet with changes
    for (sheet_name, changes) in &changes_by_sheet {
        if let Some(sheet_path) = sheet_paths.get(sheet_name) {
            if let Ok(mut ws_doc) = pkg.get_xml_part(sheet_path) {
                // Apply cell highlights
                for change in changes {
                    apply_cell_highlight(&mut ws_doc, change, &highlight_styles);
                }
                pkg.put_xml_part(sheet_path, &ws_doc)?;

                // Add comments for changes
                add_comments_for_changes(pkg, sheet_path, changes, settings)?;
            }
        }
    }

    // Add summary sheet
    add_diff_summary_sheet(pkg, result, settings)?;

    Ok(doc)
}

/// Get mapping of sheet names to their XML paths
fn get_sheet_paths(pkg: &crate::package::OoxmlPackage) -> Result<HashMap<String, String>> {
    let mut paths = HashMap::new();

    let workbook_path = "xl/workbook.xml";
    let workbook_doc = pkg.get_xml_part(workbook_path)?;

    // Get workbook relationships
    let rels_path = "xl/_rels/workbook.xml.rels";
    let rels_doc = pkg.get_xml_part(rels_path)?;

    // Build rId to target path mapping
    let mut rid_to_path: HashMap<String, String> = HashMap::new();
    if let Some(rels_root) = rels_doc.root() {
        for rel_id in rels_doc.children(rels_root) {
            if let Some(data) = rels_doc.get(rel_id) {
                if let Some(attrs) = data.attributes() {
                    let mut id = None;
                    let mut target = None;
                    for attr in attrs {
                        if attr.name.local_name == "Id" {
                            id = Some(attr.value.clone());
                        } else if attr.name.local_name == "Target" {
                            target = Some(attr.value.clone());
                        }
                    }
                    if let (Some(id), Some(target)) = (id, target) {
                        let full_path = if target.starts_with('/') {
                            target[1..].to_string()
                        } else {
                            format!("xl/{}", target)
                        };
                        rid_to_path.insert(id, full_path);
                    }
                }
            }
        }
    }

    // Find sheets in workbook
    if let Some(wb_root) = workbook_doc.root() {
        for node_id in workbook_doc.descendants(wb_root) {
            if let Some(data) = workbook_doc.get(node_id) {
                if let Some(name) = data.name() {
                    if name.local_name == "sheet" && name.namespace.as_deref() == Some(S::NS) {
                        if let Some(attrs) = data.attributes() {
                            let mut sheet_name = None;
                            let mut r_id = None;
                            for attr in attrs {
                                if attr.name.local_name == "name" {
                                    sheet_name = Some(attr.value.clone());
                                } else if attr.name.local_name == "id"
                                    && attr.name.namespace.as_deref() == Some(R::NS)
                                {
                                    r_id = Some(attr.value.clone());
                                }
                            }
                            if let (Some(name), Some(rid)) = (sheet_name, r_id) {
                                if let Some(path) = rid_to_path.get(&rid) {
                                    paths.insert(name, path.clone());
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

/// Add highlight fill styles to the workbook styles.
fn add_highlight_styles(
    styles_doc: &mut XmlDocument,
    settings: &SmlComparerSettings,
) -> HighlightStyles {
    let mut styles = HighlightStyles {
        added_fill_id: 0,
        modified_value_fill_id: 0,
        modified_formula_fill_id: 0,
        modified_format_fill_id: 0,
        added_style_id: 0,
        modified_value_style_id: 0,
        modified_formula_style_id: 0,
        modified_format_style_id: 0,
    };

    let Some(root) = styles_doc.root() else {
        return styles;
    };

    // Find or create fills element
    let fills_name = S::fills();
    let fills_id = styles_doc.find_child(root, &fills_name).unwrap_or_else(|| {
        // Create fills element
        let fills = styles_doc.add_child(
            root,
            XmlNodeData::element_with_attrs(
                fills_name.clone(),
                vec![XAttribute::new(XName::local("count"), "0")],
            ),
        );
        fills
    });

    // Count existing fills
    let fill_name = S::fill();
    let mut fill_count: usize = styles_doc
        .children(fills_id)
        .filter(|&c| {
            styles_doc
                .get(c)
                .and_then(|d| d.name())
                .map(|n| n == &fill_name)
                .unwrap_or(false)
        })
        .count();

    // Add highlight fills
    styles.added_fill_id = fill_count;
    add_solid_fill(styles_doc, fills_id, &settings.added_cell_color);
    fill_count += 1;

    styles.modified_value_fill_id = fill_count;
    add_solid_fill(styles_doc, fills_id, &settings.modified_value_color);
    fill_count += 1;

    styles.modified_formula_fill_id = fill_count;
    add_solid_fill(styles_doc, fills_id, &settings.modified_formula_color);
    fill_count += 1;

    styles.modified_format_fill_id = fill_count;
    add_solid_fill(styles_doc, fills_id, &settings.modified_format_color);
    fill_count += 1;

    // Update fills count
    styles_doc.set_attribute(fills_id, &XName::local("count"), &fill_count.to_string());

    // Find or create cellXfs element
    let cell_xfs_name = S::cellXfs();
    let cell_xfs_id = styles_doc
        .find_child(root, &cell_xfs_name)
        .unwrap_or_else(|| {
            styles_doc.add_child(
                root,
                XmlNodeData::element_with_attrs(
                    cell_xfs_name.clone(),
                    vec![XAttribute::new(XName::local("count"), "0")],
                ),
            )
        });

    // Count existing xf elements
    let xf_name = S::xf();
    let mut xf_count: usize = styles_doc
        .children(cell_xfs_id)
        .filter(|&c| {
            styles_doc
                .get(c)
                .and_then(|d| d.name())
                .map(|n| n == &xf_name)
                .unwrap_or(false)
        })
        .count();

    // Add cell formats that use the highlight fills
    styles.added_style_id = xf_count;
    add_xf_with_fill(styles_doc, cell_xfs_id, styles.added_fill_id);
    xf_count += 1;

    styles.modified_value_style_id = xf_count;
    add_xf_with_fill(styles_doc, cell_xfs_id, styles.modified_value_fill_id);
    xf_count += 1;

    styles.modified_formula_style_id = xf_count;
    add_xf_with_fill(styles_doc, cell_xfs_id, styles.modified_formula_fill_id);
    xf_count += 1;

    styles.modified_format_style_id = xf_count;
    add_xf_with_fill(styles_doc, cell_xfs_id, styles.modified_format_fill_id);
    xf_count += 1;

    // Update cellXfs count
    styles_doc.set_attribute(cell_xfs_id, &XName::local("count"), &xf_count.to_string());

    styles
}

/// Add a solid fill element
fn add_solid_fill(doc: &mut XmlDocument, fills_id: indextree::NodeId, color: &str) {
    let fill_id = doc.add_child(fills_id, XmlNodeData::element(S::fill()));

    let pattern_fill = doc.add_child(
        fill_id,
        XmlNodeData::element_with_attrs(
            S::patternFill(),
            vec![XAttribute::new(XName::local("patternType"), "solid")],
        ),
    );

    // fgColor with RGB (prepend FF for alpha)
    let rgb = format!("FF{}", color);
    doc.add_child(
        pattern_fill,
        XmlNodeData::element_with_attrs(
            S::fgColor(),
            vec![XAttribute::new(XName::local("rgb"), &rgb)],
        ),
    );

    // bgColor indexed=64
    doc.add_child(
        pattern_fill,
        XmlNodeData::element_with_attrs(
            S::bgColor(),
            vec![XAttribute::new(XName::local("indexed"), "64")],
        ),
    );
}

/// Add an xf element with the given fill ID
fn add_xf_with_fill(doc: &mut XmlDocument, cell_xfs_id: indextree::NodeId, fill_id: usize) {
    doc.add_child(
        cell_xfs_id,
        XmlNodeData::element_with_attrs(
            S::xf(),
            vec![
                XAttribute::new(XName::local("numFmtId"), "0"),
                XAttribute::new(XName::local("fontId"), "0"),
                XAttribute::new(XName::local("fillId"), &fill_id.to_string()),
                XAttribute::new(XName::local("borderId"), "0"),
                XAttribute::new(XName::local("applyFill"), "1"),
            ],
        ),
    );
}

/// Create a minimal styles.xml document
fn create_minimal_styles() -> XmlDocument {
    let mut doc = XmlDocument::new();

    let root = doc.add_root(XmlNodeData::element_with_attrs(
        S::styleSheet(),
        vec![XAttribute::new(XName::new(XMLNS::NS, "x"), S::NS)],
    ));

    // Add empty fills
    doc.add_child(
        root,
        XmlNodeData::element_with_attrs(
            S::fills(),
            vec![XAttribute::new(XName::local("count"), "0")],
        ),
    );

    // Add empty cellXfs
    doc.add_child(
        root,
        XmlNodeData::element_with_attrs(
            S::cellXfs(),
            vec![XAttribute::new(XName::local("count"), "0")],
        ),
    );

    doc
}

/// Apply highlight style to a specific cell in a worksheet.
fn apply_cell_highlight(ws_doc: &mut XmlDocument, change: &SmlChange, styles: &HighlightStyles) {
    let Some(cell_address) = &change.cell_address else {
        return;
    };

    let Some(root) = ws_doc.root() else {
        return;
    };

    // Find sheetData
    let sheet_data_name = S::sheetData();
    let Some(sheet_data_id) = ws_doc.find_child(root, &sheet_data_name) else {
        return;
    };

    let (col_index, row_index) = parse_cell_ref(cell_address);

    // Find or create row
    let row_name = S::row();
    let row_id = find_or_create_row(ws_doc, sheet_data_id, row_index, &row_name);

    // Find or create cell
    let cell_name = S::c();
    let cell_id = find_or_create_cell(ws_doc, row_id, cell_address, col_index, &cell_name);

    // Determine style ID based on change type
    let style_id = match change.change_type {
        SmlChangeType::CellAdded => Some(styles.added_style_id),
        SmlChangeType::ValueChanged => Some(styles.modified_value_style_id),
        SmlChangeType::FormulaChanged => Some(styles.modified_formula_style_id),
        SmlChangeType::FormatChanged => Some(styles.modified_format_style_id),
        _ => None,
    };

    // Apply style
    if let Some(sid) = style_id {
        ws_doc.set_attribute(cell_id, &XName::local("s"), &sid.to_string());
    }
}

/// Find or create a row element at the given index
fn find_or_create_row(
    doc: &mut XmlDocument,
    sheet_data_id: indextree::NodeId,
    row_index: usize,
    row_name: &XName,
) -> indextree::NodeId {
    // Look for existing row
    for child_id in doc.children(sheet_data_id).collect::<Vec<_>>() {
        if let Some(data) = doc.get(child_id) {
            if data.name() == Some(row_name) {
                if let Some(attrs) = data.attributes() {
                    for attr in attrs {
                        if attr.name.local_name == "r" {
                            if let Ok(r) = attr.value.parse::<usize>() {
                                if r == row_index {
                                    return child_id;
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Create new row
    let new_row = doc.add_child(
        sheet_data_id,
        XmlNodeData::element_with_attrs(
            row_name.clone(),
            vec![XAttribute::new(XName::local("r"), &row_index.to_string())],
        ),
    );

    new_row
}

/// Find or create a cell element at the given address
fn find_or_create_cell(
    doc: &mut XmlDocument,
    row_id: indextree::NodeId,
    cell_address: &str,
    _col_index: usize,
    cell_name: &XName,
) -> indextree::NodeId {
    // Look for existing cell
    for child_id in doc.children(row_id).collect::<Vec<_>>() {
        if let Some(data) = doc.get(child_id) {
            if data.name() == Some(cell_name) {
                if let Some(attrs) = data.attributes() {
                    for attr in attrs {
                        if attr.name.local_name == "r" && attr.value == cell_address {
                            return child_id;
                        }
                    }
                }
            }
        }
    }

    // Create new cell
    let new_cell = doc.add_child(
        row_id,
        XmlNodeData::element_with_attrs(
            cell_name.clone(),
            vec![XAttribute::new(XName::local("r"), cell_address)],
        ),
    );

    new_cell
}

/// Add comments for changes to a worksheet
fn add_comments_for_changes(
    pkg: &mut crate::package::OoxmlPackage,
    sheet_path: &str,
    changes: &[&SmlChange],
    settings: &SmlComparerSettings,
) -> Result<()> {
    if changes.is_empty() {
        return Ok(());
    }

    // Determine comments path from sheet path
    // e.g., xl/worksheets/sheet1.xml -> xl/worksheets/../comments1.xml
    let sheet_dir = sheet_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("xl");
    let sheet_name = sheet_path
        .rsplit_once('/')
        .map(|(_, n)| n)
        .unwrap_or(sheet_path);
    let sheet_num = sheet_name
        .trim_start_matches("sheet")
        .trim_end_matches(".xml");
    let comments_path = format!("{}/comments{}.xml", sheet_dir, sheet_num);

    // Create or get comments document
    let mut comments_doc = if let Ok(doc) = pkg.get_xml_part(&comments_path) {
        doc
    } else {
        create_comments_document(&settings.author_for_changes)
    };

    // Find commentList
    let Some(root) = comments_doc.root() else {
        return Ok(());
    };

    let comment_list_name = S::commentList();
    let comment_list_id = comments_doc
        .find_child(root, &comment_list_name)
        .unwrap_or_else(|| {
            comments_doc.add_child(root, XmlNodeData::element(comment_list_name.clone()))
        });

    // Add comments for each change
    for change in changes {
        if let Some(cell_address) = &change.cell_address {
            let comment_text = build_comment_text(change);
            add_comment(
                &mut comments_doc,
                comment_list_id,
                cell_address,
                &comment_text,
            );
        }
    }

    pkg.put_xml_part(&comments_path, &comments_doc)?;

    // Add VML drawing for comments (required for Excel to display them)
    add_vml_drawing_for_comments(pkg, sheet_path, &comments_path, changes)?;

    Ok(())
}

/// Create a new comments document
fn create_comments_document(author: &str) -> XmlDocument {
    let mut doc = XmlDocument::new();

    let root = doc.add_root(XmlNodeData::element_with_attrs(
        S::comments(),
        vec![XAttribute::new(XName::new(XMLNS::NS, "x"), S::NS)],
    ));

    // Add authors
    let authors = doc.add_child(root, XmlNodeData::element(S::authors()));
    let author_elem = doc.add_child(authors, XmlNodeData::element(S::author()));
    doc.add_child(author_elem, XmlNodeData::text(author));

    // Add commentList
    doc.add_child(root, XmlNodeData::element(S::commentList()));

    doc
}

/// Add a single comment element
fn add_comment(
    doc: &mut XmlDocument,
    comment_list_id: indextree::NodeId,
    cell_ref: &str,
    text: &str,
) {
    let comment = doc.add_child(
        comment_list_id,
        XmlNodeData::element_with_attrs(
            S::comment(),
            vec![
                XAttribute::new(XName::local("ref"), cell_ref),
                XAttribute::new(XName::local("authorId"), "0"),
            ],
        ),
    );

    let text_elem = doc.add_child(comment, XmlNodeData::element(S::text()));
    let r_elem = doc.add_child(text_elem, XmlNodeData::element(S::r()));
    let t_elem = doc.add_child(r_elem, XmlNodeData::element(S::t()));
    doc.add_child(t_elem, XmlNodeData::text(text));
}

/// Build comment text for a change
fn build_comment_text(change: &SmlChange) -> String {
    let mut lines = vec![format!("[{:?}]", change.change_type)];

    match change.change_type {
        SmlChangeType::CellAdded => {
            if let Some(new_value) = &change.new_value {
                lines.push(format!("New value: {}", new_value));
            }
            if let Some(new_formula) = &change.new_formula {
                lines.push(format!("Formula: ={}", new_formula));
            }
        }
        SmlChangeType::ValueChanged => {
            if let Some(old_value) = &change.old_value {
                lines.push(format!("Old value: {}", old_value));
            }
            if let Some(new_value) = &change.new_value {
                lines.push(format!("New value: {}", new_value));
            }
        }
        SmlChangeType::FormulaChanged => {
            if let Some(old_formula) = &change.old_formula {
                lines.push(format!("Old formula: ={}", old_formula));
            }
            if let Some(new_formula) = &change.new_formula {
                lines.push(format!("New formula: ={}", new_formula));
            }
        }
        SmlChangeType::FormatChanged => {
            if let (Some(new_format), Some(old_format)) = (&change.new_format, &change.old_format) {
                lines.push(new_format.get_difference_description(old_format));
            }
        }
        _ => {}
    }

    lines.join("\n")
}

/// Add VML drawing part for comment display (required by Excel)
fn add_vml_drawing_for_comments(
    pkg: &mut crate::package::OoxmlPackage,
    sheet_path: &str,
    comments_path: &str,
    changes: &[&SmlChange],
) -> Result<()> {
    // Determine VML path
    let sheet_dir = sheet_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("xl");
    let sheet_name = sheet_path
        .rsplit_once('/')
        .map(|(_, n)| n)
        .unwrap_or(sheet_path);
    let sheet_num = sheet_name
        .trim_start_matches("sheet")
        .trim_end_matches(".xml");
    let vml_path = format!("{}/vmlDrawing{}.vml", sheet_dir, sheet_num);

    // Build VML content
    let mut vml = String::new();
    vml.push_str(r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">"#);
    vml.push('\n');
    vml.push_str(r#"<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>"#);
    vml.push('\n');
    vml.push_str(r#"<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">"#);
    vml.push('\n');
    vml.push_str(
        r#"<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>"#,
    );
    vml.push('\n');
    vml.push_str("</v:shapetype>\n");

    let mut shape_id = 1024;
    for change in changes {
        if let Some(cell_address) = &change.cell_address {
            let (col, row) = parse_cell_ref(cell_address);
            vml.push_str(&format!(
                "<v:shape id=\"_x0000_s{}\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:80pt;margin-top:5pt;width:120pt;height:60pt;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">",
                shape_id
            ));
            vml.push('\n');
            vml.push_str("<v:fill color2=\"#ffffe1\"/>");
            vml.push('\n');
            vml.push_str("<v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>");
            vml.push('\n');
            vml.push_str("<v:path o:connecttype=\"none\"/>");
            vml.push('\n');
            vml.push_str("<v:textbox style=\"mso-direction-alt:auto\"/>");
            vml.push('\n');
            vml.push_str(&format!(
                "<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{}, 0, {}, 0, {}, 0, {}, 0</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{}</x:Row><x:Column>{}</x:Column></x:ClientData>",
                col.saturating_sub(1), row.saturating_sub(1), col + 1, row + 3, row.saturating_sub(1), col.saturating_sub(1)
            ));
            vml.push('\n');
            vml.push_str("</v:shape>\n");
            shape_id += 1;
        }
    }

    vml.push_str("</xml>");

    pkg.set_part(&vml_path, vml.into_bytes());

    // Add legacyDrawing relationship to worksheet
    let mut ws_doc = pkg.get_xml_part(sheet_path)?;
    if let Some(root) = ws_doc.root() {
        let legacy_drawing_name = S::legacyDrawing();
        if ws_doc.find_child(root, &legacy_drawing_name).is_none() {
            // Add relationship ID (simplified - in real impl would manage rels properly)
            let rel_id = format!("rId{}", shape_id);
            ws_doc.add_child(
                root,
                XmlNodeData::element_with_attrs(
                    legacy_drawing_name,
                    vec![XAttribute::new(XName::new(R::NS, "id"), &rel_id)],
                ),
            );
            pkg.put_xml_part(sheet_path, &ws_doc)?;
        }
    }

    Ok(())
}

/// Add a summary worksheet with change statistics
fn add_diff_summary_sheet(
    pkg: &mut crate::package::OoxmlPackage,
    result: &SmlComparisonResult,
    _settings: &SmlComparerSettings,
) -> Result<()> {
    // Create worksheet content
    let mut ws_doc = XmlDocument::new();

    let root = ws_doc.add_root(XmlNodeData::element_with_attrs(
        S::worksheet(),
        vec![
            XAttribute::new(XName::new(XMLNS::NS, "x"), S::NS),
            XAttribute::new(XName::new(XMLNS::NS, "r"), R::NS),
        ],
    ));

    let sheet_data = ws_doc.add_child(root, XmlNodeData::element(S::sheetData()));

    // Add summary header
    let mut row_num = 1;
    add_row(
        &mut ws_doc,
        sheet_data,
        row_num,
        &["Spreadsheet Comparison Summary"],
    );
    row_num += 1;
    add_row(&mut ws_doc, sheet_data, row_num, &[""]);
    row_num += 1;
    add_row(
        &mut ws_doc,
        sheet_data,
        row_num,
        &["Total Changes:", &result.total_changes().to_string()],
    );
    row_num += 1;
    add_row(
        &mut ws_doc,
        sheet_data,
        row_num,
        &["Value Changes:", &result.value_changes().to_string()],
    );
    row_num += 1;
    add_row(
        &mut ws_doc,
        sheet_data,
        row_num,
        &["Formula Changes:", &result.formula_changes().to_string()],
    );
    row_num += 1;
    add_row(
        &mut ws_doc,
        sheet_data,
        row_num,
        &["Format Changes:", &result.format_changes().to_string()],
    );
    row_num += 1;
    add_row(
        &mut ws_doc,
        sheet_data,
        row_num,
        &["Cells Added:", &result.cells_added().to_string()],
    );
    row_num += 1;
    add_row(
        &mut ws_doc,
        sheet_data,
        row_num,
        &["Cells Deleted:", &result.cells_deleted().to_string()],
    );
    row_num += 1;
    add_row(
        &mut ws_doc,
        sheet_data,
        row_num,
        &["Sheets Added:", &result.sheets_added().to_string()],
    );
    row_num += 1;
    add_row(
        &mut ws_doc,
        sheet_data,
        row_num,
        &["Sheets Deleted:", &result.sheets_deleted().to_string()],
    );
    row_num += 1;
    add_row(&mut ws_doc, sheet_data, row_num, &[""]);
    row_num += 1;

    // Add change details header
    add_row(
        &mut ws_doc,
        sheet_data,
        row_num,
        &[
            "Change Type",
            "Sheet",
            "Cell",
            "Old Value",
            "New Value",
            "Description",
        ],
    );
    row_num += 1;

    // Add each change
    for change in &result.changes {
        let desc = change.get_description();
        add_row(
            &mut ws_doc,
            sheet_data,
            row_num,
            &[
                &format!("{:?}", change.change_type),
                change.sheet_name.as_deref().unwrap_or(""),
                change.cell_address.as_deref().unwrap_or(""),
                change
                    .old_value
                    .as_deref()
                    .or(change.old_formula.as_deref())
                    .unwrap_or(""),
                change
                    .new_value
                    .as_deref()
                    .or(change.new_formula.as_deref())
                    .unwrap_or(""),
                &desc,
            ],
        );
        row_num += 1;
    }

    // Save to a new worksheet path
    let summary_path = "xl/worksheets/_DiffSummary.xml";
    pkg.put_xml_part(summary_path, &ws_doc)?;

    // Add sheet to workbook.xml
    let workbook_path = "xl/workbook.xml";
    let mut workbook_doc = pkg.get_xml_part(workbook_path)?;

    if let Some(wb_root) = workbook_doc.root() {
        let sheets_name = S::sheets();
        if let Some(sheets_id) = workbook_doc.find_child(wb_root, &sheets_name) {
            // Find max sheet ID
            let mut max_sheet_id: u32 = 0;
            for child_id in workbook_doc.children(sheets_id) {
                if let Some(data) = workbook_doc.get(child_id) {
                    if let Some(attrs) = data.attributes() {
                        for attr in attrs {
                            if attr.name.local_name == "sheetId" {
                                if let Ok(id) = attr.value.parse::<u32>() {
                                    max_sheet_id = max_sheet_id.max(id);
                                }
                            }
                        }
                    }
                }
            }

            // Add new sheet element
            let new_sheet_id = max_sheet_id + 1;
            workbook_doc.add_child(
                sheets_id,
                XmlNodeData::element_with_attrs(
                    S::sheet(),
                    vec![
                        XAttribute::new(XName::local("name"), "_DiffSummary"),
                        XAttribute::new(XName::local("sheetId"), &new_sheet_id.to_string()),
                        XAttribute::new(XName::new(R::NS, "id"), "rIdSummary"),
                    ],
                ),
            );

            pkg.put_xml_part(workbook_path, &workbook_doc)?;
        }
    }

    // Add relationship for the new sheet
    let rels_path = "xl/_rels/workbook.xml.rels";
    if let Ok(mut rels_doc) = pkg.get_xml_part(rels_path) {
        if let Some(rels_root) = rels_doc.root() {
            rels_doc.add_child(rels_root, XmlNodeData::element_with_attrs(
                XName::new("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship"),
                vec![
                    XAttribute::new(XName::local("Id"), "rIdSummary"),
                    XAttribute::new(XName::local("Type"), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                    XAttribute::new(XName::local("Target"), "worksheets/_DiffSummary.xml"),
                ],
            ));
            pkg.put_xml_part(rels_path, &rels_doc)?;
        }
    }

    Ok(())
}

/// Add a row with values to sheetData
fn add_row(
    doc: &mut XmlDocument,
    sheet_data_id: indextree::NodeId,
    row_num: usize,
    values: &[&str],
) {
    let row = doc.add_child(
        sheet_data_id,
        XmlNodeData::element_with_attrs(
            S::row(),
            vec![XAttribute::new(XName::local("r"), &row_num.to_string())],
        ),
    );

    for (i, value) in values.iter().enumerate() {
        let col_letter = get_column_letter(i + 1);
        let cell_ref = format!("{}{}", col_letter, row_num);

        let cell = doc.add_child(
            row,
            XmlNodeData::element_with_attrs(
                S::c(),
                vec![
                    XAttribute::new(XName::local("r"), &cell_ref),
                    XAttribute::new(XName::local("t"), "inlineStr"),
                ],
            ),
        );

        let is_elem = doc.add_child(cell, XmlNodeData::element(S::is()));
        let t_elem = doc.add_child(is_elem, XmlNodeData::element(S::t()));
        doc.add_child(t_elem, XmlNodeData::text(value));
    }
}

/// Parse cell reference like "A1" into (column, row).
fn parse_cell_ref(cell_ref: &str) -> (usize, usize) {
    let mut col = 0;
    let mut i = 0;
    let chars: Vec<char> = cell_ref.chars().collect();

    // Parse column letters (A=1, Z=26, AA=27, etc.)
    while i < chars.len() && chars[i].is_alphabetic() {
        col = col * 26 + (chars[i].to_ascii_uppercase() as usize - 'A' as usize + 1);
        i += 1;
    }

    // Parse row number
    let row_str: String = chars[i..].iter().collect();
    let row = row_str.parse::<usize>().unwrap_or(0);

    (col, row)
}

/// Get column letter from column number (1=A, 26=Z, 27=AA, etc.).
fn get_column_letter(mut column_number: usize) -> String {
    let mut result = String::new();

    while column_number > 0 {
        column_number -= 1;
        result.insert(0, (b'A' + (column_number % 26) as u8) as char);
        column_number /= 26;
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_cell_ref_works() {
        assert_eq!(parse_cell_ref("A1"), (1, 1));
        assert_eq!(parse_cell_ref("Z1"), (26, 1));
        assert_eq!(parse_cell_ref("AA1"), (27, 1));
        assert_eq!(parse_cell_ref("AB10"), (28, 10));
    }

    #[test]
    fn get_column_letter_works() {
        assert_eq!(get_column_letter(1), "A");
        assert_eq!(get_column_letter(26), "Z");
        assert_eq!(get_column_letter(27), "AA");
        assert_eq!(get_column_letter(28), "AB");
    }

    #[test]
    fn build_comment_text_formats_correctly() {
        let mut change = SmlChange::default();
        change.change_type = SmlChangeType::ValueChanged;
        change.old_value = Some("10".to_string());
        change.new_value = Some("20".to_string());

        let text = build_comment_text(&change);
        assert!(text.contains("ValueChanged"));
        assert!(text.contains("Old value: 10"));
        assert!(text.contains("New value: 20"));
    }
}
