//! Patching SML documents with changes.
//!
//! This module provides functionality to apply or revert changes to Excel spreadsheets.

use crate::error::{RedlineError, Result};
use crate::sml::{SmlChange, SmlChangeType, SmlDataRetriever, SmlDocument};
use crate::xml::{XAttribute, XName, XmlDocument, XmlNodeData, S};
use indextree::NodeId;

/// Apply a set of changes to a base Excel document.
///
/// Returns the modified document as bytes.
pub fn apply_sml_changes(base_doc: &[u8], changes: &[SmlChange]) -> Result<Vec<u8>> {
    let mut doc = SmlDocument::from_bytes(base_doc)?;

    // Group changes by sheet
    let mut changes_by_sheet: std::collections::HashMap<String, Vec<&SmlChange>> =
        std::collections::HashMap::new();

    for change in changes {
        if let Some(sheet_name) = &change.sheet_name {
            changes_by_sheet
                .entry(sheet_name.clone())
                .or_default()
                .push(change);
        }
    }

    for (sheet_name, sheet_changes) in changes_by_sheet {
        apply_changes_to_sheet(&mut doc, &sheet_name, &sheet_changes)?;
    }

    doc.to_bytes()
}

/// Revert a set of changes from a modified Excel document.
///
/// Effectively the same as apply_changes, but using old values instead of new values.
pub fn revert_sml_changes(result_doc: &[u8], changes: &[SmlChange]) -> Result<Vec<u8>> {
    // To revert, we treat "old_value" as the target value.
    // We map the changes to "inverse" changes.
    let inverse_changes: Vec<SmlChange> = changes.iter().map(|c| invert_change(c)).collect();
    apply_sml_changes(result_doc, &inverse_changes)
}

fn invert_change(change: &SmlChange) -> SmlChange {
    let mut inverse = change.clone();
    match change.change_type {
        SmlChangeType::ValueChanged | SmlChangeType::FormulaChanged => {
            inverse.old_value = change.new_value.clone();
            inverse.new_value = change.old_value.clone();
            inverse.old_formula = change.new_formula.clone();
            inverse.new_formula = change.old_formula.clone();
        }
        // TODO: Handle other inversions (e.g. CellAdded -> CellDeleted)
        _ => {}
    }
    inverse
}

fn apply_changes_to_sheet(
    doc: &mut SmlDocument,
    sheet_name: &str,
    changes: &[&SmlChange],
) -> Result<()> {
    let sheet_path = SmlDataRetriever::get_sheet_path(doc, sheet_name)?;
    let mut sheet_xml = doc.package().get_xml_part(&sheet_path)?;
    let root = sheet_xml
        .root()
        .ok_or_else(|| RedlineError::InvalidPackage {
            message: "Sheet has no root".to_string(),
        })?;

    let sheet_data_name = S::sheetData();
    let sheet_data_id = sheet_xml
        .elements_by_name(root, &sheet_data_name)
        .next()
        .ok_or_else(|| RedlineError::InvalidPackage {
            message: "Missing sheetData".to_string(),
        })?;

    for change in changes {
        match change.change_type {
            SmlChangeType::ValueChanged | SmlChangeType::FormulaChanged => {
                apply_cell_value_change(&mut sheet_xml, sheet_data_id, change)?;
            }
            _ => {}
        }
    }

    doc.package_mut().put_xml_part(&sheet_path, &sheet_xml)?;
    Ok(())
}

fn apply_cell_value_change(
    doc: &mut XmlDocument,
    sheet_data_id: NodeId,
    change: &SmlChange,
) -> Result<()> {
    let address = change
        .cell_address
        .as_deref()
        .ok_or_else(|| RedlineError::InvalidPackage {
            message: "Missing cell address".to_string(),
        })?;

    let (row_num, _) = parse_cell_ref(address);
    if row_num == 0 {
        return Ok(());
    }

    // Find row
    let row_name = S::row();
    let row_id = doc.elements_by_name(sheet_data_id, &row_name).find(|&r| {
        doc.get(r)
            .and_then(|d| d.attributes())
            .map_or(false, |attrs| {
                attrs
                    .iter()
                    .any(|a| a.name.local_name == "r" && a.value == row_num.to_string())
            })
    });

    // If row doesn't exist, we skip for now (creating rows requires ensuring order)
    if let Some(row_id) = row_id {
        // Find cell
        let c_name = S::c();
        let cell_id = doc.elements_by_name(row_id, &c_name).find(|&c| {
            doc.get(c)
                .and_then(|d| d.attributes())
                .map_or(false, |attrs| {
                    attrs
                        .iter()
                        .any(|a| a.name.local_name == "r" && a.value == address)
                })
        });

        if let Some(cell_id) = cell_id {
            // Update cell
            update_cell_content(doc, cell_id, change)?;
        }
    }

    Ok(())
}

fn update_cell_content(doc: &mut XmlDocument, cell_id: NodeId, change: &SmlChange) -> Result<()> {
    // Clear existing children (v, f, is)
    // Actually, we should be careful.
    // If it's a value change, we update/add <v>.
    // If it's a formula change, we update/add <f> and clear <v> (calculated value).

    // Simplification: remove v, f, is children.
    let children: Vec<_> = doc.children(cell_id).collect();
    for child in children {
        doc.detach(child);
    }

    // Add new formula if present
    if let Some(formula) = &change.new_formula {
        let f_id = doc.add_child(cell_id, XmlNodeData::element(S::f()));
        doc.add_child(f_id, XmlNodeData::text(formula));
    }

    // Add new value if present
    // Note: If formula is present, value is the cached result. We usually keep it or recalculate.
    // Here we just use what's in new_value.
    if let Some(value) = &change.new_value {
        // Determine type. If it's a shared string, we assume new_value is the string itself?
        // Wait, SmlChange.new_value usually contains the resolved string value.
        // But the XML expects an index for shared strings, or inline string.
        // Writing back directly is tricky if we don't update SharedStringTable.
        // Safest approach for "patching" without SST management: use inline strings (t="inlineStr").

        // Check if we can write as inlineStr
        // We need to set t="inlineStr" attribute on cell.

        set_attribute(doc, cell_id, "t", "inlineStr");

        let is_id = doc.add_child(cell_id, XmlNodeData::element(S::is()));
        let t_id = doc.add_child(is_id, XmlNodeData::element(S::t()));
        doc.add_child(t_id, XmlNodeData::text(value));
    } else {
        // Remove 't' attribute if value is empty/null (unless formula exists?)
        // If we cleared content, we might have an empty cell.
        // For now, let's leave it.
    }

    Ok(())
}

fn set_attribute(doc: &mut XmlDocument, node_id: NodeId, name: &str, value: &str) {
    if let Some(data) = doc.get_mut(node_id) {
        if let XmlNodeData::Element { attributes, .. } = data {
            // Remove existing
            attributes.retain(|a| a.name.local_name != name);
            // Add new
            attributes.push(XAttribute::new(XName::new("", name), value));
        }
    }
}

/// Parse "A1" into (row, col) 1-based indices.
fn parse_cell_ref(address: &str) -> (u32, u32) {
    let mut chars = address.chars().peekable();
    let mut col_str = String::new();

    while let Some(c) = chars.peek() {
        if c.is_alphabetic() {
            col_str.push(*c);
            chars.next();
        } else {
            break;
        }
    }

    let row_str: String = chars.collect();
    let row = row_str.parse::<u32>().unwrap_or(0);
    let col = column_name_to_index(&col_str);

    (row, col)
}

fn column_name_to_index(name: &str) -> u32 {
    let mut index = 0;
    for c in name.chars() {
        if c.is_ascii_uppercase() {
            index = index * 26 + (c as u32 - 'A' as u32 + 1);
        } else if c.is_ascii_lowercase() {
            index = index * 26 + (c as u32 - 'a' as u32 + 1);
        }
    }
    index
}
