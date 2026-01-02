// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

//! Canonicalizes spreadsheets into a normalized form for comparison.
//!
//! This module converts an SmlDocument into a WorkbookSignature, which represents
//! the normalized, expanded form of the workbook for comparison purposes. This includes:
//! - Resolving shared strings
//! - Expanding style indices to full format signatures
//! - Computing content hashes for cells
//! - Extracting comments, data validations, merged cells, and hyperlinks (Phase 3)
//! - Computing row/column signatures for alignment (Phase 2)

use crate::error::{RedlineError, Result};
use crate::package::ooxml::OoxmlPackage;
use crate::sml::signatures::{
    CellFormatSignature, CellSignature, CommentSignature, DataValidationSignature,
    HyperlinkSignature, WorkbookSignature, WorksheetSignature,
};
use crate::sml::{SmlComparerSettings, SmlDocument};
use crate::xml::{XName, XmlDocument, R, S};
use indextree::NodeId;
use std::collections::HashMap;

/// Main canonicalizer for SmlDocument to WorkbookSignature conversion.
pub struct SmlCanonicalizer;

/// Helper function to extract text content from a node.
fn get_text_content(doc: &XmlDocument, node_id: NodeId) -> Option<String> {
    let mut text = String::new();
    for child_id in doc.children(node_id) {
        if let Some(child_data) = doc.get(child_id) {
            if let Some(content) = child_data.text_content() {
                text.push_str(content);
            }
        }
    }
    if text.is_empty() {
        None
    } else {
        Some(text)
    }
}

impl SmlCanonicalizer {
    /// Canonicalize an SmlDocument into a WorkbookSignature for comparison.
    ///
    /// This is the main entry point for converting a spreadsheet into its canonical form.
    /// The resulting signature contains all data needed for comparison.
    pub fn canonicalize(
        doc: &SmlDocument,
        settings: &SmlComparerSettings,
    ) -> Result<WorkbookSignature> {
        let mut signature = WorkbookSignature::new();

        let pkg = doc.package();

        // Load workbook.xml
        let workbook_path = "xl/workbook.xml";
        let workbook_doc = pkg.get_xml_part(workbook_path)?;
        let workbook_root = workbook_doc
            .root()
            .ok_or_else(|| RedlineError::InvalidPackage {
                message: "Workbook has no root".to_string(),
            })?;

        // Get shared strings
        let shared_strings = Self::get_shared_strings(pkg)?;

        // Get style info
        let style_info = Self::get_style_info(pkg)?;

        // Process each sheet
        let sheets_elem = workbook_doc
            .elements_by_name(workbook_root, &S::sheets())
            .next()
            .ok_or_else(|| RedlineError::InvalidPackage {
                message: "Missing sheets element in workbook".to_string(),
            })?;

        for sheet_elem in workbook_doc.elements_by_name(sheets_elem, &S::sheet()) {
            let sheet_data =
                workbook_doc
                    .get(sheet_elem)
                    .ok_or_else(|| RedlineError::InvalidPackage {
                        message: "Sheet element has no data".to_string(),
                    })?;

            let attrs = sheet_data
                .attributes()
                .ok_or_else(|| RedlineError::InvalidPackage {
                    message: "Sheet has no attributes".to_string(),
                })?;

            let sheet_name = attrs
                .iter()
                .find(|a| a.name.local_name == "name")
                .map(|a| a.value.clone())
                .ok_or_else(|| RedlineError::InvalidPackage {
                    message: "Sheet has no name attribute".to_string(),
                })?;

            let r_id = attrs
                .iter()
                .find(|a| a.name == R::id())
                .map(|a| a.value.clone())
                .ok_or_else(|| RedlineError::InvalidPackage {
                    message: "Sheet has no r:id attribute".to_string(),
                })?;

            // Resolve worksheet part path
            let sheet_path = Self::resolve_worksheet_path(pkg, &r_id)?;

            // Canonicalize worksheet
            let ws_signature = Self::canonicalize_worksheet(
                pkg,
                &sheet_path,
                &sheet_name,
                &r_id,
                &shared_strings,
                &style_info,
                settings,
            )?;

            signature.sheets.insert(sheet_name, ws_signature);
        }

        // Get defined names
        if let Some(defined_names_elem) = workbook_doc
            .elements_by_name(workbook_root, &S::definedNames())
            .next()
        {
            for dn_elem in workbook_doc.elements_by_name(defined_names_elem, &S::definedName()) {
                if let Some(dn_data) = workbook_doc.get(dn_elem) {
                    if let Some(attrs) = dn_data.attributes() {
                        if let Some(name_attr) = attrs.iter().find(|a| a.name.local_name == "name")
                        {
                            let name = name_attr.value.clone();
                            let value =
                                get_text_content(&workbook_doc, dn_elem).unwrap_or_default();
                            if !name.is_empty() {
                                signature.defined_names.insert(name, value);
                            }
                        }
                    }
                }
            }
        }

        Ok(signature)
    }

    /// Resolve the worksheet part path from a relationship ID.
    fn resolve_worksheet_path(pkg: &OoxmlPackage, r_id: &str) -> Result<String> {
        let workbook_rels_path = "xl/_rels/workbook.xml.rels";
        let rels_doc = pkg.get_xml_part(workbook_rels_path)?;
        let rels_root = rels_doc
            .root()
            .ok_or_else(|| RedlineError::InvalidPackage {
                message: "Workbook rels has no root".to_string(),
            })?;

        let rel_name = XName::new(
            "http://schemas.openxmlformats.org/package/2006/relationships",
            "Relationship",
        );

        for rel_elem in rels_doc.elements_by_name(rels_root, &rel_name) {
            if let Some(rel_data) = rels_doc.get(rel_elem) {
                if let Some(attrs) = rel_data.attributes() {
                    if let Some(id_attr) = attrs.iter().find(|a| a.name.local_name == "Id") {
                        if id_attr.value == r_id {
                            if let Some(target_attr) =
                                attrs.iter().find(|a| a.name.local_name == "Target")
                            {
                                let target = &target_attr.value;
                                return Ok(if target.starts_with('/') {
                                    target[1..].to_string()
                                } else {
                                    format!("xl/{}", target)
                                });
                            }
                        }
                    }
                }
            }
        }

        Err(RedlineError::InvalidPackage {
            message: format!("Relationship not found for id: {}", r_id),
        })
    }

    /// Canonicalize a single worksheet.
    #[allow(clippy::too_many_arguments)]
    fn canonicalize_worksheet(
        pkg: &OoxmlPackage,
        sheet_path: &str,
        sheet_name: &str,
        r_id: &str,
        shared_strings: &[String],
        style_info: &StyleInfo,
        settings: &SmlComparerSettings,
    ) -> Result<WorksheetSignature> {
        let mut signature = WorksheetSignature::new(sheet_name.to_string(), r_id.to_string());

        let ws_doc = pkg.get_xml_part(sheet_path)?;
        let ws_root = ws_doc.root().ok_or_else(|| RedlineError::InvalidPackage {
            message: "Worksheet has no root".to_string(),
        })?;

        // Get sheetData element
        let sheet_data = ws_doc.elements_by_name(ws_root, &S::sheetData()).next();
        if sheet_data.is_none() {
            return Ok(signature);
        }
        let sheet_data = sheet_data.unwrap();

        // Process rows and cells
        for row_elem in ws_doc.elements_by_name(sheet_data, &S::row()) {
            let row_data = ws_doc.get(row_elem);
            if row_data.is_none() {
                continue;
            }

            for cell_elem in ws_doc.elements_by_name(row_elem, &S::c()) {
                let cell_data = ws_doc.get(cell_elem);
                if cell_data.is_none() {
                    continue;
                }

                let attrs = cell_data.unwrap().attributes();
                if attrs.is_none() {
                    continue;
                }
                let attrs = attrs.unwrap();

                let cell_ref = attrs
                    .iter()
                    .find(|a| a.name.local_name == "r")
                    .map(|a| a.value.clone());

                if cell_ref.is_none() || cell_ref.as_ref().unwrap().is_empty() {
                    continue;
                }
                let cell_ref = cell_ref.unwrap();

                let cell_sig = Self::canonicalize_cell(
                    &ws_doc,
                    cell_elem,
                    &cell_ref,
                    shared_strings,
                    style_info,
                    settings,
                )?;

                // Phase 2: Track populated rows and columns
                signature.populated_rows.insert(cell_sig.row);
                signature.populated_columns.insert(cell_sig.column);

                signature.cells.insert(cell_ref, cell_sig);
            }
        }

        // Phase 2: Compute row signatures for alignment
        if settings.enable_row_alignment {
            Self::compute_row_signatures(&mut signature, settings);
        }

        // Phase 2: Compute column signatures for alignment
        if settings.enable_column_alignment {
            Self::compute_column_signatures(&mut signature, settings);
        }

        // Phase 3: Extract comments
        if settings.compare_comments {
            Self::extract_comments(pkg, sheet_path, &mut signature)?;
        }

        // Phase 3: Extract data validations
        if settings.compare_data_validation {
            Self::extract_data_validations(&ws_doc, ws_root, &mut signature)?;
        }

        // Phase 3: Extract merged cells
        if settings.compare_merged_cells {
            Self::extract_merged_cells(&ws_doc, ws_root, &mut signature)?;
        }

        // Phase 3: Extract hyperlinks
        if settings.compare_hyperlinks {
            Self::extract_hyperlinks(pkg, sheet_path, &ws_doc, ws_root, &mut signature)?;
        }

        Ok(signature)
    }

    /// Canonicalize a single cell.
    fn canonicalize_cell(
        ws_doc: &XmlDocument,
        cell_elem: NodeId,
        cell_ref: &str,
        shared_strings: &[String],
        style_info: &StyleInfo,
        settings: &SmlComparerSettings,
    ) -> Result<CellSignature> {
        // Parse cell reference
        let (col, row) = Self::parse_cell_reference(cell_ref)?;

        // Get cell data
        let cell_data = ws_doc
            .get(cell_elem)
            .ok_or_else(|| RedlineError::InvalidPackage {
                message: "Cell has no data".to_string(),
            })?;

        // Get value
        let resolved_value = Self::resolve_value(ws_doc, cell_elem, shared_strings)?;

        // Get formula
        let formula = ws_doc
            .elements_by_name(cell_elem, &S::f())
            .next()
            .and_then(|f_elem| get_text_content(ws_doc, f_elem));

        // Get format
        let style_index = cell_data
            .attributes()
            .and_then(|attrs| attrs.iter().find(|a| a.name.local_name == "s"))
            .and_then(|a| a.value.parse::<usize>().ok())
            .unwrap_or(0);

        let format = Self::expand_style(style_index, style_info);

        let content_hash =
            CellSignature::compute_hash(resolved_value.as_deref(), formula.as_deref());

        Ok(CellSignature {
            address: cell_ref.to_string(),
            row,
            column: col,
            resolved_value,
            formula,
            content_hash,
            format,
        })
    }

    /// Resolve the value of a cell, handling shared strings and inline strings.
    fn resolve_value(
        ws_doc: &XmlDocument,
        cell_elem: NodeId,
        shared_strings: &[String],
    ) -> Result<Option<String>> {
        let cell_data = ws_doc
            .get(cell_elem)
            .ok_or_else(|| RedlineError::InvalidPackage {
                message: "Cell has no data".to_string(),
            })?;

        let cell_type = cell_data
            .attributes()
            .and_then(|attrs| attrs.iter().find(|a| a.name.local_name == "t"))
            .map(|a| a.value.as_str())
            .unwrap_or("");

        // Get <v> element
        let value_elem = ws_doc.elements_by_name(cell_elem, &S::v()).next();
        let raw_value = value_elem.and_then(|v| get_text_content(ws_doc, v));

        if raw_value.is_none() || raw_value.as_ref().unwrap().is_empty() {
            // Check for inline string
            if let Some(is_elem) = ws_doc.elements_by_name(cell_elem, &S::is()).next() {
                let mut text = String::new();
                for t_elem in ws_doc.descendants(is_elem) {
                    if let Some(t_data) = ws_doc.get(t_elem) {
                        if let Some(name) = t_data.name() {
                            if name == &S::t() {
                                if let Some(content) = get_text_content(ws_doc, t_elem) {
                                    text.push_str(&content);
                                }
                            }
                        }
                    }
                }
                return Ok(if text.is_empty() { None } else { Some(text) });
            }
            return Ok(None);
        }

        let raw_value = raw_value.unwrap();

        Ok(Some(match cell_type {
            "s" => Self::resolve_shared_string(&raw_value, shared_strings)?,
            "str" => raw_value,
            "b" => {
                if raw_value == "1" {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            "e" => raw_value, // Error value
            _ => Self::normalize_numeric(&raw_value),
        }))
    }

    /// Resolve a shared string index to its value.
    fn resolve_shared_string(index_str: &str, shared_strings: &[String]) -> Result<String> {
        if let Ok(index) = index_str.parse::<usize>() {
            if index < shared_strings.len() {
                return Ok(shared_strings[index].clone());
            }
        }
        Ok(index_str.to_string())
    }

    /// Normalize numeric values to consistent representation.
    fn normalize_numeric(value: &str) -> String {
        if value.is_empty() {
            return value.to_string();
        }

        // Try to parse as decimal and normalize
        if let Ok(d) = value.parse::<f64>() {
            // Use Rust's default formatting for consistency
            return d.to_string();
        }

        value.to_string()
    }

    /// Parse a cell reference like "A1" into (column, row).
    fn parse_cell_reference(cell_ref: &str) -> Result<(i32, i32)> {
        let mut col = 0;
        let mut row = 0;
        let mut i = 0;

        // Parse column letters
        for ch in cell_ref.chars() {
            if ch.is_ascii_alphabetic() {
                col = col * 26 + (ch.to_ascii_uppercase() as i32 - 'A' as i32 + 1);
                i += 1;
            } else {
                break;
            }
        }

        // Parse row number
        if i < cell_ref.len() {
            row = cell_ref[i..]
                .parse::<i32>()
                .map_err(|_| RedlineError::InvalidPackage {
                    message: format!("Invalid cell reference: {}", cell_ref),
                })?;
        }

        Ok((col, row))
    }

    /// Get shared strings table.
    fn get_shared_strings(pkg: &OoxmlPackage) -> Result<Vec<String>> {
        let mut result = Vec::new();

        let sst_path = "xl/sharedStrings.xml";
        let sst_doc = match pkg.get_xml_part(sst_path) {
            Ok(doc) => doc,
            Err(_) => return Ok(result), // No shared strings is valid
        };

        let sst_root = sst_doc.root().ok_or_else(|| RedlineError::InvalidPackage {
            message: "SharedStrings has no root".to_string(),
        })?;

        for si_elem in sst_doc.elements_by_name(sst_root, &S::si()) {
            let mut text = String::new();

            // Collect all <t> elements
            for descendant in sst_doc.descendants(si_elem) {
                if let Some(d_data) = sst_doc.get(descendant) {
                    if let Some(name) = d_data.name() {
                        if name == &S::t() {
                            if let Some(content) = get_text_content(&sst_doc, descendant) {
                                text.push_str(&content);
                            }
                        }
                    }
                }
            }

            result.push(text);
        }

        Ok(result)
    }

    /// Get style information from styles.xml.
    fn get_style_info(pkg: &OoxmlPackage) -> Result<StyleInfo> {
        let mut info = StyleInfo::default();

        let styles_path = "xl/styles.xml";
        let styles_doc = match pkg.get_xml_part(styles_path) {
            Ok(doc) => doc,
            Err(_) => return Ok(info), // No styles is valid
        };

        let styles_root = styles_doc
            .root()
            .ok_or_else(|| RedlineError::InvalidPackage {
                message: "Styles has no root".to_string(),
            })?;

        // Get number formats
        if let Some(num_fmts_elem) = styles_doc
            .elements_by_name(styles_root, &S::numFmts())
            .next()
        {
            for num_fmt_elem in styles_doc.elements_by_name(num_fmts_elem, &S::numFmt()) {
                if let Some(fmt_data) = styles_doc.get(num_fmt_elem) {
                    if let Some(attrs) = fmt_data.attributes() {
                        let id = attrs
                            .iter()
                            .find(|a| a.name.local_name == "numFmtId")
                            .and_then(|a| a.value.parse::<i32>().ok())
                            .unwrap_or(0);

                        let code = attrs
                            .iter()
                            .find(|a| a.name.local_name == "formatCode")
                            .map(|a| a.value.clone())
                            .unwrap_or_default();

                        info.number_formats.insert(id, code);
                    }
                }
            }
        }

        // Get fonts
        if let Some(fonts_elem) = styles_doc.elements_by_name(styles_root, &S::fonts()).next() {
            for font_elem in styles_doc.elements_by_name(fonts_elem, &S::font()) {
                let font = Self::parse_font(&styles_doc, font_elem);
                info.fonts.push(font);
            }
        }

        // Get fills
        if let Some(fills_elem) = styles_doc.elements_by_name(styles_root, &S::fills()).next() {
            for fill_elem in styles_doc.elements_by_name(fills_elem, &S::fill()) {
                let fill = Self::parse_fill(&styles_doc, fill_elem);
                info.fills.push(fill);
            }
        }

        // Get borders
        if let Some(borders_elem) = styles_doc
            .elements_by_name(styles_root, &S::borders())
            .next()
        {
            for border_elem in styles_doc.elements_by_name(borders_elem, &S::border()) {
                let border = Self::parse_border(&styles_doc, border_elem);
                info.borders.push(border);
            }
        }

        // Get cell XFs
        if let Some(cell_xfs_elem) = styles_doc
            .elements_by_name(styles_root, &S::cellXfs())
            .next()
        {
            for xf_elem in styles_doc.elements_by_name(cell_xfs_elem, &S::xf()) {
                let xf = Self::parse_xf(&styles_doc, xf_elem);
                info.cell_xfs.push(xf);
            }
        }

        Ok(info)
    }

    /// Parse a font element.
    fn parse_font(doc: &XmlDocument, font_elem: NodeId) -> FontInfo {
        let mut font = FontInfo::default();

        for child in doc.children(font_elem) {
            if let Some(child_data) = doc.get(child) {
                if let Some(name) = child_data.name() {
                    match name.local_name.as_str() {
                        "b" => font.bold = true,
                        "i" => font.italic = true,
                        "u" => font.underline = true,
                        "strike" => font.strikethrough = true,
                        "name" => {
                            if let Some(attrs) = child_data.attributes() {
                                font.name = attrs
                                    .iter()
                                    .find(|a| a.name.local_name == "val")
                                    .map(|a| a.value.clone());
                            }
                        }
                        "sz" => {
                            if let Some(attrs) = child_data.attributes() {
                                font.size = attrs
                                    .iter()
                                    .find(|a| a.name.local_name == "val")
                                    .and_then(|a| a.value.parse::<f64>().ok());
                            }
                        }
                        "color" => {
                            if let Some(attrs) = child_data.attributes() {
                                font.color = attrs
                                    .iter()
                                    .find(|a| {
                                        a.name.local_name == "rgb" || a.name.local_name == "theme"
                                    })
                                    .map(|a| a.value.clone());
                            }
                        }
                        _ => {}
                    }
                }
            }
        }

        font
    }

    /// Parse a fill element.
    fn parse_fill(doc: &XmlDocument, fill_elem: NodeId) -> FillInfo {
        let mut fill = FillInfo::default();

        // Look for patternFill
        if let Some(pattern_fill) = doc
            .elements_by_name(fill_elem, &XName::new(S::NS, "patternFill"))
            .next()
        {
            if let Some(pf_data) = doc.get(pattern_fill) {
                if let Some(attrs) = pf_data.attributes() {
                    fill.pattern = attrs
                        .iter()
                        .find(|a| a.name.local_name == "patternType")
                        .map(|a| a.value.clone());
                }

                // Foreground color
                if let Some(fg_elem) = doc
                    .elements_by_name(pattern_fill, &XName::new(S::NS, "fgColor"))
                    .next()
                {
                    if let Some(fg_data) = doc.get(fg_elem) {
                        if let Some(attrs) = fg_data.attributes() {
                            fill.fg_color = attrs
                                .iter()
                                .find(|a| {
                                    a.name.local_name == "rgb" || a.name.local_name == "theme"
                                })
                                .map(|a| a.value.clone());
                        }
                    }
                }

                // Background color
                if let Some(bg_elem) = doc
                    .elements_by_name(pattern_fill, &XName::new(S::NS, "bgColor"))
                    .next()
                {
                    if let Some(bg_data) = doc.get(bg_elem) {
                        if let Some(attrs) = bg_data.attributes() {
                            fill.bg_color = attrs
                                .iter()
                                .find(|a| {
                                    a.name.local_name == "rgb" || a.name.local_name == "theme"
                                })
                                .map(|a| a.value.clone());
                        }
                    }
                }
            }
        }

        fill
    }

    /// Parse a border element.
    fn parse_border(doc: &XmlDocument, border_elem: NodeId) -> BorderInfo {
        let mut border = BorderInfo::default();

        for child in doc.children(border_elem) {
            if let Some(child_data) = doc.get(child) {
                if let Some(name) = child_data.name() {
                    let (style, color) = Self::parse_border_side(doc, child);

                    match name.local_name.as_str() {
                        "left" => {
                            border.left_style = style;
                            border.left_color = color;
                        }
                        "right" => {
                            border.right_style = style;
                            border.right_color = color;
                        }
                        "top" => {
                            border.top_style = style;
                            border.top_color = color;
                        }
                        "bottom" => {
                            border.bottom_style = style;
                            border.bottom_color = color;
                        }
                        _ => {}
                    }
                }
            }
        }

        border
    }

    /// Parse a border side (left, right, top, bottom).
    fn parse_border_side(doc: &XmlDocument, side_elem: NodeId) -> (Option<String>, Option<String>) {
        let side_data = doc.get(side_elem);
        if side_data.is_none() {
            return (None, None);
        }

        let style = side_data
            .and_then(|d| d.attributes())
            .and_then(|attrs| attrs.iter().find(|a| a.name.local_name == "style"))
            .map(|a| a.value.clone());

        let color = doc
            .elements_by_name(side_elem, &XName::new(S::NS, "color"))
            .next()
            .and_then(|color_elem| doc.get(color_elem))
            .and_then(|color_data| color_data.attributes())
            .and_then(|attrs| {
                attrs
                    .iter()
                    .find(|a| a.name.local_name == "rgb" || a.name.local_name == "theme")
            })
            .map(|a| a.value.clone());

        (style, color)
    }

    /// Parse an XF (cell format) element.
    fn parse_xf(doc: &XmlDocument, xf_elem: NodeId) -> XfInfo {
        let mut xf = XfInfo::default();

        if let Some(xf_data) = doc.get(xf_elem) {
            if let Some(attrs) = xf_data.attributes() {
                xf.num_fmt_id = attrs
                    .iter()
                    .find(|a| a.name.local_name == "numFmtId")
                    .and_then(|a| a.value.parse::<i32>().ok());

                xf.font_id = attrs
                    .iter()
                    .find(|a| a.name.local_name == "fontId")
                    .and_then(|a| a.value.parse::<usize>().ok());

                xf.fill_id = attrs
                    .iter()
                    .find(|a| a.name.local_name == "fillId")
                    .and_then(|a| a.value.parse::<usize>().ok());

                xf.border_id = attrs
                    .iter()
                    .find(|a| a.name.local_name == "borderId")
                    .and_then(|a| a.value.parse::<usize>().ok());
            }

            // Parse alignment
            if let Some(alignment_elem) = doc
                .elements_by_name(xf_elem, &XName::new(S::NS, "alignment"))
                .next()
            {
                if let Some(align_data) = doc.get(alignment_elem) {
                    if let Some(attrs) = align_data.attributes() {
                        xf.horizontal = attrs
                            .iter()
                            .find(|a| a.name.local_name == "horizontal")
                            .map(|a| a.value.clone());

                        xf.vertical = attrs
                            .iter()
                            .find(|a| a.name.local_name == "vertical")
                            .map(|a| a.value.clone());

                        xf.wrap_text = attrs
                            .iter()
                            .find(|a| a.name.local_name == "wrapText")
                            .map(|a| a.value == "1" || a.value == "true")
                            .unwrap_or(false);

                        xf.indent = attrs
                            .iter()
                            .find(|a| a.name.local_name == "indent")
                            .and_then(|a| a.value.parse::<i32>().ok());
                    }
                }
            }
        }

        xf
    }

    /// Expand a style index to a full CellFormatSignature.
    fn expand_style(style_index: usize, style_info: &StyleInfo) -> CellFormatSignature {
        if style_index >= style_info.cell_xfs.len() {
            return CellFormatSignature::default();
        }

        let xf = &style_info.cell_xfs[style_index];

        // Get number format
        let number_format_code = xf
            .num_fmt_id
            .and_then(|id| style_info.number_formats.get(&id))
            .cloned()
            .or_else(|| Some("General".to_string()));

        // Get font
        let font = xf
            .font_id
            .and_then(|id| style_info.fonts.get(id))
            .cloned()
            .unwrap_or_default();

        // Get fill
        let fill = xf
            .fill_id
            .and_then(|id| style_info.fills.get(id))
            .cloned()
            .unwrap_or_default();

        // Get border
        let border = xf
            .border_id
            .and_then(|id| style_info.borders.get(id))
            .cloned()
            .unwrap_or_default();

        CellFormatSignature {
            number_format_code,
            bold: font.bold,
            italic: font.italic,
            underline: font.underline,
            strikethrough: font.strikethrough,
            font_name: font.name.clone(),
            font_size: font.size,
            font_color: font.color.clone(),
            fill_pattern: fill.pattern.clone(),
            fill_foreground_color: fill.fg_color.clone(),
            fill_background_color: fill.bg_color.clone(),
            border_left_style: border.left_style.clone(),
            border_left_color: border.left_color.clone(),
            border_right_style: border.right_style.clone(),
            border_right_color: border.right_color.clone(),
            border_top_style: border.top_style.clone(),
            border_top_color: border.top_color.clone(),
            border_bottom_style: border.bottom_style.clone(),
            border_bottom_color: border.bottom_color.clone(),
            horizontal_alignment: xf.horizontal.clone(),
            vertical_alignment: xf.vertical.clone(),
            wrap_text: xf.wrap_text,
            indent: xf.indent,
        }
    }

    /// Compute hash signatures for each row to enable LCS-based alignment.
    fn compute_row_signatures(signature: &mut WorksheetSignature, settings: &SmlComparerSettings) {
        for &row_index in &signature.populated_rows {
            let cells_in_row = signature.get_cells_in_row(row_index);
            if cells_in_row.is_empty() {
                continue;
            }

            // Sample cells for signature (to handle wide sheets efficiently)
            let sampled = if cells_in_row.len() <= settings.row_signature_sample_size as usize {
                cells_in_row
            } else {
                Self::sample_cells(cells_in_row, settings.row_signature_sample_size as usize)
            };

            let row_content: Vec<String> = sampled
                .iter()
                .map(|c| c.resolved_value.as_deref().unwrap_or(""))
                .map(|s| s.to_string())
                .collect();

            let row_content_str = row_content.join("|");
            signature
                .row_signatures
                .insert(row_index, Self::compute_quick_hash(&row_content_str));
        }
    }

    /// Compute hash signatures for each column to enable LCS-based alignment.
    fn compute_column_signatures(
        signature: &mut WorksheetSignature,
        settings: &SmlComparerSettings,
    ) {
        for &col_index in &signature.populated_columns {
            let cells_in_col = signature.get_cells_in_column(col_index);
            if cells_in_col.is_empty() {
                continue;
            }

            // Sample cells for signature
            let sampled = if cells_in_col.len() <= settings.row_signature_sample_size as usize {
                cells_in_col
            } else {
                Self::sample_cells(cells_in_col, settings.row_signature_sample_size as usize)
            };

            let col_content: Vec<String> = sampled
                .iter()
                .map(|c| c.resolved_value.as_deref().unwrap_or(""))
                .map(|s| s.to_string())
                .collect();

            let col_content_str = col_content.join("|");
            signature
                .column_signatures
                .insert(col_index, Self::compute_quick_hash(&col_content_str));
        }
    }

    /// Sample cells evenly from a list for signature computation.
    fn sample_cells(cells: Vec<&CellSignature>, sample_size: usize) -> Vec<&CellSignature> {
        if cells.len() <= sample_size {
            return cells;
        }

        let mut result = Vec::with_capacity(sample_size);
        let step = cells.len() as f64 / sample_size as f64;

        for i in 0..sample_size {
            let index = (i as f64 * step) as usize;
            result.push(cells[index]);
        }

        result
    }

    /// Compute a quick hash for row/column signatures.
    /// Uses a simple hash for performance; SHA256 is overkill for row signatures.
    fn compute_quick_hash(content: &str) -> String {
        let mut hash: i32 = 17;

        for ch in content.chars() {
            hash = hash.wrapping_mul(31).wrapping_add(ch as i32);
        }

        format!("{:08X}", hash as u32)
    }

    /// Extract cell comments from the worksheet.
    fn extract_comments(
        pkg: &OoxmlPackage,
        sheet_path: &str,
        signature: &mut WorksheetSignature,
    ) -> Result<()> {
        // Derive comments part path from sheet path
        // e.g., xl/worksheets/sheet1.xml -> xl/worksheets/_rels/sheet1.xml.rels
        let parts: Vec<&str> = sheet_path.rsplitn(2, '/').collect();
        if parts.len() != 2 {
            return Ok(());
        }

        let dir = parts[1];
        let filename = parts[0];
        let rels_path = format!("{}/_rels/{}.rels", dir, filename);

        // Find comments relationship
        let rels_doc = match pkg.get_xml_part(&rels_path) {
            Ok(doc) => doc,
            Err(_) => return Ok(()), // No rels is valid
        };

        let rels_root = rels_doc.root();
        if rels_root.is_none() {
            return Ok(());
        }

        let rel_name = XName::new(
            "http://schemas.openxmlformats.org/package/2006/relationships",
            "Relationship",
        );
        let comments_type =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

        let mut comments_target: Option<String> = None;
        for rel_elem in rels_doc.elements_by_name(rels_root.unwrap(), &rel_name) {
            if let Some(rel_data) = rels_doc.get(rel_elem) {
                if let Some(attrs) = rel_data.attributes() {
                    if let Some(type_attr) = attrs.iter().find(|a| a.name.local_name == "Type") {
                        if type_attr.value == comments_type {
                            if let Some(target_attr) =
                                attrs.iter().find(|a| a.name.local_name == "Target")
                            {
                                comments_target = Some(target_attr.value.clone());
                                break;
                            }
                        }
                    }
                }
            }
        }

        if comments_target.is_none() {
            return Ok(());
        }

        let comments_path = format!("{}/{}", dir, comments_target.unwrap());
        let comments_doc = match pkg.get_xml_part(&comments_path) {
            Ok(doc) => doc,
            Err(_) => return Ok(()),
        };

        let comments_root = comments_doc.root();
        if comments_root.is_none() {
            return Ok(());
        }
        let comments_root = comments_root.unwrap();

        // Get authors
        let mut authors: HashMap<usize, String> = HashMap::new();
        if let Some(authors_elem) = comments_doc
            .elements_by_name(comments_root, &S::authors())
            .next()
        {
            let mut index = 0;
            for author_elem in comments_doc.elements_by_name(authors_elem, &S::author()) {
                if let Some(author_text) = get_text_content(&comments_doc, author_elem) {
                    authors.insert(index, author_text);
                }
                index += 1;
            }
        }

        // Get comments
        if let Some(comment_list_elem) = comments_doc
            .elements_by_name(comments_root, &S::commentList())
            .next()
        {
            for comment_elem in comments_doc.elements_by_name(comment_list_elem, &S::comment()) {
                if let Some(comment_data) = comments_doc.get(comment_elem) {
                    if let Some(attrs) = comment_data.attributes() {
                        let cell_ref = attrs
                            .iter()
                            .find(|a| a.name.local_name == "ref")
                            .map(|a| a.value.clone());

                        if cell_ref.is_none() || cell_ref.as_ref().unwrap().is_empty() {
                            continue;
                        }
                        let cell_ref = cell_ref.unwrap();

                        let author_id = attrs
                            .iter()
                            .find(|a| a.name.local_name == "authorId")
                            .and_then(|a| a.value.parse::<usize>().ok())
                            .unwrap_or(0);

                        let author = authors
                            .get(&author_id)
                            .cloned()
                            .unwrap_or_else(|| "Unknown".to_string());

                        // Extract comment text
                        let mut text = String::new();
                        if let Some(text_elem) = comments_doc
                            .elements_by_name(comment_elem, &S::text())
                            .next()
                        {
                            for descendant in comments_doc.descendants(text_elem) {
                                if let Some(d_data) = comments_doc.get(descendant) {
                                    if let Some(name) = d_data.name() {
                                        if name == &S::t() {
                                            if let Some(content) =
                                                get_text_content(&comments_doc, descendant)
                                            {
                                                text.push_str(&content);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        signature.comments.insert(
                            cell_ref.clone(),
                            CommentSignature {
                                cell_address: cell_ref,
                                author,
                                text,
                            },
                        );
                    }
                }
            }
        }

        Ok(())
    }

    /// Extract data validation rules from the worksheet.
    fn extract_data_validations(
        ws_doc: &XmlDocument,
        ws_root: NodeId,
        signature: &mut WorksheetSignature,
    ) -> Result<()> {
        let data_validations_elem = ws_doc
            .elements_by_name(ws_root, &S::dataValidations())
            .next();
        if data_validations_elem.is_none() {
            return Ok(());
        }

        for dv_elem in ws_doc.elements_by_name(data_validations_elem.unwrap(), &S::dataValidation())
        {
            if let Some(dv_data) = ws_doc.get(dv_elem) {
                if let Some(attrs) = dv_data.attributes() {
                    let sqref = attrs
                        .iter()
                        .find(|a| a.name.local_name == "sqref")
                        .map(|a| a.value.clone());

                    if sqref.is_none() || sqref.as_ref().unwrap().is_empty() {
                        continue;
                    }
                    let sqref = sqref.unwrap();

                    let validation_type = attrs
                        .iter()
                        .find(|a| a.name.local_name == "type")
                        .map(|a| a.value.clone())
                        .unwrap_or_else(|| "none".to_string());

                    let operator = attrs
                        .iter()
                        .find(|a| a.name.local_name == "operator")
                        .map(|a| a.value.clone());

                    let formula1 = ws_doc
                        .elements_by_name(dv_elem, &S::formula1())
                        .next()
                        .and_then(|f| get_text_content(ws_doc, f));

                    let formula2 = ws_doc
                        .elements_by_name(dv_elem, &S::formula2())
                        .next()
                        .and_then(|f| get_text_content(ws_doc, f));

                    let allow_blank = attrs
                        .iter()
                        .find(|a| a.name.local_name == "allowBlank")
                        .map(|a| a.value == "1" || a.value == "true")
                        .unwrap_or(false);

                    // Note: showDropDown attribute means HIDE dropdown (inverted logic)
                    let show_drop_down = attrs
                        .iter()
                        .find(|a| a.name.local_name == "showDropDown")
                        .map(|a| a.value != "1" && a.value != "true")
                        .unwrap_or(true);

                    let show_input_message = attrs
                        .iter()
                        .find(|a| a.name.local_name == "showInputMessage")
                        .map(|a| a.value == "1" || a.value == "true")
                        .unwrap_or(false);

                    let show_error_message = attrs
                        .iter()
                        .find(|a| a.name.local_name == "showErrorMessage")
                        .map(|a| a.value == "1" || a.value == "true")
                        .unwrap_or(false);

                    let error_title = attrs
                        .iter()
                        .find(|a| a.name.local_name == "errorTitle")
                        .map(|a| a.value.clone());

                    let error = attrs
                        .iter()
                        .find(|a| a.name.local_name == "error")
                        .map(|a| a.value.clone());

                    let prompt_title = attrs
                        .iter()
                        .find(|a| a.name.local_name == "promptTitle")
                        .map(|a| a.value.clone());

                    let prompt = attrs
                        .iter()
                        .find(|a| a.name.local_name == "prompt")
                        .map(|a| a.value.clone());

                    // Parse sqref into individual cell references
                    for range in sqref.split_whitespace() {
                        let key = if range.contains(':') {
                            range.split(':').next().unwrap().to_string()
                        } else {
                            range.to_string()
                        };

                        signature.data_validations.insert(
                            key,
                            DataValidationSignature {
                                cell_range: range.to_string(),
                                validation_type: validation_type.clone(),
                                operator: operator.clone(),
                                formula1: formula1.clone(),
                                formula2: formula2.clone(),
                                allow_blank,
                                show_drop_down,
                                show_input_message,
                                show_error_message,
                                error_title: error_title.clone(),
                                error: error.clone(),
                                prompt_title: prompt_title.clone(),
                                prompt: prompt.clone(),
                            },
                        );
                    }
                }
            }
        }

        Ok(())
    }

    /// Extract merged cell regions from the worksheet.
    fn extract_merged_cells(
        ws_doc: &XmlDocument,
        ws_root: NodeId,
        signature: &mut WorksheetSignature,
    ) -> Result<()> {
        let merge_cells_elem = ws_doc.elements_by_name(ws_root, &S::mergeCells()).next();
        if merge_cells_elem.is_none() {
            return Ok(());
        }

        for mc_elem in ws_doc.elements_by_name(merge_cells_elem.unwrap(), &S::mergeCell()) {
            if let Some(mc_data) = ws_doc.get(mc_elem) {
                if let Some(attrs) = mc_data.attributes() {
                    if let Some(range_attr) = attrs.iter().find(|a| a.name.local_name == "ref") {
                        let range = range_attr.value.clone();
                        if !range.is_empty() {
                            signature.merged_cell_ranges.insert(range);
                        }
                    }
                }
            }
        }

        Ok(())
    }

    /// Extract hyperlinks from the worksheet.
    fn extract_hyperlinks(
        pkg: &OoxmlPackage,
        sheet_path: &str,
        ws_doc: &XmlDocument,
        ws_root: NodeId,
        signature: &mut WorksheetSignature,
    ) -> Result<()> {
        let hyperlinks_elem = ws_doc.elements_by_name(ws_root, &S::hyperlinks()).next();
        if hyperlinks_elem.is_none() {
            return Ok(());
        }

        // Load relationships for external hyperlinks
        let parts: Vec<&str> = sheet_path.rsplitn(2, '/').collect();
        let mut hyperlink_rels: HashMap<String, String> = HashMap::new();

        if parts.len() == 2 {
            let dir = parts[1];
            let filename = parts[0];
            let rels_path = format!("{}/_rels/{}.rels", dir, filename);

            if let Ok(rels_doc) = pkg.get_xml_part(&rels_path) {
                if let Some(rels_root) = rels_doc.root() {
                    let rel_name = XName::new(
                        "http://schemas.openxmlformats.org/package/2006/relationships",
                        "Relationship",
                    );
                    let hyperlink_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

                    for rel_elem in rels_doc.elements_by_name(rels_root, &rel_name) {
                        if let Some(rel_data) = rels_doc.get(rel_elem) {
                            if let Some(attrs) = rel_data.attributes() {
                                if let Some(type_attr) =
                                    attrs.iter().find(|a| a.name.local_name == "Type")
                                {
                                    if type_attr.value == hyperlink_type {
                                        let id = attrs
                                            .iter()
                                            .find(|a| a.name.local_name == "Id")
                                            .map(|a| a.value.clone());

                                        let target = attrs
                                            .iter()
                                            .find(|a| a.name.local_name == "Target")
                                            .map(|a| a.value.clone());

                                        if let (Some(id), Some(target)) = (id, target) {
                                            hyperlink_rels.insert(id, target);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        for hl_elem in ws_doc.elements_by_name(hyperlinks_elem.unwrap(), &S::hyperlink()) {
            if let Some(hl_data) = ws_doc.get(hl_elem) {
                if let Some(attrs) = hl_data.attributes() {
                    let cell_ref = attrs
                        .iter()
                        .find(|a| a.name.local_name == "ref")
                        .map(|a| a.value.clone());

                    if cell_ref.is_none() || cell_ref.as_ref().unwrap().is_empty() {
                        continue;
                    }
                    let cell_ref = cell_ref.unwrap();

                    let r_id = attrs
                        .iter()
                        .find(|a| a.name == R::id())
                        .map(|a| a.value.clone());

                    let mut target = String::new();

                    // Get external target from relationship
                    if let Some(ref id) = r_id {
                        if let Some(t) = hyperlink_rels.get(id) {
                            target = t.clone();
                        }
                    }

                    // Could also be an internal reference (location attribute)
                    if target.is_empty() {
                        if let Some(location_attr) =
                            attrs.iter().find(|a| a.name.local_name == "location")
                        {
                            target = location_attr.value.clone();
                        }
                    }

                    let display = attrs
                        .iter()
                        .find(|a| a.name.local_name == "display")
                        .map(|a| a.value.clone());

                    let tooltip = attrs
                        .iter()
                        .find(|a| a.name.local_name == "tooltip")
                        .map(|a| a.value.clone());

                    signature.hyperlinks.insert(
                        cell_ref.clone(),
                        HyperlinkSignature {
                            cell_address: cell_ref,
                            target,
                            display,
                            tooltip,
                        },
                    );
                }
            }
        }

        Ok(())
    }
}

/// Internal structure to hold parsed style information.
#[derive(Debug, Clone, Default)]
struct StyleInfo {
    number_formats: HashMap<i32, String>,
    fonts: Vec<FontInfo>,
    fills: Vec<FillInfo>,
    borders: Vec<BorderInfo>,
    cell_xfs: Vec<XfInfo>,
}

#[derive(Debug, Clone, Default)]
struct FontInfo {
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    name: Option<String>,
    size: Option<f64>,
    color: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct FillInfo {
    pattern: Option<String>,
    fg_color: Option<String>,
    bg_color: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct BorderInfo {
    left_style: Option<String>,
    left_color: Option<String>,
    right_style: Option<String>,
    right_color: Option<String>,
    top_style: Option<String>,
    top_color: Option<String>,
    bottom_style: Option<String>,
    bottom_color: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct XfInfo {
    num_fmt_id: Option<i32>,
    font_id: Option<usize>,
    fill_id: Option<usize>,
    border_id: Option<usize>,
    horizontal: Option<String>,
    vertical: Option<String>,
    wrap_text: bool,
    indent: Option<i32>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_reference() {
        assert_eq!(
            SmlCanonicalizer::parse_cell_reference("A1").unwrap(),
            (1, 1)
        );
        assert_eq!(
            SmlCanonicalizer::parse_cell_reference("B2").unwrap(),
            (2, 2)
        );
        assert_eq!(
            SmlCanonicalizer::parse_cell_reference("Z26").unwrap(),
            (26, 26)
        );
        assert_eq!(
            SmlCanonicalizer::parse_cell_reference("AA1").unwrap(),
            (27, 1)
        );
        assert_eq!(
            SmlCanonicalizer::parse_cell_reference("AB10").unwrap(),
            (28, 10)
        );
    }

    #[test]
    fn test_normalize_numeric() {
        assert_eq!(SmlCanonicalizer::normalize_numeric("123"), "123");
        assert_eq!(SmlCanonicalizer::normalize_numeric("123.45"), "123.45");
        assert_eq!(SmlCanonicalizer::normalize_numeric(""), "");
    }

    #[test]
    fn test_compute_quick_hash() {
        let hash1 = SmlCanonicalizer::compute_quick_hash("test");
        let hash2 = SmlCanonicalizer::compute_quick_hash("test");
        let hash3 = SmlCanonicalizer::compute_quick_hash("different");

        assert_eq!(hash1, hash2);
        assert_ne!(hash1, hash3);
    }
}
