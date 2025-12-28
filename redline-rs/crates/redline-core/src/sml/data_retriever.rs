use crate::error::{RedlineError, Result};
use crate::sml::SmlDocument;
use crate::xml::{XmlDocument, XmlNodeData, XName, XAttribute, S, R};

pub struct SmlDataRetriever;

impl SmlDataRetriever {
    pub fn retrieve_sheet(sml_doc: &SmlDocument, sheet_name: &str) -> Result<XmlDocument> {
        let pkg = sml_doc.package();
        
        let workbook_path = "xl/workbook.xml";
        let workbook_doc = pkg.get_xml_part(workbook_path)?;
        let workbook_root = workbook_doc.root().ok_or_else(|| RedlineError::InvalidPackage { message: "Workbook has no root".to_string() })?;
        
        // Find sheets element
        let sheets_elem = workbook_doc.elements_by_name(workbook_root, &XName::new(S::NS, "sheets")).next()
            .ok_or_else(|| RedlineError::InvalidPackage { message: "Missing sheets element in workbook".to_string() })?;
            
        // Find specific sheet
        let sheet_elem = workbook_doc.elements_by_name(sheets_elem, &XName::new(S::NS, "sheet"))
            .find(|&s| {
                workbook_doc.get(s)
                    .and_then(|d| d.attributes())
                    .map_or(false, |attrs| attrs.iter().any(|a| a.name.local_name == "name" && a.value == sheet_name))
            })
            .ok_or_else(|| RedlineError::InvalidPackage { message: format!("Invalid sheet name: {}", sheet_name) })?;
            
        // Get r:id
        let r_id = workbook_doc.get(sheet_elem)
            .and_then(|d| d.attributes())
            .and_then(|attrs| attrs.iter().find(|a| a.name == R::id()))
            .map(|a| a.value.clone())
            .ok_or_else(|| RedlineError::InvalidPackage { message: "Sheet has no r:id".to_string() })?;

        // Resolve relationship
        let workbook_rels_path = "xl/_rels/workbook.xml.rels";
        let workbook_rels_doc = pkg.get_xml_part(workbook_rels_path)?;
        let rels_root = workbook_rels_doc.root().ok_or_else(|| RedlineError::InvalidPackage { message: "Workbook rels has no root".to_string() })?;
        
        let rel_elem = workbook_rels_doc.elements_by_name(rels_root, &XName::new("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship"))
            .find(|&r| {
                workbook_rels_doc.get(r)
                    .and_then(|d| d.attributes())
                    .map_or(false, |attrs| attrs.iter().any(|a| a.name.local_name == "Id" && a.value == r_id))
            })
            .ok_or_else(|| RedlineError::InvalidPackage { message: format!("Relationship not found for id: {}", r_id) })?;
            
        let target = workbook_rels_doc.get(rel_elem)
            .and_then(|d| d.attributes())
            .and_then(|attrs| attrs.iter().find(|a| a.name.local_name == "Target"))
            .map(|a| a.value.clone())
            .ok_or_else(|| RedlineError::InvalidPackage { message: "Relationship has no Target".to_string() })?;
            
        let sheet_path = if target.starts_with('/') {
            target[1..].to_string()
        } else {
            format!("xl/{}", target)
        };
        
        let range = "A1:XFD1048576";
        let (left_col, top_row, right_col, bottom_row) = xlsx_tables::parse_range(range)?;
        
        Self::retrieve_range_impl(sml_doc, &sheet_path, left_col, top_row, right_col, bottom_row)
    }

    pub fn retrieve_range(sml_doc: &SmlDocument, sheet_name: &str, range: &str) -> Result<XmlDocument> {
        let pkg = sml_doc.package();
        let workbook_path = "xl/workbook.xml";
        let workbook_doc = pkg.get_xml_part(workbook_path)?;
        let workbook_root = workbook_doc.root().ok_or_else(|| RedlineError::InvalidPackage { message: "Workbook has no root".to_string() })?;
        
        let sheets_elem = workbook_doc.elements_by_name(workbook_root, &XName::new(S::NS, "sheets")).next()
            .ok_or_else(|| RedlineError::InvalidPackage { message: "Missing sheets element in workbook".to_string() })?;
            
        let sheet_elem = workbook_doc.elements_by_name(sheets_elem, &XName::new(S::NS, "sheet"))
            .find(|&s| {
                workbook_doc.get(s)
                    .and_then(|d| d.attributes())
                    .map_or(false, |attrs| attrs.iter().any(|a| a.name.local_name == "name" && a.value == sheet_name))
            })
            .ok_or_else(|| RedlineError::InvalidPackage { message: format!("Invalid sheet name: {}", sheet_name) })?;
            
        let r_id = workbook_doc.get(sheet_elem)
            .and_then(|d| d.attributes())
            .and_then(|attrs| attrs.iter().find(|a| a.name == R::id()))
            .map(|a| a.value.clone())
            .ok_or_else(|| RedlineError::InvalidPackage { message: "Sheet has no r:id".to_string() })?;

        let workbook_rels_path = "xl/_rels/workbook.xml.rels";
        let workbook_rels_doc = pkg.get_xml_part(workbook_rels_path)?;
        let rels_root = workbook_rels_doc.root().ok_or_else(|| RedlineError::InvalidPackage { message: "Workbook rels has no root".to_string() })?;
        
        let rel_elem = workbook_rels_doc.elements_by_name(rels_root, &XName::new("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship"))
            .find(|&r| {
                workbook_rels_doc.get(r)
                    .and_then(|d| d.attributes())
                    .map_or(false, |attrs| attrs.iter().any(|a| a.name.local_name == "Id" && a.value == r_id))
            })
            .ok_or_else(|| RedlineError::InvalidPackage { message: format!("Relationship not found for id: {}", r_id) })?;
            
        let target = workbook_rels_doc.get(rel_elem)
            .and_then(|d| d.attributes())
            .and_then(|attrs| attrs.iter().find(|a| a.name.local_name == "Target"))
            .map(|a| a.value.clone())
            .ok_or_else(|| RedlineError::InvalidPackage { message: "Relationship has no Target".to_string() })?;
            
        let sheet_path = if target.starts_with('/') {
            target[1..].to_string()
        } else {
            format!("xl/{}", target)
        };
        
        let (left_col, top_row, right_col, bottom_row) = xlsx_tables::parse_range(range)?;
        
        Self::retrieve_range_impl(sml_doc, &sheet_path, left_col, top_row, right_col, bottom_row)
    }

    fn retrieve_range_impl(
        sml_doc: &SmlDocument,
        sheet_path: &str,
        left_col: i32,
        top_row: i32,
        right_col: i32,
        bottom_row: i32,
    ) -> Result<XmlDocument> {
        let pkg = sml_doc.package();
        let sh_xdoc = pkg.get_xml_part(sheet_path)?;
        let sh_root = sh_xdoc.root().ok_or_else(|| RedlineError::InvalidPackage { message: "Sheet has no root".to_string() })?;

        // Shared Strings
        let sst_path = "xl/sharedStrings.xml";
        let sst_xdoc = pkg.get_xml_part(sst_path).ok();

        // New Document for output
        let mut out_doc = XmlDocument::new();
        let out_root_id = out_doc.add_root(XmlNodeData::element(XName::new("", "Data")));

        let sheet_data_name = S::sheetData();
        let sheet_data_elem = sh_xdoc.elements_by_name(sh_root, &sheet_data_name).next();
        
        if let Some(sheet_data_id) = sheet_data_elem {
            let out_sheet_data_id = out_doc.add_child(out_root_id, XmlNodeData::element(S::sheetData()));

            let row_name = S::row();
            for row_id in sh_xdoc.elements_by_name(sheet_data_id, &row_name) {
                let r_attr_val = sh_xdoc.get(row_id)
                    .and_then(|d| d.attributes())
                    .and_then(|attrs| attrs.iter().find(|a| a.name.local_name == "r"))
                    .map(|a| a.value.as_str());
                    
                if r_attr_val.is_none() { continue; }
                let row_nbr: i32 = r_attr_val.unwrap().parse().unwrap_or(0);
                
                if row_nbr < top_row || row_nbr > bottom_row {
                    continue;
                }

                let mut out_row_data = XmlNodeData::element(XName::new("", "Row"));
                if let Some(attrs) = out_row_data.attributes_mut() {
                    attrs.push(XAttribute::new(XName::new("", "RowNumber"), &row_nbr.to_string()));
                }
                
                let mut out_cells = Vec::new();

                let c_name = S::c();
                for cell_id in sh_xdoc.elements_by_name(row_id, &c_name) {
                    let cell_addr_val = sh_xdoc.get(cell_id)
                        .and_then(|d| d.attributes())
                        .and_then(|attrs| attrs.iter().find(|a| a.name.local_name == "r"))
                        .map(|a| a.value.clone());
                        
                    if cell_addr_val.is_none() { continue; }
                    let cell_address = cell_addr_val.unwrap();
                    
                    let (col_addr, _) = xlsx_tables::split_address(&cell_address)?;
                    let col_index = xlsx_tables::column_address_to_index(&col_addr)?;

                    if col_index < left_col || col_index > right_col {
                        continue;
                    }

                    // Process cell
                    let cell_type = sh_xdoc.get(cell_id)
                        .and_then(|d| d.attributes())
                        .and_then(|attrs| attrs.iter().find(|a| a.name.local_name == "t"))
                        .map(|a| a.value.as_str())
                        .unwrap_or("");

                    let mut shared_string = None;

                    if cell_type == "s" {
                         let v_name = S::v();
                         let v_elem = sh_xdoc.elements_by_name(cell_id, &v_name).next();
                         if let Some(v_id) = v_elem {
                             if let Some(val) = sh_xdoc.get(v_id).and_then(|d| d.text_content()) {
                                 if let Ok(idx) = val.parse::<usize>() {
                                     if let Some(ref sst) = sst_xdoc {
                                         if let Some(sst_root) = sst.root() {
                                             let si_name = XName::new(S::NS, "si");
                                             // Find si by index
                                             let si_node = sst.elements_by_name(sst_root, &si_name).nth(idx);
                                             
                                             if let Some(si) = si_node {
                                                  let mut s = String::new();
                                                  // Find t elements manually
                                                  let t_name = S::t();
                                                  for descendant in sst.descendants(si) {
                                                      if let Some(d_data) = sst.get(descendant) {
                                                          // Check if it's <t> element
                                                          if let Some(name) = d_data.name() {
                                                              if name == &t_name {
                                                                  // Get text content of this element (its children text nodes)
                                                                  // XmlDocument doesn't have element_text_content helper
                                                                  for child_id in sst.children(descendant) {
                                                                      if let Some(c_data) = sst.get(child_id) {
                                                                          if let Some(text) = c_data.text_content() {
                                                                              s.push_str(text);
                                                                          }
                                                                      }
                                                                  }
                                                              }
                                                          }
                                                      }
                                                  }
                                                  if !s.is_empty() {
                                                      shared_string = Some(s);
                                                  }
                                             }
                                         }
                                     }
                                 }
                             }
                         }
                    } else if cell_type == "inlineStr" {
                         let is_name = S::is();
                         let is_elem = sh_xdoc.elements_by_name(cell_id, &is_name).next();
                         if let Some(is_id) = is_elem {
                             let t_name = S::t();
                             let t_elem = sh_xdoc.elements_by_name(is_id, &t_name).next();
                             if let Some(t_id) = t_elem {
                                 let mut s = String::new();
                                 for child in sh_xdoc.children(t_id) {
                                     if let Some(c_data) = sh_xdoc.get(child) {
                                         if let Some(txt) = c_data.text_content() {
                                             s.push_str(txt);
                                         }
                                     }
                                 }
                                 shared_string = Some(s);
                             }
                         }
                    }

                    let v_name = S::v();
                    let v_val = sh_xdoc.elements_by_name(cell_id, &v_name).next()
                        .and_then(|v| {
                            sh_xdoc.children(v).find_map(|c| sh_xdoc.get(c).and_then(|d| d.text_content()).map(|s| s.to_string()))
                        })
                        .unwrap_or_default();

                    let value = shared_string.clone().unwrap_or(v_val);
                    let display_value = value.clone(); // Placeholder


                    // Construct Cell Node
                    let mut cell_data = XmlNodeData::element(XName::new("", "Cell"));
                    if let Some(attrs) = cell_data.attributes_mut() {
                        attrs.push(XAttribute::new(XName::new("", "Ref"), &cell_address));
                        attrs.push(XAttribute::new(XName::new("", "ColumnId"), &col_addr));
                        attrs.push(XAttribute::new(XName::new("", "ColumnNumber"), &col_index.to_string()));
                        
                        if !cell_type.is_empty() {
                             let type_val = if cell_type == "inlineStr" { "s" } else { cell_type };
                             attrs.push(XAttribute::new(XName::new("", "Type"), type_val));
                        }
                        
                        // Formula
                        if let Some(f) = sh_xdoc.get(cell_id).and_then(|d| d.attributes()).and_then(|a| a.iter().find(|x| x.name.local_name == "f")) {
                            attrs.push(XAttribute::new(XName::new("", "Formula"), &f.value));
                        }
                         // Style
                        if let Some(s) = sh_xdoc.get(cell_id).and_then(|d| d.attributes()).and_then(|a| a.iter().find(|x| x.name.local_name == "s")) {
                            attrs.push(XAttribute::new(XName::new("", "Style"), &s.value));
                        }
                    }
                    
                    // Children: Value, DisplayValue
                    out_cells.push((cell_data, value, display_value));
                }

                if !out_cells.is_empty() {
                    let out_row_id = out_doc.add_child(out_sheet_data_id, out_row_data);
                    for (cell_data, val, disp_val) in out_cells {
                        let cell_id = out_doc.add_child(out_row_id, cell_data);
                        
                        let val_id = out_doc.add_child(cell_id, XmlNodeData::element(XName::new("", "Value")));
                        out_doc.add_child(val_id, XmlNodeData::text(&val));
                        
                        let disp_val_id = out_doc.add_child(cell_id, XmlNodeData::element(XName::new("", "DisplayValue")));
                        out_doc.add_child(disp_val_id, XmlNodeData::text(&disp_val));
                    }
                }
            }
        }

        Ok(out_doc)
    }
}

mod xlsx_tables {
    use crate::error::{RedlineError, Result};

    pub fn parse_range(range: &str) -> Result<(i32, i32, i32, i32)> {
        let parts: Vec<&str> = range.split(':').collect();
        if parts.len() != 2 {
            return Err(RedlineError::InvalidPackage { message: "Invalid range format".to_string() });
        }
        
        let start = parts[0];
        let end = parts[1];
        
        let (start_col_str, start_row_str) = split_address(start)?;
        let left_col = column_address_to_index(&start_col_str)?;
        let top_row = start_row_str.parse::<i32>().map_err(|_| RedlineError::InvalidPackage { message: "Invalid row number".to_string() })?;
        
        let (end_col_str, end_row_str) = split_address(end)?;
        let right_col = column_address_to_index(&end_col_str)?;
        let bottom_row = end_row_str.parse::<i32>().map_err(|_| RedlineError::InvalidPackage { message: "Invalid row number".to_string() })?;
        
        Ok((left_col, top_row, right_col, bottom_row))
    }

    pub fn split_address(address: &str) -> Result<(String, String)> {
        let mut i = 0;
        for c in address.chars() {
            if c.is_digit(10) {
                break;
            }
            i += 1;
        }
        
        if i == 0 || i == address.len() {
             return Err(RedlineError::InvalidPackage { message: "Invalid cell address".to_string() });
        }
        
        Ok((address[0..i].to_string(), address[i..].to_string()))
    }

    pub fn column_address_to_index(address: &str) -> Result<i32> {
        let mut index = 0;
        let mut mul = 1;
        
        for c in address.chars().rev() {
            let val = c as i32 - 'A' as i32 + 1;
            index += val * mul;
            mul *= 26;
        }
        
        Ok(index - 1)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_column_address_to_index() {

        assert_eq!(xlsx_tables::column_address_to_index("A").unwrap(), 0);
        assert_eq!(xlsx_tables::column_address_to_index("B").unwrap(), 1);
        assert_eq!(xlsx_tables::column_address_to_index("Z").unwrap(), 25);
        assert_eq!(xlsx_tables::column_address_to_index("AA").unwrap(), 26);
        assert_eq!(xlsx_tables::column_address_to_index("AB").unwrap(), 27);
    }

    #[test]
    fn test_split_address() {
        assert_eq!(xlsx_tables::split_address("A1").unwrap(), ("A".to_string(), "1".to_string()));
        assert_eq!(xlsx_tables::split_address("AA10").unwrap(), ("AA".to_string(), "10".to_string()));
    }
}
