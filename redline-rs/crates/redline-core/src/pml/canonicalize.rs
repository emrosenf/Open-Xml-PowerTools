// Canonicalizer: Extracts semantic signatures from PowerPoint presentations
// Ported from C# PmlComparer.cs lines 583-1121

use super::document::PmlDocument;
use super::settings::PmlComparerSettings;
use super::slide_matching::{
    ParagraphSignature, PlaceholderInfo, PmlShapeType, PresentationSignature,
    RunPropertiesSignature, RunSignature, ShapeSignature, SlideSignature, TextBodySignature,
    TransformSignature,
};
use crate::error::{RedlineError, Result};
use crate::hash::sha256::{sha256_hash_bytes, sha256_hash_string};
use crate::xml::xname::XName;
use crate::xml::{XmlDocument, A, P, R};

/// Canonicalizer: Extracts semantic signatures from presentations.
pub struct PmlCanonicalizer;

impl PmlCanonicalizer {
    /// Canonicalize a PmlDocument into a PresentationSignature.
    ///
    /// This extracts all semantic information needed for comparison:
    /// - Slide dimensions
    /// - Slide layouts and backgrounds
    /// - Shape hierarchies with positions, types, and content
    /// - Text content with formatting
    /// - Images, tables, and charts
    pub fn canonicalize(
        doc: &PmlDocument,
        settings: &PmlComparerSettings,
    ) -> Result<PresentationSignature> {
        let package = doc.package();

        // Get presentation.xml
        let pres_path = "ppt/presentation.xml";
        let pres_doc = package.get_xml_part(pres_path)?;
        let pres_root = pres_doc.root().ok_or_else(|| RedlineError::XmlParse {
            message: "Missing presentation root".to_string(),
            location: pres_path.to_string(),
        })?;

        let mut signature = PresentationSignature {
            slide_cx: 0,
            slide_cy: 0,
            slides: Vec::new(),
            theme_hash: None,
        };

        // Get slide size
        if let Some(sld_sz) = pres_doc.find_child(pres_root, &P::sld_sz()) {
            signature.slide_cx = pres_doc.get_attribute_i64(sld_sz, "cx").unwrap_or(0);
            signature.slide_cy = pres_doc.get_attribute_i64(sld_sz, "cy").unwrap_or(0);
        }

        // Get slide list
        let sld_id_lst = pres_doc
            .find_child(pres_root, &P::sld_id_lst())
            .ok_or_else(|| RedlineError::XmlParse {
                message: "Missing sldIdLst".to_string(),
                location: pres_path.to_string(),
            })?;

        let mut slide_index = 1;
        for sld_id in pres_doc.children(sld_id_lst) {
            if pres_doc.name(sld_id) != Some(&P::sld_id()) {
                continue;
            }

            let r_id = match pres_doc.get_attribute_string(sld_id, &R::id()) {
                Some(id) if !id.is_empty() => id,
                _ => continue,
            };

            // Get slide part path from relationship
            let slide_path = Self::resolve_relationship(package, pres_path, &r_id)?;

            // Canonicalize slide
            match Self::canonicalize_slide(package, &slide_path, slide_index, &r_id, settings) {
                Ok(slide_sig) => signature.slides.push(slide_sig),
                Err(_) => {
                    // Skip invalid slide references (as C# does with empty catch)
                }
            }

            slide_index += 1;
        }

        Ok(signature)
    }

    fn canonicalize_slide(
        package: &crate::package::OoxmlPackage,
        slide_path: &str,
        index: usize,
        r_id: &str,
        settings: &PmlComparerSettings,
    ) -> Result<SlideSignature> {
        let slide_doc = package.get_xml_part(slide_path)?;
        let slide_root = slide_doc.root().ok_or_else(|| RedlineError::XmlParse {
            message: "Missing slide root".to_string(),
            location: slide_path.to_string(),
        })?;

        let mut signature = SlideSignature {
            index,
            relationship_id: r_id.to_string(),
            layout_relationship_id: None,
            layout_hash: None,
            shapes: Vec::new(),
            notes_text: None,
            title_text: None,
            content_hash: String::new(),
            background_hash: None,
        };

        // Get layout reference
        // In PPTX, slide layout relationship is in slide.xml.rels
        let slide_rels_path = Self::get_rels_path(slide_path);
        if let Ok(layout_rid) = Self::get_layout_relationship(package, &slide_rels_path) {
            signature.layout_relationship_id = Some(layout_rid.clone());

            // Get layout part and compute hash
            if let Ok(layout_path) = Self::resolve_relationship(package, slide_path, &layout_rid) {
                if let Ok(layout_doc) = package.get_xml_part(&layout_path) {
                    if let Some(layout_root) = layout_doc.root() {
                        let type_name = XName::local("type");
                        let layout_type = layout_doc
                            .get_attribute_string(layout_root, &type_name)
                            .unwrap_or_else(|| "custom".to_string());
                        signature.layout_hash = Some(sha256_hash_string(&layout_type));
                    }
                }
            }
        }

        // Get common slide data
        let c_sld = slide_doc
            .find_child(slide_root, &P::c_sld())
            .ok_or_else(|| RedlineError::XmlParse {
                message: "Missing cSld".to_string(),
                location: slide_path.to_string(),
            })?;

        // Get background hash
        if let Some(bg) = slide_doc.find_child(c_sld, &P::bg()) {
            let bg_xml = slide_doc.to_xml_string(bg);
            signature.background_hash = Some(sha256_hash_string(&bg_xml));
        }

        // Get shape tree
        let sp_tree =
            slide_doc
                .find_child(c_sld, &P::sp_tree())
                .ok_or_else(|| RedlineError::XmlParse {
                    message: "Missing spTree".to_string(),
                    location: slide_path.to_string(),
                })?;

        let mut z_order = 0;
        for element in slide_doc.children(sp_tree) {
            let name = slide_doc.name(element);
            if name == Some(&P::sp())
                || name == Some(&P::pic())
                || name == Some(&P::graphic_frame())
                || name == Some(&P::grp_sp())
                || name == Some(&P::cxn_sp())
            {
                if let Ok(shape_sig) = Self::canonicalize_shape(
                    &slide_doc, package, slide_path, element, z_order, settings,
                ) {
                    // Extract title text
                    if let Some(ref ph) = shape_sig.placeholder {
                        if ph.type_ == "title" || ph.type_ == "ctrTitle" {
                            if let Some(ref tb) = shape_sig.text_body {
                                signature.title_text = Some(tb.plain_text.clone());
                            }
                        }
                    }

                    signature.shapes.push(shape_sig);
                    z_order += 1;
                }
            }
        }

        // Get notes text
        if settings.add_notes_annotations {
            let notes_rels_path = Self::get_rels_path(slide_path);
            if let Ok(notes_rid) = Self::get_notes_relationship(package, &notes_rels_path) {
                if let Ok(notes_path) = Self::resolve_relationship(package, slide_path, &notes_rid)
                {
                    signature.notes_text = Self::extract_notes_text(package, &notes_path).ok();
                }
            }
        }

        // Compute content hash
        let mut content_builder = String::new();
        content_builder.push_str(signature.title_text.as_deref().unwrap_or(""));
        for shape in &signature.shapes {
            content_builder.push('|');
            content_builder.push_str(&shape.name);
            content_builder.push(':');
            content_builder.push_str(&shape.type_.to_string());
            content_builder.push(':');
            if let Some(ref tb) = shape.text_body {
                content_builder.push_str(&tb.plain_text);
            }
        }
        signature.content_hash = sha256_hash_string(&content_builder);

        Ok(signature)
    }

    fn canonicalize_shape(
        slide_doc: &XmlDocument,
        package: &crate::package::OoxmlPackage,
        slide_path: &str,
        element: indextree::NodeId,
        z_order: i32,
        settings: &PmlComparerSettings,
    ) -> Result<ShapeSignature> {
        let name = slide_doc.name(element);

        let mut signature = ShapeSignature {
            name: String::new(),
            id: 0,
            type_: PmlShapeType::Unknown,
            placeholder: None,
            transform: None,
            z_order,
            geometry_hash: None,
            text_body: None,
            image_hash: None,
            table_hash: None,
            chart_hash: None,
            children: None,
            content_hash: String::new(),
        };

        // Determine shape type
        if name == Some(&P::sp()) {
            signature.type_ = PmlShapeType::AutoShape;
        } else if name == Some(&P::pic()) {
            signature.type_ = PmlShapeType::Picture;
        } else if name == Some(&P::graphic_frame()) {
            // Could be table, chart, or diagram
            if let Some(graphic) = slide_doc.find_child(element, &A::graphic()) {
                if let Some(graphic_data) = slide_doc.find_child(graphic, &A::graphic_data()) {
                    let uri_name = XName::local("uri");
                    let uri = slide_doc
                        .get_attribute_string(graphic_data, &uri_name)
                        .unwrap_or_default();
                    signature.type_ = match uri.as_str() {
                        "http://schemas.openxmlformats.org/drawingml/2006/table" => {
                            PmlShapeType::Table
                        }
                        "http://schemas.openxmlformats.org/drawingml/2006/chart" => {
                            PmlShapeType::Chart
                        }
                        "http://schemas.openxmlformats.org/drawingml/2006/diagram" => {
                            PmlShapeType::SmartArt
                        }
                        _ => PmlShapeType::OleObject,
                    };
                }
            }
        } else if name == Some(&P::grp_sp()) {
            signature.type_ = PmlShapeType::Group;
        } else if name == Some(&P::cxn_sp()) {
            signature.type_ = PmlShapeType::Connector;
        }

        // Get non-visual properties
        let nv_sp_pr = slide_doc
            .find_child(element, &P::nv_sp_pr())
            .or_else(|| slide_doc.find_child(element, &P::nv_pic_pr()))
            .or_else(|| slide_doc.find_child(element, &P::nv_graphic_frame_pr()))
            .or_else(|| slide_doc.find_child(element, &P::nv_grp_sp_pr()))
            .or_else(|| slide_doc.find_child(element, &P::nv_cxn_sp_pr()));

        if let Some(nv_pr_node) = nv_sp_pr {
            if let Some(c_nv_pr) = slide_doc.find_child(nv_pr_node, &P::c_nv_pr()) {
                let name_attr = XName::local("name");
                signature.name = slide_doc
                    .get_attribute_string(c_nv_pr, &name_attr)
                    .unwrap_or_default();
                signature.id = slide_doc.get_attribute_i64(c_nv_pr, "id").unwrap_or(0) as u32;
            }

            // Get placeholder info
            if let Some(nv_pr) = slide_doc.find_child(nv_pr_node, &P::nv_pr()) {
                if let Some(ph) = slide_doc.find_child(nv_pr, &P::ph()) {
                    let type_name = XName::local("type");
                    signature.placeholder = Some(PlaceholderInfo {
                        type_: slide_doc
                            .get_attribute_string(ph, &type_name)
                            .unwrap_or_else(|| "body".to_string()),
                        index: slide_doc.get_attribute_i64(ph, "idx").map(|v| v as u32),
                    });
                }
            }
        }

        // Get transform
        let sp_pr = slide_doc
            .find_child(element, &P::sp_pr())
            .or_else(|| slide_doc.find_child(element, &P::grp_sp_pr()));

        if let Some(sp_pr_node) = sp_pr {
            if let Some(xfrm) = slide_doc.find_child(sp_pr_node, &A::xfrm()) {
                signature.transform = Some(Self::extract_transform(slide_doc, xfrm));
            }

            // Get geometry hash
            if let Some(prst_geom) = slide_doc.find_child(sp_pr_node, &A::prst_geom()) {
                let prst_name = XName::local("prst");
                signature.geometry_hash = slide_doc.get_attribute_string(prst_geom, &prst_name);
            } else if let Some(cust_geom) = slide_doc.find_child(sp_pr_node, &A::cust_geom()) {
                let geom_xml = slide_doc.to_xml_string(cust_geom);
                signature.geometry_hash = Some(sha256_hash_string(&geom_xml));
            }
        }

        // For groups, check grpSpPr for transform
        if name == Some(&P::grp_sp()) {
            if let Some(grp_sp_pr) = slide_doc.find_child(element, &P::grp_sp_pr()) {
                if signature.transform.is_none() {
                    if let Some(xfrm) = slide_doc.find_child(grp_sp_pr, &A::xfrm()) {
                        signature.transform = Some(Self::extract_transform(slide_doc, xfrm));
                    }
                }
            }
        }

        // Get text body
        if let Some(tx_body) = slide_doc.find_child(element, &P::tx_body()) {
            signature.text_body = Some(Self::extract_text_body(slide_doc, tx_body));
            if signature.type_ == PmlShapeType::AutoShape {
                if let Some(ref tb) = signature.text_body {
                    if !tb.plain_text.is_empty() {
                        signature.type_ = PmlShapeType::TextBox;
                    }
                }
            }
        }

        // Get image hash for pictures
        if signature.type_ == PmlShapeType::Picture {
            signature.image_hash =
                Self::extract_image_hash(slide_doc, package, slide_path, element).ok();
        }

        // Get table hash
        if signature.type_ == PmlShapeType::Table {
            signature.table_hash = Self::extract_table_hash(slide_doc, element).ok();
        }

        // Get chart hash
        if signature.type_ == PmlShapeType::Chart {
            signature.chart_hash =
                Self::extract_chart_hash(slide_doc, package, slide_path, element).ok();
        }

        // Handle group children
        if signature.type_ == PmlShapeType::Group {
            let mut children = Vec::new();
            let mut child_z_order = 0;
            for child in slide_doc.children(element) {
                let child_name = slide_doc.name(child);
                if child_name == Some(&P::sp())
                    || child_name == Some(&P::pic())
                    || child_name == Some(&P::graphic_frame())
                    || child_name == Some(&P::grp_sp())
                    || child_name == Some(&P::cxn_sp())
                {
                    if let Ok(child_sig) = Self::canonicalize_shape(
                        slide_doc,
                        package,
                        slide_path,
                        child,
                        child_z_order,
                        settings,
                    ) {
                        children.push(child_sig);
                        child_z_order += 1;
                    }
                }
            }
            signature.children = Some(children);
        }

        // Compute content hash
        let mut content_builder = String::new();
        content_builder.push_str(&signature.type_.to_string());
        content_builder.push('|');
        if let Some(ref tb) = signature.text_body {
            content_builder.push_str(&tb.plain_text);
        }
        content_builder.push('|');
        if let Some(ref img) = signature.image_hash {
            content_builder.push_str(img);
        }
        content_builder.push('|');
        if let Some(ref tbl) = signature.table_hash {
            content_builder.push_str(tbl);
        }
        content_builder.push('|');
        if let Some(ref chart) = signature.chart_hash {
            content_builder.push_str(chart);
        }
        signature.content_hash = sha256_hash_string(&content_builder);

        Ok(signature)
    }

    fn extract_transform(doc: &XmlDocument, xfrm: indextree::NodeId) -> TransformSignature {
        let mut transform = TransformSignature {
            x: 0,
            y: 0,
            cx: 0,
            cy: 0,
            rotation: 0,
            flip_h: false,
            flip_v: false,
        };

        if let Some(off) = doc.find_child(xfrm, &A::off()) {
            transform.x = doc.get_attribute_i64(off, "x").unwrap_or(0);
            transform.y = doc.get_attribute_i64(off, "y").unwrap_or(0);
        }

        if let Some(ext) = doc.find_child(xfrm, &A::ext()) {
            transform.cx = doc.get_attribute_i64(ext, "cx").unwrap_or(0);
            transform.cy = doc.get_attribute_i64(ext, "cy").unwrap_or(0);
        }

        transform.rotation = doc.get_attribute_i32(xfrm, "rot").unwrap_or(0);
        transform.flip_h = doc.get_attribute_bool(xfrm, "flipH").unwrap_or(false);
        transform.flip_v = doc.get_attribute_bool(xfrm, "flipV").unwrap_or(false);

        transform
    }

    fn extract_text_body(doc: &XmlDocument, tx_body: indextree::NodeId) -> TextBodySignature {
        let mut signature = TextBodySignature {
            paragraphs: Vec::new(),
            plain_text: String::new(),
        };

        let mut plain_text_builder = String::new();

        for p in doc.children(tx_body) {
            if doc.name(p) != Some(&A::p()) {
                continue;
            }

            let mut para = ParagraphSignature {
                runs: Vec::new(),
                plain_text: String::new(),
                alignment: None,
                has_bullet: false,
            };

            let mut para_text_builder = String::new();

            // Get paragraph properties
            if let Some(p_pr) = doc.find_child(p, &A::p_pr()) {
                let algn_name = XName::local("algn");
                para.alignment = doc.get_attribute_string(p_pr, &algn_name);
                para.has_bullet = doc.find_child(p_pr, &A::bu_char()).is_some()
                    || doc.find_child(p_pr, &A::bu_auto_num()).is_some();
            }

            // Get runs
            for r in doc.children(p) {
                if doc.name(r) == Some(&A::r()) {
                    let mut run = RunSignature {
                        text: String::new(),
                        properties: None,
                        content_hash: String::new(),
                    };

                    if let Some(t) = doc.find_child(r, &A::t()) {
                        run.text = doc.text(t).unwrap_or_default();
                        para_text_builder.push_str(&run.text);
                    }

                    // Get run properties
                    if let Some(r_pr) = doc.find_child(r, &A::r_pr()) {
                        run.properties = Some(Self::extract_run_properties(doc, r_pr));
                    }

                    run.content_hash = sha256_hash_string(&run.text);
                    para.runs.push(run);
                } else if doc.name(r) == Some(&A::fld()) {
                    // Handle field codes
                    let text = if let Some(t) = doc.find_child(r, &A::t()) {
                        doc.text(t).unwrap_or_default()
                    } else {
                        String::new()
                    };
                    para_text_builder.push_str(&text);

                    let run = RunSignature {
                        text,
                        properties: None,
                        content_hash: String::new(),
                    };
                    para.runs.push(run);
                }
            }

            para.plain_text = para_text_builder.clone();
            if !plain_text_builder.is_empty() {
                plain_text_builder.push('\n');
            }
            plain_text_builder.push_str(&para.plain_text);
            signature.paragraphs.push(para);
        }

        signature.plain_text = plain_text_builder;
        signature
    }

    fn extract_run_properties(
        doc: &XmlDocument,
        r_pr: indextree::NodeId,
    ) -> RunPropertiesSignature {
        let mut props = RunPropertiesSignature {
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            font_name: None,
            font_size: None,
            font_color: None,
        };

        props.bold = doc.get_attribute_bool(r_pr, "b").unwrap_or(false);
        props.italic = doc.get_attribute_bool(r_pr, "i").unwrap_or(false);

        let u_name = XName::local("u");
        if let Some(u_val) = doc.get_attribute_string(r_pr, &u_name) {
            props.underline = !u_val.is_empty() && u_val != "none";
        }

        let strike_name = XName::local("strike");
        if let Some(strike_val) = doc.get_attribute_string(r_pr, &strike_name) {
            props.strikethrough = !strike_val.is_empty() && strike_val != "noStrike";
        }

        props.font_size = doc.get_attribute_i32(r_pr, "sz");

        // Get font name
        if let Some(latin) = doc.find_child(r_pr, &A::latin()) {
            let typeface_name = XName::local("typeface");
            props.font_name = doc.get_attribute_string(latin, &typeface_name);
        }

        // Get font color
        if let Some(solid_fill) = doc.find_child(r_pr, &A::solid_fill()) {
            if let Some(srgb_clr) = doc.find_child(solid_fill, &A::srgb_clr()) {
                let val_name = XName::local("val");
                props.font_color = doc.get_attribute_string(srgb_clr, &val_name);
            }
        }

        props
    }

    fn extract_image_hash(
        doc: &XmlDocument,
        package: &crate::package::OoxmlPackage,
        slide_path: &str,
        element: indextree::NodeId,
    ) -> Result<String> {
        let blip_fill =
            doc.find_child(element, &P::blip_fill())
                .ok_or_else(|| RedlineError::XmlParse {
                    message: "Missing blipFill".to_string(),
                    location: slide_path.to_string(),
                })?;
        let blip = doc
            .find_child(blip_fill, &A::blip())
            .ok_or_else(|| RedlineError::XmlParse {
                message: "Missing blip".to_string(),
                location: slide_path.to_string(),
            })?;
        let embed =
            doc.get_attribute_string(blip, &R::embed())
                .ok_or_else(|| RedlineError::XmlParse {
                    message: "Missing embed".to_string(),
                    location: slide_path.to_string(),
                })?;

        // Resolve image part path
        let image_path = Self::resolve_relationship(package, slide_path, &embed)?;
        let image_data =
            package
                .get_part(&image_path)
                .ok_or_else(|| RedlineError::MissingPart {
                    part_path: image_path.clone(),
                    document_type: "PPTX".to_string(),
                })?;

        Ok(sha256_hash_bytes(image_data))
    }

    fn extract_table_hash(doc: &XmlDocument, element: indextree::NodeId) -> Result<String> {
        let graphic =
            doc.find_child(element, &A::graphic())
                .ok_or_else(|| RedlineError::XmlParse {
                    message: "Missing graphic".to_string(),
                    location: "table element".to_string(),
                })?;
        let graphic_data =
            doc.find_child(graphic, &A::graphic_data())
                .ok_or_else(|| RedlineError::XmlParse {
                    message: "Missing graphicData".to_string(),
                    location: "table element".to_string(),
                })?;
        let tbl =
            doc.find_child(graphic_data, &A::tbl())
                .ok_or_else(|| RedlineError::XmlParse {
                    message: "Missing tbl".to_string(),
                    location: "table element".to_string(),
                })?;

        // Hash table content
        let mut content_builder = String::new();
        for tr in doc.children(tbl) {
            if doc.name(tr) != Some(&A::tr()) {
                continue;
            }
            for tc in doc.children(tr) {
                if doc.name(tc) != Some(&A::tc()) {
                    continue;
                }
                if let Some(tx_body) = doc.find_child(tc, &A::tx_body()) {
                    let text = Self::extract_text_body(doc, tx_body);
                    content_builder.push_str(&text.plain_text);
                    content_builder.push('|');
                }
            }
            content_builder.push_str("||");
        }

        Ok(sha256_hash_string(&content_builder))
    }

    fn extract_chart_hash(
        doc: &XmlDocument,
        package: &crate::package::OoxmlPackage,
        slide_path: &str,
        element: indextree::NodeId,
    ) -> Result<String> {
        let graphic =
            doc.find_child(element, &A::graphic())
                .ok_or_else(|| RedlineError::XmlParse {
                    message: "Missing graphic".to_string(),
                    location: slide_path.to_string(),
                })?;
        let graphic_data =
            doc.find_child(graphic, &A::graphic_data())
                .ok_or_else(|| RedlineError::XmlParse {
                    message: "Missing graphicData".to_string(),
                    location: slide_path.to_string(),
                })?;

        // Note: C# uses C.chart here, but we need to define that namespace
        // For now, look for the chart element by local name
        let chart_ref = doc
            .children(graphic_data)
            .find(|child| doc.name(*child).map(|n| n.local_name.as_str()) == Some("chart"))
            .ok_or_else(|| RedlineError::XmlParse {
                message: "Missing chart".to_string(),
                location: slide_path.to_string(),
            })?;

        let r_id = doc
            .get_attribute_string(chart_ref, &R::id())
            .ok_or_else(|| RedlineError::XmlParse {
                message: "Missing chart rId".to_string(),
                location: slide_path.to_string(),
            })?;

        // Resolve chart part path
        let chart_path = Self::resolve_relationship(package, slide_path, &r_id)?;
        let chart_doc = package.get_xml_part(&chart_path)?;
        let chart_xml = if let Some(root) = chart_doc.root() {
            chart_doc.to_xml_string(root)
        } else {
            String::new()
        };

        Ok(sha256_hash_string(&chart_xml))
    }

    fn extract_notes_text(
        package: &crate::package::OoxmlPackage,
        notes_path: &str,
    ) -> Result<String> {
        let notes_doc = package.get_xml_part(notes_path)?;
        let notes_root = notes_doc.root().ok_or_else(|| RedlineError::XmlParse {
            message: "Missing notes root".to_string(),
            location: notes_path.to_string(),
        })?;

        let c_sld = notes_doc
            .find_child(notes_root, &P::c_sld())
            .ok_or_else(|| RedlineError::XmlParse {
                message: "Missing cSld in notes".to_string(),
                location: notes_path.to_string(),
            })?;
        let sp_tree =
            notes_doc
                .find_child(c_sld, &P::sp_tree())
                .ok_or_else(|| RedlineError::XmlParse {
                    message: "Missing spTree in notes".to_string(),
                    location: notes_path.to_string(),
                })?;

        let mut text_builder = String::new();
        for sp in notes_doc.children(sp_tree) {
            if notes_doc.name(sp) != Some(&P::sp()) {
                continue;
            }
            if let Some(tx_body) = notes_doc.find_child(sp, &P::tx_body()) {
                let text_body = Self::extract_text_body(&notes_doc, tx_body);
                if !text_builder.is_empty() {
                    text_builder.push('\n');
                }
                text_builder.push_str(&text_body.plain_text);
            }
        }

        Ok(text_builder)
    }

    // Helper functions for resolving relationships

    fn resolve_relationship(
        package: &crate::package::OoxmlPackage,
        source_path: &str,
        r_id: &str,
    ) -> Result<String> {
        // Get relationship file path
        let rels_path = Self::get_rels_path(source_path);

        // Parse relationships
        let rels_doc = package.get_xml_part(&rels_path)?;
        let rels_root = rels_doc.root().ok_or_else(|| RedlineError::XmlParse {
            message: "Missing rels root".to_string(),
            location: rels_path.clone(),
        })?;

        let id_name = XName::local("Id");
        let target_name = XName::local("Target");

        for rel in rels_doc.children(rels_root) {
            if let Some(id) = rels_doc.get_attribute_string(rel, &id_name) {
                if id == r_id {
                    if let Some(target) = rels_doc.get_attribute_string(rel, &target_name) {
                        // Resolve relative path
                        return Ok(Self::resolve_relative_path(source_path, &target));
                    }
                }
            }
        }

        Err(RedlineError::MissingPart {
            part_path: format!("relationship {} from {}", r_id, source_path),
            document_type: "PPTX".to_string(),
        })
    }

    fn get_rels_path(source_path: &str) -> String {
        // Convert "ppt/slides/slide1.xml" to "ppt/slides/_rels/slide1.xml.rels"
        let parts: Vec<&str> = source_path.rsplitn(2, '/').collect();
        if parts.len() == 2 {
            format!("{}/_rels/{}.rels", parts[1], parts[0])
        } else {
            format!("_rels/{}.rels", source_path)
        }
    }

    fn resolve_relative_path(source_path: &str, target: &str) -> String {
        if target.starts_with('/') {
            return target[1..].to_string();
        }

        // Get directory of source
        let parts: Vec<&str> = source_path.rsplitn(2, '/').collect();
        let base_dir = if parts.len() == 2 { parts[1] } else { "" };

        // Resolve ../ references
        let mut path_parts: Vec<&str> = if base_dir.is_empty() {
            Vec::new()
        } else {
            base_dir.split('/').collect()
        };

        for part in target.split('/') {
            if part == ".." {
                path_parts.pop();
            } else if part != "." && !part.is_empty() {
                path_parts.push(part);
            }
        }

        path_parts.join("/")
    }

    fn get_layout_relationship(
        package: &crate::package::OoxmlPackage,
        rels_path: &str,
    ) -> Result<String> {
        let rels_doc = package.get_xml_part(rels_path)?;
        let rels_root = rels_doc.root().ok_or_else(|| RedlineError::XmlParse {
            message: "Missing rels root".to_string(),
            location: rels_path.to_string(),
        })?;

        let type_name = XName::local("Type");
        let id_name = XName::local("Id");

        for rel in rels_doc.children(rels_root) {
            if let Some(type_) = rels_doc.get_attribute_string(rel, &type_name) {
                if type_.ends_with("/slideLayout") {
                    if let Some(id) = rels_doc.get_attribute_string(rel, &id_name) {
                        return Ok(id);
                    }
                }
            }
        }

        Err(RedlineError::MissingPart {
            part_path: "slideLayout relationship".to_string(),
            document_type: "PPTX".to_string(),
        })
    }

    fn get_notes_relationship(
        package: &crate::package::OoxmlPackage,
        rels_path: &str,
    ) -> Result<String> {
        let rels_doc = package.get_xml_part(rels_path)?;
        let rels_root = rels_doc.root().ok_or_else(|| RedlineError::XmlParse {
            message: "Missing rels root".to_string(),
            location: rels_path.to_string(),
        })?;

        let type_name = XName::local("Type");
        let id_name = XName::local("Id");

        for rel in rels_doc.children(rels_root) {
            if let Some(type_) = rels_doc.get_attribute_string(rel, &type_name) {
                if type_.ends_with("/notesSlide") {
                    if let Some(id) = rels_doc.get_attribute_string(rel, &id_name) {
                        return Ok(id);
                    }
                }
            }
        }

        Err(RedlineError::MissingPart {
            part_path: "notesSlide relationship".to_string(),
            document_type: "PPTX".to_string(),
        })
    }
}
