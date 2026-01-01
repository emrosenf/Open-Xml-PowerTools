//! Comments handling for WML documents
//!
//! This module handles the five comment-related files in OOXML:
//! - comments.xml - Main comment content
//! - commentsExtended.xml - Extended properties (done status, reply threading)
//! - commentsIds.xml - Durable IDs for collaboration
//! - commentsExtensible.xml - UTC timestamps
//! - people.xml - Author identity information

use crate::error::RedlineError;
use crate::package::content_types::content_type_values;
use crate::package::relationships::relationship_types;
use crate::package::OoxmlPackage;
use crate::xml::arena::XmlDocument;
use crate::xml::namespaces::{MC, W, W14, W15, W16CEX, W16CID};
use crate::xml::node::XmlNodeData;
use crate::xml::xname::{XAttribute, XName};
use crate::Result;
use indextree::NodeId;
use std::collections::{HashMap, HashSet};

/// Information about a single comment
#[derive(Debug, Clone)]
pub struct CommentInfo {
    /// Comment ID (w:id attribute)
    pub id: String,
    /// Author display name
    pub author: String,
    /// Author initials
    pub initials: String,
    /// Date in local time (ISO 8601)
    pub date: String,
    /// Date in UTC (ISO 8601)
    pub date_utc: Option<String>,
    /// Paragraph ID (w14:paraId)
    pub para_id: String,
    /// Durable ID for collaboration
    pub durable_id: String,
    /// Whether the comment is resolved
    pub done: bool,
    /// Parent comment's paraId (for replies)
    pub parent_para_id: Option<String>,
    /// Comment text content
    pub text: String,
}

/// Information about a person/author
#[derive(Debug, Clone)]
pub struct PersonInfo {
    /// Author display name
    pub author: String,
    /// Identity provider (e.g., "Windows Live", "AD")
    pub provider_id: String,
    /// User ID from the identity provider
    pub user_id: String,
}

/// Collected comment data from source documents
#[derive(Debug, Default)]
pub struct CommentsData {
    pub comments: Vec<CommentInfo>,
    pub people: Vec<PersonInfo>,
}

impl CommentsData {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn is_empty(&self) -> bool {
        self.comments.is_empty()
    }
}

/// Extract comments data from an OOXML package
pub fn extract_comments_data(package: &OoxmlPackage) -> Result<CommentsData> {
    let mut data = CommentsData::new();

    // Parse comments.xml if present
    if let Some(_) = package.get_part("word/comments.xml") {
        let comments_doc = package.get_xml_part("word/comments.xml")?;
        parse_comments_xml(&comments_doc, &mut data)?;
    }

    // Parse commentsExtended.xml if present
    if let Some(_) = package.get_part("word/commentsExtended.xml") {
        let ext_doc = package.get_xml_part("word/commentsExtended.xml")?;
        parse_comments_extended(&ext_doc, &mut data)?;
    }

    // Parse commentsIds.xml if present
    if let Some(_) = package.get_part("word/commentsIds.xml") {
        let ids_doc = package.get_xml_part("word/commentsIds.xml")?;
        parse_comments_ids(&ids_doc, &mut data)?;
    }

    // Parse commentsExtensible.xml if present
    if let Some(_) = package.get_part("word/commentsExtensible.xml") {
        let ext_doc = package.get_xml_part("word/commentsExtensible.xml")?;
        parse_comments_extensible(&ext_doc, &mut data)?;
    }

    // Parse people.xml if present
    if let Some(_) = package.get_part("word/people.xml") {
        let people_doc = package.get_xml_part("word/people.xml")?;
        parse_people_xml(&people_doc, &mut data)?;
    }

    Ok(data)
}

fn parse_comments_xml(doc: &XmlDocument, data: &mut CommentsData) -> Result<()> {
    let root = doc.root().ok_or_else(|| RedlineError::XmlParse {
        message: "No root element".to_string(),
        location: "comments.xml".to_string(),
    })?;

    for node_id in doc.descendants(root) {
        if let Some(node_data) = doc.get(node_id) {
            if let Some(name) = node_data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "comment" {
                    if let Some(attrs) = node_data.attributes() {
                        let id = get_attr(attrs, W::NS, "id").unwrap_or_default();
                        let author = get_attr(attrs, W::NS, "author").unwrap_or_default();
                        let initials = get_attr(attrs, W::NS, "initials").unwrap_or_default();
                        let date = get_attr(attrs, W::NS, "date").unwrap_or_default();

                        // Find paraId from first paragraph child
                        let para_id = find_para_id(doc, node_id).unwrap_or_else(generate_para_id);

                        // Extract text content
                        let text = extract_comment_text(doc, node_id);

                        data.comments.push(CommentInfo {
                            id,
                            author,
                            initials,
                            date: date.clone(),
                            date_utc: None, // Will be filled from commentsExtensible
                            para_id,
                            durable_id: generate_durable_id(), // Will be overwritten if found in commentsIds
                            done: false,
                            parent_para_id: None,
                            text,
                        });
                    }
                }
            }
        }
    }

    Ok(())
}

fn parse_comments_extended(doc: &XmlDocument, data: &mut CommentsData) -> Result<()> {
    let root = doc.root().ok_or_else(|| RedlineError::XmlParse {
        message: "No root element".to_string(),
        location: "commentsExtended.xml".to_string(),
    })?;

    // Build a map of paraId -> comment index (using owned strings)
    let para_id_map: HashMap<String, usize> = data
        .comments
        .iter()
        .enumerate()
        .map(|(i, c)| (c.para_id.clone(), i))
        .collect();

    // Collect updates first
    let mut updates: Vec<(usize, bool, Option<String>)> = Vec::new();

    for node_id in doc.descendants(root) {
        if let Some(node_data) = doc.get(node_id) {
            if let Some(name) = node_data.name() {
                if name.namespace.as_deref() == Some(W15::NS) && name.local_name == "commentEx" {
                    if let Some(attrs) = node_data.attributes() {
                        let para_id = get_attr(attrs, W15::NS, "paraId");
                        let done = get_attr(attrs, W15::NS, "done");
                        let parent = get_attr(attrs, W15::NS, "paraIdParent");

                        if let Some(pid) = para_id {
                            if let Some(&idx) = para_id_map.get(&pid) {
                                updates.push((idx, done.as_deref() == Some("1"), parent));
                            }
                        }
                    }
                }
            }
        }
    }

    // Apply updates
    for (idx, done, parent) in updates {
        data.comments[idx].done = done;
        data.comments[idx].parent_para_id = parent;
    }

    Ok(())
}

fn parse_comments_ids(doc: &XmlDocument, data: &mut CommentsData) -> Result<()> {
    let root = doc.root().ok_or_else(|| RedlineError::XmlParse {
        message: "No root element".to_string(),
        location: "commentsIds.xml".to_string(),
    })?;

    let para_id_map: HashMap<String, usize> = data
        .comments
        .iter()
        .enumerate()
        .map(|(i, c)| (c.para_id.clone(), i))
        .collect();

    // Collect updates first
    let mut updates: Vec<(usize, String)> = Vec::new();

    for node_id in doc.descendants(root) {
        if let Some(node_data) = doc.get(node_id) {
            if let Some(name) = node_data.name() {
                if name.namespace.as_deref() == Some(W16CID::NS) && name.local_name == "commentId" {
                    if let Some(attrs) = node_data.attributes() {
                        let para_id = get_attr(attrs, W16CID::NS, "paraId");
                        let durable_id = get_attr(attrs, W16CID::NS, "durableId");

                        if let (Some(pid), Some(did)) = (para_id, durable_id) {
                            if let Some(&idx) = para_id_map.get(&pid) {
                                updates.push((idx, did));
                            }
                        }
                    }
                }
            }
        }
    }

    // Apply updates
    for (idx, durable_id) in updates {
        data.comments[idx].durable_id = durable_id;
    }

    Ok(())
}

fn parse_comments_extensible(doc: &XmlDocument, data: &mut CommentsData) -> Result<()> {
    let root = doc.root().ok_or_else(|| RedlineError::XmlParse {
        message: "No root element".to_string(),
        location: "commentsExtensible.xml".to_string(),
    })?;

    // Build map of durableId -> comment index (using owned strings)
    let durable_id_map: HashMap<String, usize> = data
        .comments
        .iter()
        .enumerate()
        .map(|(i, c)| (c.durable_id.clone(), i))
        .collect();

    // Collect updates first
    let mut updates: Vec<(usize, String)> = Vec::new();

    for node_id in doc.descendants(root) {
        if let Some(node_data) = doc.get(node_id) {
            if let Some(name) = node_data.name() {
                if name.namespace.as_deref() == Some(W16CEX::NS)
                    && name.local_name == "commentExtensible"
                {
                    if let Some(attrs) = node_data.attributes() {
                        let durable_id = get_attr(attrs, W16CEX::NS, "durableId");
                        let date_utc = get_attr(attrs, W16CEX::NS, "dateUtc");

                        if let (Some(did), Some(utc)) = (durable_id, date_utc) {
                            if let Some(&idx) = durable_id_map.get(&did) {
                                updates.push((idx, utc));
                            }
                        }
                    }
                }
            }
        }
    }

    // Apply updates
    for (idx, date_utc) in updates {
        data.comments[idx].date_utc = Some(date_utc);
    }

    Ok(())
}

fn parse_people_xml(doc: &XmlDocument, data: &mut CommentsData) -> Result<()> {
    let root = doc.root().ok_or_else(|| RedlineError::XmlParse {
        message: "No root element".to_string(),
        location: "people.xml".to_string(),
    })?;

    for node_id in doc.descendants(root) {
        if let Some(node_data) = doc.get(node_id) {
            if let Some(name) = node_data.name() {
                if name.namespace.as_deref() == Some(W15::NS) && name.local_name == "person" {
                    if let Some(attrs) = node_data.attributes() {
                        let author = get_attr(attrs, W15::NS, "author").unwrap_or_default();

                        // Find presenceInfo child
                        let (provider_id, user_id) = find_presence_info(doc, node_id);

                        data.people.push(PersonInfo {
                            author,
                            provider_id,
                            user_id,
                        });
                    }
                }
            }
        }
    }

    Ok(())
}

fn find_presence_info(doc: &XmlDocument, person_node: NodeId) -> (String, String) {
    for child_id in doc.children(person_node) {
        if let Some(child_data) = doc.get(child_id) {
            if let Some(name) = child_data.name() {
                if name.namespace.as_deref() == Some(W15::NS) && name.local_name == "presenceInfo" {
                    if let Some(attrs) = child_data.attributes() {
                        let provider = get_attr(attrs, W15::NS, "providerId").unwrap_or_default();
                        let user = get_attr(attrs, W15::NS, "userId").unwrap_or_default();
                        return (provider, user);
                    }
                }
            }
        }
    }
    (String::new(), String::new())
}

fn get_attr(attrs: &[XAttribute], ns: &str, local: &str) -> Option<String> {
    attrs
        .iter()
        .find(|a| a.name.local_name == local && a.name.namespace.as_deref() == Some(ns))
        .map(|a| a.value.clone())
}

fn find_para_id(doc: &XmlDocument, comment_node: NodeId) -> Option<String> {
    for child_id in doc.children(comment_node) {
        if let Some(child_data) = doc.get(child_id) {
            if let Some(name) = child_data.name() {
                if name.namespace.as_deref() == Some(W::NS) && name.local_name == "p" {
                    if let Some(attrs) = child_data.attributes() {
                        return get_attr(attrs, W14::NS, "paraId");
                    }
                }
            }
        }
    }
    None
}

fn extract_comment_text(doc: &XmlDocument, comment_node: NodeId) -> String {
    let mut text = String::new();
    for node_id in doc.descendants(comment_node) {
        if let Some(XmlNodeData::Text(t)) = doc.get(node_id) {
            text.push_str(t);
        }
    }
    text
}

fn generate_para_id() -> String {
    // Use first 8 chars of UUID for a random hex ID
    let uuid = uuid::Uuid::new_v4();
    uuid.as_simple().to_string()[..8].to_uppercase()
}

fn generate_durable_id() -> String {
    // Use first 8 chars of UUID for a random hex ID
    let uuid = uuid::Uuid::new_v4();
    uuid.as_simple().to_string()[..8].to_uppercase()
}

// ============================================================================
// XML Generation Functions
// ============================================================================

/// Build comments.xml document
pub fn build_comments_xml(data: &CommentsData) -> XmlDocument {
    let mut doc = XmlDocument::new();

    let root_attrs = vec![
        XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", "w"), W::NS),
        XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", "w14"), W14::NS),
        XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", "w15"), W15::NS),
        XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", "mc"), MC::NS),
        XAttribute::new(MC::ignorable(), "w14 w15"),
    ];

    let root = doc.add_root(XmlNodeData::element_with_attrs(
        XName::new(W::NS, "comments"),
        root_attrs,
    ));

    for comment in &data.comments {
        let comment_attrs = vec![
            XAttribute::new(W::id(), &comment.id),
            XAttribute::new(W::author(), &comment.author),
            XAttribute::new(XName::new(W::NS, "initials"), &comment.initials),
            XAttribute::new(W::date(), &comment.date),
        ];

        let comment_node = doc.add_child(
            root,
            XmlNodeData::element_with_attrs(XName::new(W::NS, "comment"), comment_attrs),
        );

        // Add paragraph with paraId
        let para_attrs = vec![
            XAttribute::new(W14::paraId(), &comment.para_id),
            XAttribute::new(W14::textId(), "77777777"),
        ];

        let para = doc.add_child(
            comment_node,
            XmlNodeData::element_with_attrs(W::p(), para_attrs),
        );

        // Add paragraph properties with CommentText style
        let p_pr = doc.add_child(para, XmlNodeData::element(W::p_pr()));
        doc.add_child(
            p_pr,
            XmlNodeData::element_with_attrs(
                W::p_style(),
                vec![XAttribute::new(W::val(), "CommentText")],
            ),
        );

        // Add first run with annotationRef (required for comment reference marker)
        let ref_run = doc.add_child(para, XmlNodeData::element(W::r()));

        // Add run properties with CommentReference style
        let ref_r_pr = doc.add_child(ref_run, XmlNodeData::element(W::r_pr()));
        doc.add_child(
            ref_r_pr,
            XmlNodeData::element_with_attrs(
                W::r_style(),
                vec![XAttribute::new(W::val(), "CommentReference")],
            ),
        );

        // Add annotationRef element (the actual reference marker)
        doc.add_child(ref_run, XmlNodeData::element(W::annotation_ref()));

        // Add run with text
        let run = doc.add_child(para, XmlNodeData::element(W::r()));
        let text_elem = doc.add_child(run, XmlNodeData::element(W::t()));
        doc.add_child(text_elem, XmlNodeData::Text(comment.text.clone()));
    }

    doc
}

/// Build commentsExtended.xml document
pub fn build_comments_extended_xml(data: &CommentsData) -> XmlDocument {
    let mut doc = XmlDocument::new();

    let root_attrs = vec![
        XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", "w15"), W15::NS),
        XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", "mc"), MC::NS),
        XAttribute::new(MC::ignorable(), "w15"),
    ];

    let root = doc.add_root(XmlNodeData::element_with_attrs(
        W15::commentsEx(),
        root_attrs,
    ));

    for comment in &data.comments {
        let mut attrs = vec![
            XAttribute::new(W15::paraId(), &comment.para_id),
            XAttribute::new(W15::done(), if comment.done { "1" } else { "0" }),
        ];

        if let Some(ref parent) = comment.parent_para_id {
            attrs.push(XAttribute::new(W15::paraIdParent(), parent));
        }

        doc.add_child(
            root,
            XmlNodeData::element_with_attrs(W15::commentEx(), attrs),
        );
    }

    doc
}

/// Build commentsIds.xml document
pub fn build_comments_ids_xml(data: &CommentsData) -> XmlDocument {
    let mut doc = XmlDocument::new();

    let root_attrs = vec![
        XAttribute::new(
            XName::new("http://www.w3.org/2000/xmlns/", "w16cid"),
            W16CID::NS,
        ),
        XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", "mc"), MC::NS),
        XAttribute::new(MC::ignorable(), "w16cid"),
    ];

    let root = doc.add_root(XmlNodeData::element_with_attrs(
        W16CID::commentsIds(),
        root_attrs,
    ));

    for comment in &data.comments {
        let attrs = vec![
            XAttribute::new(W16CID::paraId(), &comment.para_id),
            XAttribute::new(W16CID::durableId(), &comment.durable_id),
        ];

        doc.add_child(
            root,
            XmlNodeData::element_with_attrs(W16CID::commentId(), attrs),
        );
    }

    doc
}

/// Build commentsExtensible.xml document
pub fn build_comments_extensible_xml(data: &CommentsData) -> XmlDocument {
    let mut doc = XmlDocument::new();

    let root_attrs = vec![
        XAttribute::new(
            XName::new("http://www.w3.org/2000/xmlns/", "w16cex"),
            W16CEX::NS,
        ),
        XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", "mc"), MC::NS),
        XAttribute::new(MC::ignorable(), "w16cex"),
    ];

    let root = doc.add_root(XmlNodeData::element_with_attrs(
        W16CEX::commentsExtensible(),
        root_attrs,
    ));

    for comment in &data.comments {
        // Use UTC date if available, otherwise convert local date
        let date_utc = comment.date_utc.clone().unwrap_or_else(|| {
            // Simple fallback - just use the local date as-is
            comment.date.clone()
        });

        let attrs = vec![
            XAttribute::new(W16CEX::durableId(), &comment.durable_id),
            XAttribute::new(W16CEX::dateUtc(), &date_utc),
        ];

        doc.add_child(
            root,
            XmlNodeData::element_with_attrs(W16CEX::commentExtensible(), attrs),
        );
    }

    doc
}

/// Build people.xml document
pub fn build_people_xml(data: &CommentsData) -> XmlDocument {
    let mut doc = XmlDocument::new();

    let root_attrs = vec![
        XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", "w15"), W15::NS),
        XAttribute::new(XName::new("http://www.w3.org/2000/xmlns/", "mc"), MC::NS),
        XAttribute::new(MC::ignorable(), "w15"),
    ];

    let root = doc.add_root(XmlNodeData::element_with_attrs(W15::people(), root_attrs));

    // Collect unique authors
    let mut seen_authors = HashSet::new();
    let mut people_to_add = Vec::new();

    // First add from parsed people data
    for person in &data.people {
        if !seen_authors.contains(&person.author) {
            seen_authors.insert(person.author.clone());
            people_to_add.push(person.clone());
        }
    }

    // Then add any authors from comments that aren't in people list
    for comment in &data.comments {
        if !seen_authors.contains(&comment.author) {
            seen_authors.insert(comment.author.clone());
            people_to_add.push(PersonInfo {
                author: comment.author.clone(),
                provider_id: String::new(),
                user_id: String::new(),
            });
        }
    }

    for person in people_to_add {
        let person_attrs = vec![XAttribute::new(W15::author(), &person.author)];

        let person_node = doc.add_child(
            root,
            XmlNodeData::element_with_attrs(W15::person(), person_attrs),
        );

        // Only add presenceInfo if we have provider data
        if !person.provider_id.is_empty() || !person.user_id.is_empty() {
            let presence_attrs = vec![
                XAttribute::new(W15::providerId(), &person.provider_id),
                XAttribute::new(W15::userId(), &person.user_id),
            ];

            doc.add_child(
                person_node,
                XmlNodeData::element_with_attrs(W15::presenceInfo(), presence_attrs),
            );
        }
    }

    doc
}

// ============================================================================
// Integration Functions
// ============================================================================

/// Add comment-related parts to a package
pub fn add_comments_to_package(package: &mut OoxmlPackage, data: &CommentsData) -> Result<()> {
    if data.is_empty() {
        return Ok(());
    }

    // Build and add the comment XML files
    let comments_doc = build_comments_xml(data);
    package.put_xml_part("word/comments.xml", &comments_doc)?;

    let comments_ext_doc = build_comments_extended_xml(data);
    package.put_xml_part("word/commentsExtended.xml", &comments_ext_doc)?;

    let comments_ids_doc = build_comments_ids_xml(data);
    package.put_xml_part("word/commentsIds.xml", &comments_ids_doc)?;

    let comments_extensible_doc = build_comments_extensible_xml(data);
    package.put_xml_part("word/commentsExtensible.xml", &comments_extensible_doc)?;

    let people_doc = build_people_xml(data);
    package.put_xml_part("word/people.xml", &people_doc)?;

    // Update [Content_Types].xml to include the new parts
    update_content_types(package)?;

    // Update word/_rels/document.xml.rels to include relationships
    update_document_relationships(package)?;

    Ok(())
}

const RELS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

fn update_content_types(package: &mut OoxmlPackage) -> Result<()> {
    let ct_doc = package.get_xml_part("[Content_Types].xml")?;
    let mut new_doc = XmlDocument::new();

    let root = ct_doc.root().ok_or_else(|| RedlineError::XmlParse {
        message: "No root element".to_string(),
        location: "[Content_Types].xml".to_string(),
    })?;

    // Clone the root element
    let root_data = ct_doc.get(root).ok_or_else(|| RedlineError::XmlParse {
        message: "Cannot get root data".to_string(),
        location: "[Content_Types].xml".to_string(),
    })?;

    let new_root = new_doc.add_root(root_data.clone());

    // Collect existing PartName values to avoid duplicates
    let mut existing_part_names: HashSet<String> = HashSet::new();
    for child_id in ct_doc.children(root) {
        if let Some(child_data) = ct_doc.get(child_id) {
            if let Some(attrs) = child_data.attributes() {
                for attr in attrs {
                    if attr.name.local_name == "PartName" {
                        existing_part_names.insert(attr.value.clone());
                    }
                }
            }
        }
    }

    // Clone all existing children
    for child_id in ct_doc.children(root) {
        clone_node_to_doc(&ct_doc, child_id, &mut new_doc, new_root);
    }

    // Add new Override entries for comment parts (only if not already present)
    let new_parts = [
        ("/word/comments.xml", content_type_values::WORD_COMMENTS),
        (
            "/word/commentsExtended.xml",
            content_type_values::WORD_COMMENTS_EXTENDED,
        ),
        (
            "/word/commentsIds.xml",
            content_type_values::WORD_COMMENTS_IDS,
        ),
        (
            "/word/commentsExtensible.xml",
            content_type_values::WORD_COMMENTS_EXTENSIBLE,
        ),
        ("/word/people.xml", content_type_values::WORD_PEOPLE),
    ];

    for (part_name, content_type) in new_parts {
        // Skip if this part already has a content type entry
        if existing_part_names.contains(part_name) {
            continue;
        }
        let attrs = vec![
            XAttribute::new(XName::local("PartName"), part_name),
            XAttribute::new(XName::local("ContentType"), content_type),
        ];
        new_doc.add_child(
            new_root,
            XmlNodeData::element_with_attrs(XName::new(CONTENT_TYPES_NS, "Override"), attrs),
        );
    }

    package.put_xml_part("[Content_Types].xml", &new_doc)?;
    Ok(())
}

fn update_document_relationships(package: &mut OoxmlPackage) -> Result<()> {
    let rels_path = "word/_rels/document.xml.rels";
    let rels_doc = package.get_xml_part(rels_path)?;
    let mut new_doc = XmlDocument::new();

    let root = rels_doc.root().ok_or_else(|| RedlineError::XmlParse {
        message: "No root element".to_string(),
        location: rels_path.to_string(),
    })?;

    // Clone the root element
    let root_data = rels_doc.get(root).ok_or_else(|| RedlineError::XmlParse {
        message: "Cannot get root data".to_string(),
        location: rels_path.to_string(),
    })?;

    let new_root = new_doc.add_root(root_data.clone());

    // Find the max existing rId and collect existing relationship types
    let mut max_id = 0u32;
    let mut existing_rel_types: HashSet<String> = HashSet::new();
    for child_id in rels_doc.children(root) {
        if let Some(child_data) = rels_doc.get(child_id) {
            if let Some(attrs) = child_data.attributes() {
                for attr in attrs {
                    if attr.name.local_name == "Id" {
                        if let Some(num_str) = attr.value.strip_prefix("rId") {
                            if let Ok(num) = num_str.parse::<u32>() {
                                max_id = max_id.max(num);
                            }
                        }
                    }
                    if attr.name.local_name == "Type" {
                        existing_rel_types.insert(attr.value.clone());
                    }
                }
            }
        }
    }

    // Clone all existing children
    for child_id in rels_doc.children(root) {
        clone_node_to_doc(&rels_doc, child_id, &mut new_doc, new_root);
    }

    // Add new relationships for comment parts (only if not already present)
    let new_rels = [
        ("comments.xml", relationship_types::COMMENTS),
        (
            "commentsExtended.xml",
            relationship_types::COMMENTS_EXTENDED,
        ),
        ("commentsIds.xml", relationship_types::COMMENTS_IDS),
        (
            "commentsExtensible.xml",
            relationship_types::COMMENTS_EXTENSIBLE,
        ),
        ("people.xml", relationship_types::PEOPLE),
    ];

    let mut next_id = max_id + 1;
    for (target, rel_type) in new_rels.iter() {
        // Skip if this relationship type already exists
        if existing_rel_types.contains(*rel_type) {
            continue;
        }
        let rid = format!("rId{}", next_id);
        next_id += 1;
        let attrs = vec![
            XAttribute::new(XName::local("Id"), &rid),
            XAttribute::new(XName::local("Type"), *rel_type),
            XAttribute::new(XName::local("Target"), *target),
        ];
        new_doc.add_child(
            new_root,
            XmlNodeData::element_with_attrs(XName::new(RELS_NS, "Relationship"), attrs),
        );
    }

    package.put_xml_part(rels_path, &new_doc)?;
    Ok(())
}

fn clone_node_to_doc(
    src_doc: &XmlDocument,
    src_node: NodeId,
    dst_doc: &mut XmlDocument,
    dst_parent: NodeId,
) {
    if let Some(src_data) = src_doc.get(src_node) {
        // Clone the node data, filtering out redundant xmlns declarations on child elements
        let filtered_data = match src_data {
            XmlNodeData::Element { name, attributes } => {
                // Filter out default xmlns declarations - these are redundant on child elements
                // since the namespace is already declared on the root element
                let filtered_attrs: Vec<XAttribute> = attributes
                    .iter()
                    .filter(|attr| {
                        // Keep attribute if it's NOT a default xmlns declaration
                        // (i.e., xmlns="..." without a prefix)
                        !(attr.name.namespace.is_none() && attr.name.local_name == "xmlns")
                    })
                    .cloned()
                    .collect();
                XmlNodeData::Element {
                    name: name.clone(),
                    attributes: filtered_attrs,
                }
            }
            other => other.clone(),
        };

        let new_node = dst_doc.add_child(dst_parent, filtered_data);

        // Recursively clone children
        for child_id in src_doc.children(src_node) {
            clone_node_to_doc(src_doc, child_id, dst_doc, new_node);
        }
    }
}

/// Merge comments from two source documents
pub fn merge_comments(source1: &CommentsData, source2: &CommentsData) -> CommentsData {
    let mut merged = CommentsData::new();

    // Add all comments from source1
    merged.comments.extend(source1.comments.clone());

    // Track existing paraIds to avoid duplicates (owned strings to avoid borrow issues)
    let existing_para_ids: HashSet<String> =
        merged.comments.iter().map(|c| c.para_id.clone()).collect();

    // Add comments from source2 that don't duplicate paraIds
    for comment in &source2.comments {
        if !existing_para_ids.contains(&comment.para_id) {
            merged.comments.push(comment.clone());
        }
    }

    // Merge people - deduplicate by author name
    let mut seen_authors = HashSet::new();
    for person in source1.people.iter().chain(source2.people.iter()) {
        if !seen_authors.contains(&person.author) {
            seen_authors.insert(person.author.clone());
            merged.people.push(person.clone());
        }
    }

    merged
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_generate_para_id() {
        let id = generate_para_id();
        assert_eq!(id.len(), 8);
        // Should be valid hex
        assert!(id.chars().all(|c| c.is_ascii_hexdigit()));
    }

    #[test]
    fn test_build_empty_comments() {
        let data = CommentsData::new();
        let doc = build_comments_xml(&data);
        assert!(doc.root().is_some());
    }
}
