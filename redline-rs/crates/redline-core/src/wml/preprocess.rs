//! PreProcessMarkup pipeline - Document preparation for comparison
//!
//! Port of C# WmlComparer.PreProcessMarkup and related functions.
//! This module handles the document preparation steps that must occur
//! before the actual comparison algorithm can run.

use crate::wml::atom_list::assign_unid_to_all_elements;
use crate::wml::block_hash::hash_block_level_content;
use crate::wml::simplify::{simplify_markup, SimplifyMarkupSettings};
use crate::xml::arena::XmlDocument;
use indextree::NodeId;

/// Settings for PreProcessMarkup pipeline
#[derive(Debug, Clone, Default)]
pub struct PreProcessSettings {
    /// Starting ID for footnotes/endnotes (to avoid ID collisions)
    pub starting_id_for_footnotes_endnotes: i32,

    /// Settings for SimplifyMarkup step
    pub simplify_settings: SimplifyMarkupSettings,
}

impl PreProcessSettings {
    /// Create default settings for document comparison
    pub fn for_comparison() -> Self {
        let mut settings = Self::default();

        // Configure SimplifyMarkup settings (from C# line 441-455)
        settings.simplify_settings = SimplifyMarkupSettings {
            remove_bookmarks: true,
            accept_revisions: false,
            remove_comments: true,
            remove_content_controls: true,
            remove_field_codes: true,
            remove_go_back_bookmark: true,
            remove_last_rendered_page_break: true,
            remove_permissions: true,
            remove_proof: true,
            remove_smart_tags: true,
            remove_soft_hyphens: true,
            remove_hyperlinks: true,
            ..Default::default()
        };

        settings
    }
}

/// PreProcessMarkup pipeline - Main entry point
///
/// Corresponds to C# WmlComparer.PreProcessMarkup (line 392)
///
/// # Pipeline Steps (EXACT order from C#):
/// 1. SimplifyMarkup - Normalize document (remove bookmarks, comments, etc.)
/// 2. AssignUnidToAllElements - Add pt:Unid GUIDs to all elements
/// 3. (Note: Revision acceptance/rejection happens AFTER this in CompareInternal)
/// 4. (Note: HashBlockLevelContent happens AFTER this in CompareInternal)
///
/// The C# version does extensive OOXML package manipulation (MC processing, etc.)
/// which is handled separately in the Rust port's package layer.
/// This function focuses on the XML tree transformations.
pub fn preprocess_markup(
    doc: &mut XmlDocument,
    root: NodeId,
    settings: &PreProcessSettings,
) -> Result<(), String> {
    // Step 1: SimplifyMarkup (C# line 456)
    simplify_markup(doc, root, &settings.simplify_settings);

    // Step 2: AssignUnidToAllElements (C# line 477 via AddUnidsToMarkupInContentParts -> line 953)
    assign_unid_to_all_elements(doc, root);

    Ok(())
}

/// Hash block-level content for correlation
///
/// Corresponds to C# WmlComparer.HashBlockLevelContent (line 343)
///
/// This is a wrapper around hash_block_level_content that handles
/// the typical workflow in document comparison where:
/// 1. Source document (original) has revisions accepted
/// 2. After-processing document is hashed
/// 3. Hashes are added back to the source document as CorrelatedSHA1Hash
///
/// NOTE: The caller is responsible for:
/// - Creating a separate document with revisions accepted/rejected
/// - Passing both documents to this function
/// - The source_doc will be modified with CorrelatedSHA1Hash attributes
pub fn add_correlated_hashes_from_processed_doc(
    source_doc: &mut XmlDocument,
    source_root: NodeId,
    processed_doc: &XmlDocument,
    processed_root: NodeId,
) {
    // Use default hashing settings (case-insensitive, conflate spaces, track formatting)
    use crate::wml::block_hash::HashingSettings;
    let settings = HashingSettings {
        case_insensitive: true,
        conflate_spaces: true,
        track_formatting_changes: true,
    };

    // Hash the block-level content in the processed document and
    // add CorrelatedSHA1Hash attributes to the source document (C# line 368-383)
    hash_block_level_content(
        source_doc,
        source_root,
        processed_doc,
        processed_root,
        &settings,
    );
}

/// Repair Unids after revision acceptance
///
/// Corresponds to C# WmlComparer logic at line 270-279
///
/// After accepting revisions, some unids may have been removed by the revision accepter.
/// This function adds GUIDs back to elements that had them removed.
pub fn repair_unids_after_revision_acceptance(
    doc: &mut XmlDocument,
    root: NodeId,
) -> Result<(), String> {
    // Simply re-assign unids to all elements
    // (existing elements with unids will keep them, new ones get fresh unids)
    assign_unid_to_all_elements(doc, root);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_preprocess_settings_for_comparison() {
        let settings = PreProcessSettings::for_comparison();

        assert!(settings.simplify_settings.remove_bookmarks);
        assert!(!settings.simplify_settings.accept_revisions);
        assert!(settings.simplify_settings.remove_comments);
        assert!(settings.simplify_settings.remove_content_controls);
        assert!(settings.simplify_settings.remove_field_codes);
        assert!(settings.simplify_settings.remove_go_back_bookmark);
        assert!(settings.simplify_settings.remove_last_rendered_page_break);
        assert!(settings.simplify_settings.remove_permissions);
        assert!(settings.simplify_settings.remove_proof);
        assert!(settings.simplify_settings.remove_smart_tags);
        assert!(settings.simplify_settings.remove_soft_hyphens);
        assert!(settings.simplify_settings.remove_hyperlinks);
    }
}
