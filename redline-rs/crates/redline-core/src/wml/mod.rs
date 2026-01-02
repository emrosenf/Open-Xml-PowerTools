mod atom_list;
mod block_hash;
mod change_event;
mod change_list;
mod coalesce;
mod comments;
mod comparer;
mod comparison_unit;
mod consolidate;
mod document;
mod drawing_identity;
mod extract_changes;
mod formatting;
mod get_revisions;
mod lcs_algorithm;
mod order;
pub use order::order_elements_per_standard;
mod preprocess;
mod repro_test;
mod revision;
mod revision_accepter;
mod revision_processor;
mod settings;
mod simplify;
mod types;
mod visual_redline;

pub use atom_list::{
    assign_unid_to_all_elements, create_comparison_unit_atom_list,
    create_comparison_unit_atom_list_with_package,
};
pub use block_hash::{
    clone_block_level_content_for_hashing, compute_block_hash, hash_block_level_content,
    HashingSettings,
};
pub use change_event::{
    count_revisions_from_events, detect_format_changes, emit_change_events, group_adjacent_events,
    ChangeEvent, ChangeEventResult,
};
pub use change_list::build_change_list;
pub use coalesce::{coalesce, pt_status, pt_unid, CoalesceResult, PT_STATUS_NS};
pub use comments::{
    add_comments_to_package, build_comments_extended_xml, build_comments_extensible_xml,
    build_comments_ids_xml, build_comments_xml, build_people_xml, extract_comments_data,
    merge_comments, CommentInfo, CommentsData, PersonInfo,
};
pub use comparer::WmlComparer;
pub use comparison_unit::{
    generate_unid, get_comparison_unit_list, AncestorInfo, ComparisonCorrelationStatus,
    ComparisonUnit, ComparisonUnitAtom, ComparisonUnitGroup, ComparisonUnitGroupContents,
    ComparisonUnitGroupType, ComparisonUnitWord, ContentElement, ContentType,
    WordSeparatorSettings,
};
pub use consolidate::{consolidate, consolidate_with_settings};
pub use document::{
    extract_all_text, extract_paragraph_text, find_document_body, find_endnotes_root,
    find_footnotes_root, find_note_by_id, find_note_paragraphs, find_paragraphs, WmlDocument,
};
pub use drawing_identity::{
    compute_drawing_identity, get_drawing_info, has_textbox_content, DrawingInfo,
};
pub use extract_changes::extract_changes_from_document;
pub use formatting::{
    compute_formatting_signature, compute_formatting_signature_hash, compute_normalized_rpr,
    formatting_differs,
};
pub use get_revisions::get_revisions;
pub use lcs_algorithm::{
    flatten_to_atoms, get_lcs_counters, lcs, reset_lcs_counters, CorrelatedSequence,
    CorrelationStatus,
};
#[cfg(feature = "trace")]
pub use lcs_algorithm::{
    generate_focused_trace, generate_lcs_trace, units_match_filter, MatchedParagraphInfo,
};
pub use preprocess::{
    add_correlated_hashes_from_processed_doc, preprocess_markup,
    repair_unids_after_revision_acceptance, PreProcessSettings,
};
pub use revision::{
    count_revisions, create_deletion, create_insertion, create_paragraph,
    create_paragraph_property_change, create_run_property_change, create_text_run,
    find_max_revision_id, fix_up_revision_ids, get_next_revision_id, is_deletion, is_format_change,
    is_insertion, is_revision_element, is_revision_element_tag, reset_revision_id_counter,
    RevisionSettings,
};
pub use revision_accepter::{accept_revisions, accept_revisions_by_id, reject_revisions_by_id};
pub use revision_processor::{
    accept_revisions as accept_revisions_processor, reject_revisions, BlockContentInfo,
};
#[cfg(feature = "trace")]
pub use settings::{LcsTraceFilter, LcsTraceOperation};
pub use settings::{
    LcsTraceOutput, WmlComparerConsolidateSettings, WmlComparerRevision, WmlComparerRevisionType,
    WmlComparerSettings, WmlRevisedDocumentInfo,
};
pub use simplify::{
    merge_adjacent_superfluous_runs, simplify_markup, transform_element_to_single_character_runs,
    SimplifyMarkupSettings,
};
pub use types::{
    RevisionCounts, WmlChange, WmlChangeDetails, WmlChangeListItem, WmlChangeListOptions,
    WmlChangeType, WmlComparisonResult, WmlWordCount,
};
pub use visual_redline::{render_visual_redline, VisualRedlineResult, VisualRedlineSettings};
