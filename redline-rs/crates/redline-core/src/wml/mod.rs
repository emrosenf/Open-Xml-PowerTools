mod atom_list;
mod block_hash;
mod coalesce;
mod comparison_unit;
mod comparer;
mod consolidate;
mod document;
mod formatting;
mod get_revisions;
mod lcs_algorithm;
mod order;
mod preprocess;
mod revision;
mod revision_accepter;
mod revision_processor;
mod settings;
mod simplify;
mod types;

pub use atom_list::{assign_unid_to_all_elements, create_comparison_unit_atom_list};
pub use block_hash::{
    clone_block_level_content_for_hashing, compute_block_hash, hash_block_level_content,
    HashingSettings,
};
pub use comparison_unit::{
    get_comparison_unit_list, AncestorInfo, ComparisonCorrelationStatus, ComparisonUnit,
    ComparisonUnitAtom, ComparisonUnitGroup, ComparisonUnitGroupContents, ComparisonUnitGroupType,
    ComparisonUnitWord, ContentElement, WordSeparatorSettings, generate_unid,
};
pub use comparer::WmlComparer;
pub use lcs_algorithm::{flatten_to_atoms, lcs, CorrelatedSequence, CorrelationStatus};
pub use document::{
    extract_all_text, extract_paragraph_text,
    find_document_body, find_paragraphs,
    WmlDocument,
};
pub use revision::{
    count_revisions, create_deletion, create_insertion, create_paragraph,
    create_paragraph_property_change, create_run_property_change, create_text_run,
    find_max_revision_id, fix_up_revision_ids, get_next_revision_id, is_deletion,
    is_format_change, is_insertion, is_revision_element, is_revision_element_tag,
    reset_revision_id_counter, RevisionSettings,
};
pub use revision_accepter::accept_revisions;
pub use revision_processor::{
    accept_revisions as accept_revisions_processor,
    reject_revisions,
    BlockContentInfo,
};
pub use coalesce::{
    coalesce, CoalesceResult, pt_status, pt_unid, PT_STATUS_NS
};
pub use consolidate::{consolidate, consolidate_with_settings};
pub use get_revisions::get_revisions;
pub use settings::{
    WmlComparerSettings, WmlComparerConsolidateSettings, WmlRevisedDocumentInfo,
    WmlComparerRevisionType, WmlComparerRevision,
};
pub use simplify::{
    SimplifyMarkupSettings, simplify_markup, merge_adjacent_superfluous_runs,
    transform_element_to_single_character_runs,
};
pub use types::{
    RevisionCounts, WmlChange, WmlChangeDetails, WmlChangeListItem, WmlChangeListOptions,
    WmlChangeType, WmlComparisonResult, WmlWordCount,
};
pub use formatting::{
    compute_normalized_rpr, compute_formatting_signature, compute_formatting_signature_hash,
    formatting_differs,
};
pub use preprocess::{
    preprocess_markup, add_correlated_hashes_from_processed_doc,
    repair_unids_after_revision_acceptance, PreProcessSettings,
};
