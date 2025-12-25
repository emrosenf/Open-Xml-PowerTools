mod comparer;
mod document;
mod revision;
mod revision_accepter;
mod settings;
mod types;

pub use comparer::WmlComparer;
pub use document::{
    extract_all_text, extract_paragraph_text, find_document_body, find_paragraphs, WmlDocument,
};
pub use revision::{
    count_revisions, create_deletion, create_insertion, create_paragraph,
    create_paragraph_property_change, create_run_property_change, create_text_run,
    find_max_revision_id, fix_up_revision_ids, get_next_revision_id, is_deletion,
    is_format_change, is_insertion, is_revision_element, is_revision_element_tag,
    reset_revision_id_counter, RevisionSettings,
};
pub use revision_accepter::accept_revisions;
pub use settings::WmlComparerSettings;
pub use types::{
    RevisionCounts, WmlChange, WmlChangeDetails, WmlChangeListItem, WmlChangeListOptions,
    WmlChangeType, WmlComparisonResult, WmlWordCount,
};
