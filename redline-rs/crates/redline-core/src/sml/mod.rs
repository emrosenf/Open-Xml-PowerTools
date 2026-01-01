mod canonicalize;
mod change_list;
mod comparer;
pub mod data_retriever;
mod diff;
mod document;
mod markup;
mod patch;
mod result;
mod settings;
mod signatures;
mod types;

pub use change_list::build_change_list;
pub use comparer::SmlComparer;
pub use data_retriever::SmlDataRetriever;
pub use document::SmlDocument;
pub use patch::{apply_sml_changes, revert_sml_changes};
pub use result::SmlComparisonResult;
pub use settings::SmlComparerSettings;
pub use signatures::CellFormatSignature;
pub use types::{
    SmlChange, SmlChangeDetails, SmlChangeListItem, SmlChangeListOptions, SmlChangeType,
};

// Internal signature types used by comparer
pub(crate) use signatures::{
    CellSignature, CommentSignature, DataValidationSignature, HyperlinkSignature,
    WorkbookSignature, WorksheetSignature,
};

// Internal diff engine
pub(crate) use diff::compute_diff;

// Internal canonicalizer
pub(crate) use canonicalize::SmlCanonicalizer;
