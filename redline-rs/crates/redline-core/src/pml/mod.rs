pub mod canonicalize;
mod change_list;
mod comparer;
mod diff;
mod document;
pub mod markup;
mod patch;
mod result;
mod settings;
pub mod shape_match;
pub mod slide_matching;
mod types;

pub use canonicalize::PmlCanonicalizer;
pub use change_list::build_change_list;
pub use comparer::PmlComparer;
pub use diff::PmlDiffEngine;
pub use document::PmlDocument;
pub use markup::render_marked_presentation;
pub use patch::{apply_pml_changes, revert_pml_changes};
pub use result::PmlComparisonResult;
pub use settings::PmlComparerSettings;
pub use types::{
    PmlChange, PmlChangeDetails, PmlChangeListItem, PmlChangeListOptions, PmlChangeType,
    PmlTextChange, PmlWordCount, TextChangeType,
};
