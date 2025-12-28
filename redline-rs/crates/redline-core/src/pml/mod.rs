mod comparer;
mod settings;
mod document;
mod result;
mod diff;
pub mod slide_matching;
pub mod shape_match;
pub mod markup;
pub mod canonicalize;

pub use comparer::PmlComparer;
pub use settings::PmlComparerSettings;
pub use document::PmlDocument;
pub use result::{PmlComparisonResult, PmlChange, PmlChangeType};
pub use diff::PmlDiffEngine;
pub use markup::render_marked_presentation;
pub use canonicalize::PmlCanonicalizer;
