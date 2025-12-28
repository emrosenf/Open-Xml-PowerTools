mod comparer;
mod settings;
mod document;
mod result;
mod signatures;
mod types;
pub mod data_retriever;

pub use comparer::SmlComparer;
pub use settings::SmlComparerSettings;
pub use document::SmlDocument;
pub use result::SmlComparisonResult;
pub use data_retriever::SmlDataRetriever;
pub use signatures::CellFormatSignature;
pub use types::{SmlChange, SmlChangeType};

// Internal signature types used by comparer
pub(crate) use signatures::{
    WorkbookSignature, WorksheetSignature, CellSignature,
    CommentSignature, DataValidationSignature, HyperlinkSignature,
};
