pub mod error;
pub mod types;
pub mod xml;
pub mod hash;
pub mod package;
pub mod util;
pub mod wml;
pub mod sml;
// TODO: Re-enable PML after fixing namespace syntax errors (bd-XXXXX)
// pub mod pml;

pub use error::{RedlineError, Result};

pub use wml::{WmlComparer, WmlComparerSettings, WmlDocument};
pub use sml::{SmlComparer, SmlComparerSettings, SmlDocument, SmlComparisonResult};
// pub use pml::{PmlComparer, PmlComparerSettings, PmlDocument, PmlComparisonResult};
