pub mod descendants;
pub mod group;
pub mod strings;
pub mod culture;
pub mod lcs;

pub use descendants::descendants_trimmed;
pub use group::{group_adjacent, rollup};
pub use strings::{make_valid_xml, string_concatenate};
pub use culture::{to_upper_invariant, to_upper_culture};
pub use lcs::{
    CorrelationStatus, Hashable, CorrelatedSequence, LcsSettings, MatchResult,
    find_longest_match, compute_correlation, flatten_correlation,
    DiffType, DiffResult, diff_text,
};
