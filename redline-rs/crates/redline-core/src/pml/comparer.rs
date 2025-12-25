use super::document::PmlDocument;
use super::result::PmlComparisonResult;
use super::settings::PmlComparerSettings;
use crate::error::Result;

pub struct PmlComparer;

impl PmlComparer {
    pub fn compare(
        _source1: &PmlDocument,
        _source2: &PmlDocument,
        settings: Option<&PmlComparerSettings>,
    ) -> Result<PmlComparisonResult> {
        let _settings = settings.cloned().unwrap_or_default();
        
        todo!("PmlComparer.compare not yet implemented - Phase 4")
    }

    pub fn produce_marked_presentation(
        _source1: &PmlDocument,
        _source2: &PmlDocument,
        settings: Option<&PmlComparerSettings>,
    ) -> Result<PmlDocument> {
        let _settings = settings.cloned().unwrap_or_default();
        
        todo!("PmlComparer.produce_marked_presentation not yet implemented - Phase 4")
    }
}
