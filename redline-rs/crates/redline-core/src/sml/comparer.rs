use super::document::SmlDocument;
use super::result::SmlComparisonResult;
use super::settings::SmlComparerSettings;
use crate::error::Result;

pub struct SmlComparer;

impl SmlComparer {
    pub fn compare(
        source1: &SmlDocument,
        source2: &SmlDocument,
        settings: Option<&SmlComparerSettings>,
    ) -> Result<SmlComparisonResult> {
        let _settings = settings.cloned().unwrap_or_default();
        
        todo!("SmlComparer.compare not yet implemented - Phase 3")
    }

    pub fn produce_marked_workbook(
        source1: &SmlDocument,
        source2: &SmlDocument,
        settings: Option<&SmlComparerSettings>,
    ) -> Result<SmlDocument> {
        let _settings = settings.cloned().unwrap_or_default();
        
        todo!("SmlComparer.produce_marked_workbook not yet implemented - Phase 3")
    }
}
