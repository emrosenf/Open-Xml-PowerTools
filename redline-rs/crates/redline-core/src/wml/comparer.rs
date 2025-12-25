use super::document::WmlDocument;
use super::settings::WmlComparerSettings;
use crate::error::Result;
use crate::types::Revision;

pub struct WmlComparer;

impl WmlComparer {
    pub fn compare(
        _source1: &WmlDocument,
        _source2: &WmlDocument,
        settings: Option<&WmlComparerSettings>,
    ) -> Result<WmlDocument> {
        let _settings = settings.cloned().unwrap_or_default();
        
        todo!("WmlComparer.compare not yet implemented - Phase 2")
    }

    pub fn get_revisions(
        _document: &WmlDocument,
        settings: Option<&WmlComparerSettings>,
    ) -> Result<Vec<Revision>> {
        let _settings = settings.cloned().unwrap_or_default();
        
        todo!("WmlComparer.get_revisions not yet implemented - Phase 2")
    }
}
