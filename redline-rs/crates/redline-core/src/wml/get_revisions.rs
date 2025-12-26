use crate::error::Result;
use crate::wml::settings::{WmlComparerRevision, WmlComparerSettings};

pub fn get_revisions(
    _source: &[u8],
    _settings: &WmlComparerSettings,
) -> Result<Vec<WmlComparerRevision>> {
    Err(crate::error::RedlineError::UnsupportedFeature {
        feature: "GetRevisions implementation requires full OpenXML document handling".to_string()
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn get_revisions_not_implemented() {
        let source = vec![];
        let settings = WmlComparerSettings::default();

        let result = get_revisions(&source, &settings);
        assert!(result.is_err());
        assert!(result
            .unwrap_err()
            .to_string()
            .contains("requires full OpenXML document handling"));
    }
}
