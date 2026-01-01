use crate::pml::types::{PmlChange, PmlChangeType};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PmlComparisonResult {
    pub changes: Vec<PmlChange>,
    pub slides_inserted: i32,
    pub slides_deleted: i32,
    pub shapes_inserted: i32,
    pub shapes_deleted: i32,
    pub shapes_moved: i32,
    pub shapes_resized: i32,
    pub text_changes: i32,
    pub total_changes: i32,
}

impl PmlComparisonResult {
    pub fn new() -> Self {
        Self {
            changes: Vec::new(),
            slides_inserted: 0,
            slides_deleted: 0,
            shapes_inserted: 0,
            shapes_deleted: 0,
            shapes_moved: 0,
            shapes_resized: 0,
            text_changes: 0,
            total_changes: 0,
        }
    }

    pub fn to_json(&self) -> String {
        serde_json::to_string_pretty(self).unwrap_or_default()
    }

    pub fn get_changes_by_slide(&self) -> HashMap<usize, Vec<&PmlChange>> {
        let mut by_slide: HashMap<usize, Vec<&PmlChange>> = HashMap::new();
        for change in &self.changes {
            if let Some(index) = change.slide_index {
                by_slide.entry(index).or_default().push(change);
            }
        }
        by_slide
    }

    pub fn get_changes_by_type(&self) -> HashMap<PmlChangeType, Vec<&PmlChange>> {
        let mut by_type: HashMap<PmlChangeType, Vec<&PmlChange>> = HashMap::new();
        for change in &self.changes {
            by_type
                .entry(change.change_type.clone())
                .or_default()
                .push(change);
        }
        by_type
    }
}

impl Default for PmlComparisonResult {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn result_serializes_to_json() {
        let result = PmlComparisonResult::new();
        let json = result.to_json();
        assert!(json.contains("\"changes\""));
        assert!(json.contains("\"total_changes\""));
    }
}
