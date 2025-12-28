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
            by_slide.entry(change.slide_index).or_default().push(change);
        }
        by_slide
    }

    pub fn get_changes_by_type(&self) -> HashMap<PmlChangeType, Vec<&PmlChange>> {
        let mut by_type: HashMap<PmlChangeType, Vec<&PmlChange>> = HashMap::new();
        for change in &self.changes {
            by_type.entry(change.change_type.clone()).or_default().push(change);
        }
        by_type
    }
}

impl Default for PmlComparisonResult {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PmlChange {
    pub slide_index: usize,
    pub shape_id: Option<String>,
    pub shape_name: Option<String>,
    pub change_type: PmlChangeType,
    pub description: Option<String>,
    /// New X coordinate in EMUs (for moved/inserted shapes)
    pub new_x: Option<i64>,
    /// New Y coordinate in EMUs (for moved/inserted shapes)
    pub new_y: Option<i64>,
    /// Old X coordinate in EMUs (for moved shapes)
    pub old_x: Option<i64>,
    /// Old Y coordinate in EMUs (for moved shapes)
    pub old_y: Option<i64>,
}

impl PmlChange {
    /// Create a new PmlChange with minimal fields (others default to None).
    pub fn new(slide_index: usize, change_type: PmlChangeType) -> Self {
        Self {
            slide_index,
            shape_id: None,
            shape_name: None,
            change_type,
            description: None,
            new_x: None,
            new_y: None,
            old_x: None,
            old_y: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum PmlChangeType {
    // Presentation-level
    SlideSizeChanged,
    ThemeChanged,
    
    // Slide-level structure
    SlideInserted,
    SlideDeleted,
    SlideMoved,
    SlideModified,
    SlideLayoutChanged,
    SlideBackgroundChanged,
    SlideNotesChanged,
    
    // Shape-level structure  
    ShapeInserted,
    ShapeDeleted,
    ShapeMoved,
    ShapeResized,
    ShapeRotated,
    ShapeZOrderChanged,
    ShapeModified,
    
    // Shape content
    TextChanged,
    FormattingChanged,
    ImageReplaced,
    TableContentChanged,
    ChartDataChanged,
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
