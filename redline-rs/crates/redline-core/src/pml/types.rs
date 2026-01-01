// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

//! Public types for PmlComparer results and changes.
//!
//! This module defines PmlChange and PmlChangeType which are used to represent
//! individual changes detected during presentation comparison.

use serde::{Deserialize, Serialize};

/// Types of changes detected during presentation comparison.
/// 100% parity with C# PmlChangeType enum.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub enum PmlChangeType {
    // Presentation level
    SlideSizeChanged,
    ThemeChanged,

    // Slide level
    SlideInserted,
    SlideDeleted,
    SlideMoved,
    SlideLayoutChanged,
    SlideBackgroundChanged,
    SlideTransitionChanged,
    SlideNotesChanged,

    // Shape level
    ShapeInserted,
    ShapeDeleted,
    ShapeMoved,
    ShapeResized,
    ShapeRotated,
    ShapeZOrderChanged,
    ShapeTypeChanged,

    // Content level
    TextChanged,
    TextFormattingChanged,
    ImageReplaced,
    TableContentChanged,
    TableStructureChanged,
    ChartDataChanged,
    ChartFormatChanged,

    // Shape style
    ShapeFillChanged,
    ShapeLineChanged,
    ShapeEffectsChanged,
    GroupMembershipChanged,
}

/// Type of text change.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub enum TextChangeType {
    Insert,
    Delete,
    Replace,
    FormatOnly,
}

/// Word count statistics for a PML change.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PmlWordCount {
    pub deleted: usize,
    pub inserted: usize,
}

/// Detailed text change information.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PmlTextChange {
    pub r#type: TextChangeType,
    pub paragraph_index: usize,
    pub run_index: usize,
    pub old_text: Option<String>,
    pub new_text: Option<String>,
}

/// Represents a single change between two presentations.
/// 100% parity with C# PmlChange class.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PmlChange {
    pub change_type: PmlChangeType,
    pub slide_index: Option<usize>,
    pub old_slide_index: Option<usize>,
    pub shape_name: Option<String>,
    pub shape_id: Option<String>,
    pub old_value: Option<String>,
    pub new_value: Option<String>,

    // Transform changes
    pub old_x: Option<f64>,
    pub old_y: Option<f64>,
    pub old_cx: Option<f64>,
    pub old_cy: Option<f64>,
    pub new_x: Option<f64>,
    pub new_y: Option<f64>,
    pub new_cx: Option<f64>,
    pub new_cy: Option<f64>,

    pub text_changes: Option<Vec<PmlTextChange>>,
    pub match_confidence: Option<f64>,
}

impl Default for PmlChange {
    fn default() -> Self {
        Self {
            change_type: PmlChangeType::SlideInserted,
            slide_index: None,
            old_slide_index: None,
            shape_name: None,
            shape_id: None,
            old_value: None,
            new_value: None,
            old_x: None,
            old_y: None,
            old_cx: None,
            old_cy: None,
            new_x: None,
            new_y: None,
            new_cx: None,
            new_cy: None,
            text_changes: None,
            match_confidence: None,
        }
    }
}

impl PmlChange {
    pub fn get_description(&self) -> String {
        match self.change_type {
            PmlChangeType::SlideInserted => {
                format!("Slide {} inserted", self.slide_index.unwrap_or(0))
            }
            PmlChangeType::SlideDeleted => {
                format!("Slide {} deleted", self.old_slide_index.unwrap_or(0))
            }
            PmlChangeType::SlideMoved => {
                format!(
                    "Slide moved from position {} to {}",
                    self.old_slide_index.unwrap_or(0),
                    self.slide_index.unwrap_or(0)
                )
            }
            PmlChangeType::SlideLayoutChanged => {
                format!("Slide {} layout changed", self.slide_index.unwrap_or(0))
            }
            PmlChangeType::SlideBackgroundChanged => {
                format!("Slide {} background changed", self.slide_index.unwrap_or(0))
            }
            PmlChangeType::SlideNotesChanged => {
                format!("Slide {} notes changed", self.slide_index.unwrap_or(0))
            }
            PmlChangeType::ShapeInserted => {
                format!(
                    "Shape '{}' inserted on slide {}",
                    self.shape_name.as_deref().unwrap_or(""),
                    self.slide_index.unwrap_or(0)
                )
            }
            PmlChangeType::ShapeDeleted => {
                format!(
                    "Shape '{}' deleted from slide {}",
                    self.shape_name.as_deref().unwrap_or(""),
                    self.slide_index.unwrap_or(0)
                )
            }
            PmlChangeType::ShapeMoved => {
                format!(
                    "Shape '{}' moved on slide {}",
                    self.shape_name.as_deref().unwrap_or(""),
                    self.slide_index.unwrap_or(0)
                )
            }
            PmlChangeType::ShapeResized => {
                format!(
                    "Shape '{}' resized on slide {}",
                    self.shape_name.as_deref().unwrap_or(""),
                    self.slide_index.unwrap_or(0)
                )
            }
            PmlChangeType::ShapeRotated => {
                format!(
                    "Shape '{}' rotated on slide {}",
                    self.shape_name.as_deref().unwrap_or(""),
                    self.slide_index.unwrap_or(0)
                )
            }
            PmlChangeType::ShapeZOrderChanged => {
                format!(
                    "Shape '{}' z-order changed on slide {}",
                    self.shape_name.as_deref().unwrap_or(""),
                    self.slide_index.unwrap_or(0)
                )
            }
            PmlChangeType::TextChanged => {
                format!(
                    "Text changed in '{}' on slide {}",
                    self.shape_name.as_deref().unwrap_or(""),
                    self.slide_index.unwrap_or(0)
                )
            }
            PmlChangeType::TextFormattingChanged => {
                format!(
                    "Text formatting changed in '{}' on slide {}",
                    self.shape_name.as_deref().unwrap_or(""),
                    self.slide_index.unwrap_or(0)
                )
            }
            PmlChangeType::ImageReplaced => {
                format!(
                    "Image replaced in '{}' on slide {}",
                    self.shape_name.as_deref().unwrap_or(""),
                    self.slide_index.unwrap_or(0)
                )
            }
            PmlChangeType::TableContentChanged => {
                format!(
                    "Table content changed in '{}' on slide {}",
                    self.shape_name.as_deref().unwrap_or(""),
                    self.slide_index.unwrap_or(0)
                )
            }
            PmlChangeType::ChartDataChanged => {
                format!(
                    "Chart data changed in '{}' on slide {}",
                    self.shape_name.as_deref().unwrap_or(""),
                    self.slide_index.unwrap_or(0)
                )
            }
            _ => format!(
                "{:?} on slide {}",
                self.change_type,
                self.slide_index.unwrap_or(0)
            ),
        }
    }
}

/// UI-friendly representation of a change for display in a change list.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PmlChangeListItem {
    pub id: String,
    pub change_type: PmlChangeType,
    pub slide_index: Option<usize>,
    pub shape_name: Option<String>,
    pub shape_id: Option<String>,
    pub summary: String,
    pub preview_text: Option<String>,
    pub word_count: Option<PmlWordCount>,
    pub count: Option<usize>,
    pub details: Option<PmlChangeDetails>,
    pub anchor: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PmlChangeDetails {
    pub old_value: Option<String>,
    pub new_value: Option<String>,
    pub old_slide_index: Option<usize>,
    pub text_changes: Option<Vec<PmlTextChange>>,
    pub match_confidence: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PmlChangeListOptions {
    pub group_by_slide: bool,
    pub max_preview_length: usize,
}

impl Default for PmlChangeListOptions {
    fn default() -> Self {
        Self {
            group_by_slide: true,
            max_preview_length: 100,
        }
    }
}
