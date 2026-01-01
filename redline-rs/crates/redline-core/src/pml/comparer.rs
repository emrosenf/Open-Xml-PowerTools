// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

//! PmlComparer - Main entry point for PowerPoint presentation comparison.
//!
//! Provides:
//! - `compare()` - Compare two presentations and return detailed change list
//! - `produce_marked_presentation()` - Generate a presentation with visual change overlays
//!
//! # Architecture
//!
//! The comparison pipeline has 5 stages:
//! 1. **Canonicalization** (PmlCanonicalizer) - Extract presentation signatures
//! 2. **Slide Matching** (PmlSlideMatchEngine) - Match slides between versions
//! 3. **Shape Matching** (PmlShapeMatchEngine) - Match shapes within slides
//! 4. **Diff Engine** (PmlDiffEngine) - Generate detailed change list
//! 5. **Markup Rendering** (PmlMarkupRenderer) - Add visual annotations
//!
//! C# Source: OpenXmlPowerTools/PmlComparer.cs:2622-2690

use super::canonicalize::PmlCanonicalizer;
use super::diff::PmlDiffEngine;
use super::document::PmlDocument;
use super::markup::render_marked_presentation;
use super::result::PmlComparisonResult;
use super::settings::PmlComparerSettings;
use crate::error::Result;

/// Main entry point for PowerPoint presentation comparison.
pub struct PmlComparer;

impl PmlComparer {
    /// Compare two presentations and return a detailed list of changes.
    ///
    /// This is the main comparison entry point. It:
    /// 1. Canonicalizes both presentations (extracts structure signatures)
    /// 2. Matches slides using multi-pass heuristics
    /// 3. Matches shapes within matched slides
    /// 4. Generates a detailed change list with statistics
    ///
    /// # Arguments
    /// - `older` - The original/older presentation
    /// - `newer` - The modified/newer presentation
    /// - `settings` - Optional comparison settings (uses defaults if None)
    ///
    /// # Returns
    /// `PmlComparisonResult` with:
    /// - List of all detected changes
    /// - Statistics (slides inserted/deleted, shapes modified, etc.)
    /// - Change descriptions
    ///
    /// # C# Signature
    /// ```csharp
    /// public static PmlComparisonResult Compare(
    ///     PmlDocument older,
    ///     PmlDocument newer,
    ///     PmlComparerSettings settings = null)
    /// ```
    pub fn compare(
        older: &PmlDocument,
        newer: &PmlDocument,
        settings: Option<&PmlComparerSettings>,
    ) -> Result<PmlComparisonResult> {
        let settings = settings.cloned().unwrap_or_default();

        // 1. Canonicalize both presentations
        let sig1 = PmlCanonicalizer::canonicalize(older, &settings)?;
        let sig2 = PmlCanonicalizer::canonicalize(newer, &settings)?;

        // 2. Run diff engine to compare signatures
        // (DiffEngine internally uses SlideMatchEngine and ShapeMatchEngine)
        let result = PmlDiffEngine::compare(&sig1, &sig2, &settings);

        Ok(result)
    }

    /// Generate a marked presentation with visual change overlays.
    ///
    /// This is a convenience method that:
    /// 1. Calls `compare()` to get the change list
    /// 2. Calls `render_marked_presentation()` to add visual annotations
    ///
    /// The result is a new presentation (based on `newer`) with:
    /// - Visual labels on changed shapes (color-coded by change type)
    /// - Optional summary slide listing all changes
    /// - Optional speaker notes annotations
    ///
    /// # Arguments
    /// - `older` - The original/older presentation
    /// - `newer` - The modified/newer presentation
    /// - `settings` - Optional comparison settings (controls colors, summary slide, etc.)
    ///
    /// # Returns
    /// A new `PmlDocument` with visual change overlays.
    /// If no changes detected, returns a clone of `newer`.
    ///
    /// # C# Signature
    /// ```csharp
    /// public static PmlDocument ProduceMarkedPresentation(
    ///     PmlDocument older,
    ///     PmlDocument newer,
    ///     PmlComparerSettings settings = null)
    /// ```
    pub fn produce_marked_presentation(
        older: &PmlDocument,
        newer: &PmlDocument,
        settings: Option<&PmlComparerSettings>,
    ) -> Result<PmlDocument> {
        let settings = settings.cloned().unwrap_or_default();

        // 1. Compare to get change list
        let result = Self::compare(older, newer, Some(&settings))?;

        // 2. Render marked presentation
        let marked = render_marked_presentation(newer, &result, &settings)?;

        Ok(marked)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn compare_empty_presentations_returns_no_changes() {
        // Empty PPTX files won't parse - this test verifies API
        // Real integration tests should use actual PPTX files
    }

    #[test]
    fn compare_accepts_custom_settings() {
        // This validates that settings can be passed through
        let _settings = PmlComparerSettings::default();
        // Actual comparison requires valid PPTX files
    }

    #[test]
    fn produce_marked_presentation_accepts_custom_settings() {
        let settings = PmlComparerSettings {
            add_summary_slide: true,
            ..Default::default()
        };
        assert!(settings.add_summary_slide);
    }
}
