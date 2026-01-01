// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

//! SmlComparer - Entry point for Excel workbook comparison.
//!
//! This module provides the public API for comparing two Excel workbooks and producing
//! marked workbooks with highlighted differences. It orchestrates the entire comparison
//! pipeline: canonicalization, diff computation, and markup rendering.
//!
//! 100% parity with C# SmlComparer.cs

use super::canonicalize::SmlCanonicalizer;
use super::diff::compute_diff;
use super::document::SmlDocument;
use super::markup::render_marked_workbook;
use super::result::SmlComparisonResult;
use super::settings::SmlComparerSettings;
use crate::error::Result;

/// Main comparer for Excel workbooks.
/// Provides static methods for comparing workbooks and producing marked output.
pub struct SmlComparer;

impl SmlComparer {
    /// Compare two Excel workbooks and return a detailed change report.
    ///
    /// This method:
    /// 1. Canonicalizes both workbooks into normalized signatures
    /// 2. Computes differences using the diff engine
    /// 3. Returns a structured result with all detected changes
    ///
    /// # Arguments
    ///
    /// * `older` - The original/older workbook
    /// * `newer` - The revised/newer workbook
    /// * `settings` - Optional comparison settings (uses defaults if None)
    ///
    /// # Returns
    ///
    /// A `SmlComparisonResult` containing all detected changes between the workbooks.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use redline_core::sml::{SmlComparer, SmlDocument, SmlComparerSettings};
    ///
    /// let older_bytes = std::fs::read("old.xlsx").unwrap();
    /// let newer_bytes = std::fs::read("new.xlsx").unwrap();
    ///
    /// let older = SmlDocument::from_bytes(&older_bytes).unwrap();
    /// let newer = SmlDocument::from_bytes(&newer_bytes).unwrap();
    ///
    /// let settings = SmlComparerSettings::default();
    /// let result = SmlComparer::compare(&older, &newer, Some(&settings)).unwrap();
    ///
    /// println!("Total changes: {}", result.total_changes());
    /// println!("Value changes: {}", result.value_changes());
    /// ```
    pub fn compare(
        older: &SmlDocument,
        newer: &SmlDocument,
        settings: Option<&SmlComparerSettings>,
    ) -> Result<SmlComparisonResult> {
        let settings = settings.cloned().unwrap_or_default();

        settings.log("SmlComparer.Compare: Starting comparison");

        // Canonicalize both workbooks
        let sig1 = SmlCanonicalizer::canonicalize(older, &settings)?;
        let sig2 = SmlCanonicalizer::canonicalize(newer, &settings)?;

        settings.log(&format!(
            "SmlComparer.Compare: Canonicalized older workbook: {} sheets",
            sig1.sheets.len()
        ));
        settings.log(&format!(
            "SmlComparer.Compare: Canonicalized newer workbook: {} sheets",
            sig2.sheets.len()
        ));

        // Compute diff
        let result = compute_diff(&sig1, &sig2, &settings);

        settings.log(&format!(
            "SmlComparer.Compare: Found {} changes",
            result.total_changes()
        ));

        Ok(result)
    }

    /// Produce a marked workbook highlighting all differences between two workbooks.
    ///
    /// The output is based on the newer workbook with highlights and comments showing changes.
    /// This method:
    /// 1. Compares the workbooks using `compare()`
    /// 2. Applies visual markup to the newer workbook to highlight changes
    /// 3. Returns a new workbook with all changes annotated
    ///
    /// # Arguments
    ///
    /// * `older` - The original/older workbook
    /// * `newer` - The revised/newer workbook
    /// * `settings` - Optional comparison settings (uses defaults if None)
    ///
    /// # Returns
    ///
    /// A new `SmlDocument` based on the newer workbook with changes highlighted.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use redline_core::sml::{SmlComparer, SmlDocument, SmlComparerSettings};
    ///
    /// let older_bytes = std::fs::read("old.xlsx").unwrap();
    /// let newer_bytes = std::fs::read("new.xlsx").unwrap();
    ///
    /// let older = SmlDocument::from_bytes(&older_bytes).unwrap();
    /// let newer = SmlDocument::from_bytes(&newer_bytes).unwrap();
    ///
    /// let settings = SmlComparerSettings::default();
    /// let marked = SmlComparer::produce_marked_workbook(&older, &newer, Some(&settings)).unwrap();
    ///
    /// std::fs::write("marked.xlsx", marked.to_bytes().unwrap()).unwrap();
    /// ```
    pub fn produce_marked_workbook(
        older: &SmlDocument,
        newer: &SmlDocument,
        settings: Option<&SmlComparerSettings>,
    ) -> Result<SmlDocument> {
        let settings = settings.cloned().unwrap_or_default();

        settings.log("SmlComparer.ProduceMarkedWorkbook: Starting");

        // First compute the diff
        let result = Self::compare(older, newer, Some(&settings))?;

        // Then render the marked workbook
        let marked_workbook = render_marked_workbook(newer, &result, &settings)?;

        settings.log("SmlComparer.ProduceMarkedWorkbook: Complete");

        Ok(marked_workbook)
    }

    /// Compare two Excel workbooks and return both the marked workbook and the change list.
    ///
    /// This combines `compare` and `render_marked_workbook` into a single operation,
    /// returning both artifacts needed for a full UI experience.
    pub fn compare_and_render(
        older: &SmlDocument,
        newer: &SmlDocument,
        settings: Option<&SmlComparerSettings>,
    ) -> Result<(SmlDocument, SmlComparisonResult)> {
        let settings = settings.cloned().unwrap_or_default();

        // First compute the diff
        let result = Self::compare(older, newer, Some(&settings))?;

        // Then render the marked workbook
        let marked_workbook = render_marked_workbook(newer, &result, &settings)?;

        Ok((marked_workbook, result))
    }
}
