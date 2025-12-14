// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlPowerTools;
using Xunit;

// Aliases to resolve ambiguous references
using SpreadsheetCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using SpreadsheetRow = DocumentFormat.OpenXml.Spreadsheet.Row;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class SmlComparerTests
    {
        /// <summary>
        /// Helper to create a simple test workbook with specified cell values.
        /// </summary>
        private static SmlDocument CreateTestWorkbook(Dictionary<string, Dictionary<string, object>> sheetData)
        {
            using var ms = new MemoryStream();
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Add shared string table
                var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringPart.SharedStringTable = new SharedStringTable();

                // Add styles part
                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = CreateDefaultStylesheet();

                var sheets = new Sheets();
                uint sheetId = 1;

                foreach (var sheet in sheetData)
                {
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetDataElement = new SheetData();

                    foreach (var cell in sheet.Value)
                    {
                        var cellRef = cell.Key;
                        var (col, row) = ParseCellRef(cellRef);

                        // Find or create row
                        var rowElement = sheetDataElement.Elements<SpreadsheetRow>()
                            .FirstOrDefault(r => r.RowIndex == (uint)row);
                        if (rowElement == null)
                        {
                            rowElement = new SpreadsheetRow { RowIndex = (uint)row };
                            sheetDataElement.Append(rowElement);
                        }

                        // Create cell
                        var cellElement = new SpreadsheetCell { CellReference = cellRef };

                        if (cell.Value is string strValue)
                        {
                            // Add to shared string table
                            var ssIndex = AddSharedString(sharedStringPart.SharedStringTable, strValue);
                            cellElement.DataType = CellValues.SharedString;
                            cellElement.CellValue = new CellValue(ssIndex.ToString());
                        }
                        else if (cell.Value is double dblValue)
                        {
                            cellElement.CellValue = new CellValue(dblValue.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        }
                        else if (cell.Value is int intValue)
                        {
                            cellElement.CellValue = new CellValue(intValue.ToString());
                        }
                        else if (cell.Value is CellWithFormula cwf)
                        {
                            cellElement.CellFormula = new CellFormula(cwf.Formula);
                            if (cwf.Value != null)
                                cellElement.CellValue = new CellValue(cwf.Value);
                        }

                        rowElement.Append(cellElement);
                    }

                    worksheetPart.Worksheet = new Worksheet(sheetDataElement);

                    sheets.Append(new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = sheetId++,
                        Name = sheet.Key
                    });
                }

                workbookPart.Workbook.Append(sheets);
            }

            return new SmlDocument("test.xlsx", ms.ToArray());
        }

        private static Stylesheet CreateDefaultStylesheet()
        {
            return new Stylesheet(
                new Fonts(
                    new Font(
                        new FontSize { Val = 11 },
                        new FontName { Val = "Calibri" }
                    )
                ) { Count = 1 },
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
                ) { Count = 2 },
                new Borders(
                    new Border(
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()
                    )
                ) { Count = 1 },
                new CellFormats(
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 0 }
                ) { Count = 1 }
            );
        }

        private static int AddSharedString(SharedStringTable sst, string text)
        {
            var items = sst.Elements<SharedStringItem>().ToList();
            for (int i = 0; i < items.Count; i++)
            {
                if (items[i].InnerText == text)
                    return i;
            }
            sst.Append(new SharedStringItem(new Text(text)));
            sst.Count = (uint)sst.Elements<SharedStringItem>().Count();
            sst.UniqueCount = sst.Count;
            return (int)(uint)sst.Count - 1;
        }

        private static (int col, int row) ParseCellRef(string cellRef)
        {
            var col = 0;
            var i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i]))
            {
                col = col * 26 + (char.ToUpper(cellRef[i]) - 'A' + 1);
                i++;
            }
            int.TryParse(cellRef.Substring(i), out var row);
            return (col, row);
        }

        // Helper class for cells with formulas
        private class CellWithFormula
        {
            public string Formula { get; set; }
            public string Value { get; set; }
        }

        #region Basic Comparison Tests

        [Fact]
        public void SC001_IdenticalWorkbooks_NoChanges()
        {
            // Arrange
            var data = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello",
                    ["B1"] = 123.45,
                    ["A2"] = "World"
                }
            };

            var doc1 = CreateTestWorkbook(data);
            var doc2 = CreateTestWorkbook(data);
            var settings = new SmlComparerSettings();

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.TotalChanges);
        }

        [Fact]
        public void SC002_SingleCellValueChange_DetectedCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello",
                    ["B1"] = 123.45
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Goodbye",  // Changed
                    ["B1"] = 123.45
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings();

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Equal(1, result.TotalChanges);
            Assert.Equal(1, result.ValueChanges);

            var change = result.Changes.First();
            Assert.Equal(SmlChangeType.ValueChanged, change.ChangeType);
            Assert.Equal("Sheet1", change.SheetName);
            Assert.Equal("A1", change.CellAddress);
            Assert.Equal("Hello", change.OldValue);
            Assert.Equal("Goodbye", change.NewValue);
        }

        [Fact]
        public void SC003_CellAdded_DetectedCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello",
                    ["B1"] = "World"  // Added
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings();

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Equal(1, result.TotalChanges);
            Assert.Equal(1, result.CellsAdded);

            var change = result.Changes.First();
            Assert.Equal(SmlChangeType.CellAdded, change.ChangeType);
            Assert.Equal("B1", change.CellAddress);
            Assert.Equal("World", change.NewValue);
        }

        [Fact]
        public void SC004_CellDeleted_DetectedCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello",
                    ["B1"] = "World"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello"
                    // B1 deleted
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings();

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Equal(1, result.TotalChanges);
            Assert.Equal(1, result.CellsDeleted);

            var change = result.Changes.First();
            Assert.Equal(SmlChangeType.CellDeleted, change.ChangeType);
            Assert.Equal("B1", change.CellAddress);
            Assert.Equal("World", change.OldValue);
        }

        [Fact]
        public void SC005_SheetAdded_DetectedCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello"
                },
                ["Sheet2"] = new Dictionary<string, object>  // Added
                {
                    ["A1"] = "New Sheet"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings();

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Equal(1, result.SheetsAdded);
            var sheetChange = result.Changes.First(c => c.ChangeType == SmlChangeType.SheetAdded);
            Assert.Equal("Sheet2", sheetChange.SheetName);
        }

        [Fact]
        public void SC006_SheetDeleted_DetectedCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello"
                },
                ["Sheet2"] = new Dictionary<string, object>
                {
                    ["A1"] = "Will be deleted"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello"
                }
                // Sheet2 deleted
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings();

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Equal(1, result.SheetsDeleted);
            var sheetChange = result.Changes.First(c => c.ChangeType == SmlChangeType.SheetDeleted);
            Assert.Equal("Sheet2", sheetChange.SheetName);
        }

        #endregion

        #region Formula Tests

        [Fact]
        public void SC007_FormulaChange_DetectedCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = 10,
                    ["A2"] = 20,
                    ["A3"] = new CellWithFormula { Formula = "A1+A2", Value = "30" }
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = 10,
                    ["A2"] = 20,
                    ["A3"] = new CellWithFormula { Formula = "A1*A2", Value = "200" }  // Formula changed
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings();

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - Value changed because formula result is different
            Assert.True(result.TotalChanges >= 1);
        }

        #endregion

        #region Settings Tests

        [Fact]
        public void SC008_CaseInsensitiveComparison()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "HELLO"  // Same content, different case
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);

            // Act - Case sensitive (default)
            var resultSensitive = SmlComparer.Compare(doc1, doc2, new SmlComparerSettings { CaseInsensitiveValues = false });

            // Act - Case insensitive
            var resultInsensitive = SmlComparer.Compare(doc1, doc2, new SmlComparerSettings { CaseInsensitiveValues = true });

            // Assert
            Assert.Equal(1, resultSensitive.ValueChanges);  // Should detect change
            Assert.Equal(0, resultInsensitive.ValueChanges);  // Should not detect change
        }

        [Fact]
        public void SC009_NumericTolerance()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = 100.0
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = 100.001  // Slightly different
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);

            // Act - No tolerance
            var resultNoTolerance = SmlComparer.Compare(doc1, doc2, new SmlComparerSettings { NumericTolerance = 0 });

            // Act - With tolerance
            var resultWithTolerance = SmlComparer.Compare(doc1, doc2, new SmlComparerSettings { NumericTolerance = 0.01 });

            // Assert
            Assert.Equal(1, resultNoTolerance.ValueChanges);  // Should detect change
            Assert.Equal(0, resultWithTolerance.ValueChanges);  // Should not detect change
        }

        [Fact]
        public void SC010_DisableFormattingComparison()
        {
            // This test would require creating workbooks with different formatting
            // For now, just verify the setting is respected
            var settings = new SmlComparerSettings { CompareFormatting = false };
            Assert.False(settings.CompareFormatting);
        }

        #endregion

        #region Output Tests

        [Fact]
        public void SC011_ProduceMarkedWorkbook_CreatesValidOutput()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Original"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Modified"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings();

            // Act
            var markedDoc = SmlComparer.ProduceMarkedWorkbook(doc1, doc2, settings);

            // Assert
            Assert.NotNull(markedDoc);
            Assert.NotNull(markedDoc.DocumentByteArray);
            Assert.True(markedDoc.DocumentByteArray.Length > 0);

            // Verify the document can be opened
            using var ms = new MemoryStream(markedDoc.DocumentByteArray);
            using var sDoc = SpreadsheetDocument.Open(ms, false);
            Assert.NotNull(sDoc.WorkbookPart);

            // Verify _DiffSummary sheet exists
            var sheets = sDoc.WorkbookPart.Workbook.Sheets.Cast<Sheet>().ToList();
            Assert.Contains(sheets, s => s.Name == "_DiffSummary");
        }

        [Fact]
        public void SC012_ComparisonResult_ToJson()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Hello"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "World"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);

            // Act
            var result = SmlComparer.Compare(doc1, doc2, new SmlComparerSettings());
            var json = result.ToJson();

            // Assert
            Assert.NotNull(json);
            Assert.Contains("TotalChanges", json);
            Assert.Contains("ValueChanges", json);
            Assert.Contains("Changes", json);
        }

        #endregion

        #region Statistics Tests

        [Fact]
        public void SC013_Statistics_CorrectlySummarized()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Value1",
                    ["A2"] = "Value2",
                    ["A3"] = "Value3"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Changed1",  // Value change
                    ["A2"] = "Value2",    // No change
                    // A3 deleted
                    ["A4"] = "New"        // Added
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);

            // Act
            var result = SmlComparer.Compare(doc1, doc2, new SmlComparerSettings());

            // Assert
            Assert.Equal(1, result.ValueChanges);
            Assert.Equal(1, result.CellsAdded);
            Assert.Equal(1, result.CellsDeleted);
            Assert.Equal(3, result.TotalChanges);
            Assert.Equal(2, result.StructuralChanges);  // Added + Deleted
        }

        #endregion

        #region CellFormatSignature Tests

        [Fact]
        public void SC014_CellFormatSignature_Equality()
        {
            var format1 = new CellFormatSignature
            {
                Bold = true,
                FontSize = 12,
                FontName = "Arial"
            };

            var format2 = new CellFormatSignature
            {
                Bold = true,
                FontSize = 12,
                FontName = "Arial"
            };

            var format3 = new CellFormatSignature
            {
                Bold = false,
                FontSize = 12,
                FontName = "Arial"
            };

            Assert.True(format1.Equals(format2));
            Assert.False(format1.Equals(format3));
        }

        [Fact]
        public void SC015_CellFormatSignature_GetDifferenceDescription()
        {
            var format1 = new CellFormatSignature
            {
                Bold = false,
                FontSize = 11,
                FontName = "Calibri"
            };

            var format2 = new CellFormatSignature
            {
                Bold = true,
                FontSize = 14,
                FontName = "Arial"
            };

            var description = format2.GetDifferenceDescription(format1);

            Assert.Contains("Made bold", description);
            Assert.Contains("Font", description);
            Assert.Contains("Size", description);
        }

        #endregion

        #region Edge Cases

        [Fact]
        public void SC016_EmptyWorkbooks_NoChanges()
        {
            // Arrange
            var data = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>()
            };

            var doc1 = CreateTestWorkbook(data);
            var doc2 = CreateTestWorkbook(data);

            // Act
            var result = SmlComparer.Compare(doc1, doc2, new SmlComparerSettings());

            // Assert
            Assert.Equal(0, result.TotalChanges);
        }

        [Fact]
        public void SC017_MultipleSheets_ComparedCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object> { ["A1"] = "S1" },
                ["Sheet2"] = new Dictionary<string, object> { ["A1"] = "S2" },
                ["Sheet3"] = new Dictionary<string, object> { ["A1"] = "S3" }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object> { ["A1"] = "S1-changed" },  // Changed
                ["Sheet2"] = new Dictionary<string, object> { ["A1"] = "S2" },          // Same
                ["Sheet3"] = new Dictionary<string, object> { ["A1"] = "S3-changed" }   // Changed
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);

            // Act
            var result = SmlComparer.Compare(doc1, doc2, new SmlComparerSettings());

            // Assert
            Assert.Equal(2, result.ValueChanges);
            Assert.Contains(result.Changes, c => c.SheetName == "Sheet1");
            Assert.Contains(result.Changes, c => c.SheetName == "Sheet3");
        }

        [Fact]
        public void SC018_NumericValues_ComparedAsNumbers()
        {
            // Arrange - Same numeric value, different string representation
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = 100.0
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = 100  // int vs double, same value
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);

            // Act
            var result = SmlComparer.Compare(doc1, doc2, new SmlComparerSettings());

            // Assert - Should be considered the same after normalization
            Assert.Equal(0, result.ValueChanges);
        }

        #endregion

        #region Integration Tests

        [Fact]
        public void SC019_RoundTrip_MarkedWorkbookCanBeComparedAgain()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Original"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Modified"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings();

            // Act - Create marked workbook
            var markedDoc = SmlComparer.ProduceMarkedWorkbook(doc1, doc2, settings);

            // Act - Compare marked workbook with itself (should have no changes)
            var result = SmlComparer.Compare(markedDoc, markedDoc, settings);

            // Assert
            Assert.Equal(0, result.TotalChanges);
        }

        [Fact]
        public void SC020_GetChangesBySheet_FiltersCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object> { ["A1"] = "S1" },
                ["Sheet2"] = new Dictionary<string, object> { ["A1"] = "S2" }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object> { ["A1"] = "S1-changed" },
                ["Sheet2"] = new Dictionary<string, object> { ["A1"] = "S2-changed" }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);

            // Act
            var result = SmlComparer.Compare(doc1, doc2, new SmlComparerSettings());

            // Assert
            var sheet1Changes = result.GetChangesBySheet("Sheet1").ToList();
            var sheet2Changes = result.GetChangesBySheet("Sheet2").ToList();

            Assert.Single(sheet1Changes);
            Assert.Single(sheet2Changes);
            Assert.All(sheet1Changes, c => Assert.Equal("Sheet1", c.SheetName));
            Assert.All(sheet2Changes, c => Assert.Equal("Sheet2", c.SheetName));
        }

        #endregion

        #region Phase 2: Row Alignment Tests

        [Fact]
        public void SC021_RowInserted_DetectedCorrectly()
        {
            // Arrange - Original has 3 rows, new has 4 rows (row inserted in middle)
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Header",
                    ["A2"] = "Row2Data",
                    ["A3"] = "Row3Data"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Header",
                    ["A2"] = "InsertedRow",  // New row inserted
                    ["A3"] = "Row2Data",     // Original row 2 moved to row 3
                    ["A4"] = "Row3Data"      // Original row 3 moved to row 4
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings { EnableRowAlignment = true };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - With row alignment, should detect the insertion
            Assert.True(result.RowsInserted >= 1);
        }

        [Fact]
        public void SC022_RowDeleted_DetectedCorrectly()
        {
            // Arrange - Original has 4 rows, new has 3 rows (middle row deleted)
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Header",
                    ["A2"] = "Row2ToDelete",
                    ["A3"] = "Row3Data",
                    ["A4"] = "Row4Data"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Header",
                    ["A2"] = "Row3Data",  // Original row 3 moved up
                    ["A3"] = "Row4Data"   // Original row 4 moved up
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings { EnableRowAlignment = true };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - With row alignment, should detect the deletion
            Assert.True(result.RowsDeleted >= 1);
        }

        [Fact]
        public void SC023_RowAlignment_DisabledFallsBackToCellByCell()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Header",
                    ["A2"] = "OriginalRow2"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Header",
                    ["A2"] = "NewRow",
                    ["A3"] = "OriginalRow2"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);

            // Act with row alignment disabled
            var settings = new SmlComparerSettings { EnableRowAlignment = false };
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - Without alignment, it just sees cell changes
            Assert.Equal(0, result.RowsInserted);
            Assert.True(result.ValueChanges > 0 || result.CellsAdded > 0);
        }

        [Fact]
        public void SC024_MultipleRowChanges_DetectedCorrectly()
        {
            // Arrange - Multiple rows inserted and deleted
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Row1",
                    ["A2"] = "Row2ToDelete",
                    ["A3"] = "Row3",
                    ["A4"] = "Row4ToDelete",
                    ["A5"] = "Row5"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Row1",
                    ["A2"] = "InsertedRowA",
                    ["A3"] = "Row3",
                    ["A4"] = "Row5",
                    ["A5"] = "InsertedRowB"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings { EnableRowAlignment = true };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - Should detect both insertions and deletions
            Assert.True(result.TotalChanges > 0);
        }

        #endregion

        #region Phase 2: Sheet Rename Detection Tests

        [Fact]
        public void SC025_SheetRenamed_DetectedCorrectly()
        {
            // Arrange - Sheet with same content but different name
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["OldSheetName"] = new Dictionary<string, object>
                {
                    ["A1"] = "Data1",
                    ["A2"] = "Data2",
                    ["A3"] = "Data3"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["NewSheetName"] = new Dictionary<string, object>
                {
                    ["A1"] = "Data1",
                    ["A2"] = "Data2",
                    ["A3"] = "Data3"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings { EnableSheetRenameDetection = true };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Equal(1, result.SheetsRenamed);
            var renameChange = result.Changes.First(c => c.ChangeType == SmlChangeType.SheetRenamed);
            Assert.Equal("OldSheetName", renameChange.OldSheetName);
            Assert.Equal("NewSheetName", renameChange.SheetName);
        }

        [Fact]
        public void SC026_SheetRenameDetection_Disabled()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["OldName"] = new Dictionary<string, object>
                {
                    ["A1"] = "Data"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["NewName"] = new Dictionary<string, object>
                {
                    ["A1"] = "Data"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings { EnableSheetRenameDetection = false };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - Without rename detection, should see as add + delete
            Assert.Equal(0, result.SheetsRenamed);
            Assert.Equal(1, result.SheetsAdded);
            Assert.Equal(1, result.SheetsDeleted);
        }

        [Fact]
        public void SC027_SheetRenamed_BelowSimilarityThreshold_TreatedAsAddDelete()
        {
            // Arrange - Sheet with different content (different sheet)
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["OldSheet"] = new Dictionary<string, object>
                {
                    ["A1"] = "OldData1",
                    ["A2"] = "OldData2",
                    ["A3"] = "OldData3"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["NewSheet"] = new Dictionary<string, object>
                {
                    ["A1"] = "CompletelyDifferent1",
                    ["A2"] = "CompletelyDifferent2",
                    ["A3"] = "CompletelyDifferent3"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings
            {
                EnableSheetRenameDetection = true,
                SheetRenameSimilarityThreshold = 0.7
            };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - Different content, so should be add + delete, not rename
            Assert.Equal(0, result.SheetsRenamed);
            Assert.Equal(1, result.SheetsAdded);
            Assert.Equal(1, result.SheetsDeleted);
        }

        [Fact]
        public void SC028_SheetRenamed_PartialContentMatch()
        {
            // Arrange - Sheet with partially matching content
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["OldSheet"] = new Dictionary<string, object>
                {
                    ["A1"] = "Same1",
                    ["A2"] = "Same2",
                    ["A3"] = "Same3",
                    ["A4"] = "Same4",
                    ["A5"] = "Old5"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["NewSheet"] = new Dictionary<string, object>
                {
                    ["A1"] = "Same1",
                    ["A2"] = "Same2",
                    ["A3"] = "Same3",
                    ["A4"] = "Same4",
                    ["A5"] = "New5"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings
            {
                EnableSheetRenameDetection = true,
                SheetRenameSimilarityThreshold = 0.7  // 80% similar should be > 70% threshold
            };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - Should detect as rename since content is mostly the same
            Assert.Equal(1, result.SheetsRenamed);
        }

        #endregion

        #region Phase 2: Settings Tests

        [Fact]
        public void SC029_RowSignatureSampleSize_Setting()
        {
            // Verify the setting can be configured
            var settings = new SmlComparerSettings
            {
                RowSignatureSampleSize = 5
            };
            Assert.Equal(5, settings.RowSignatureSampleSize);

            settings.RowSignatureSampleSize = 20;
            Assert.Equal(20, settings.RowSignatureSampleSize);
        }

        [Fact]
        public void SC030_Phase2Statistics_InJson()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Row1",
                    ["A2"] = "Row2"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Row1",
                    ["A2"] = "NewRow",
                    ["A3"] = "Row2"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings { EnableRowAlignment = true };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);
            var json = result.ToJson();

            // Assert - JSON should include Phase 2 statistics
            Assert.Contains("SheetsRenamed", json);
            Assert.Contains("RowsInserted", json);
            Assert.Contains("RowsDeleted", json);
            Assert.Contains("ColumnsInserted", json);
            Assert.Contains("ColumnsDeleted", json);
        }

        #endregion

        #region Phase 2: Change Description Tests

        [Fact]
        public void SC031_RowInserted_GetDescription()
        {
            var change = new SmlChange
            {
                ChangeType = SmlChangeType.RowInserted,
                SheetName = "Sheet1",
                RowIndex = 5
            };

            var description = change.GetDescription();

            Assert.Contains("Row 5", description);
            Assert.Contains("inserted", description);
            Assert.Contains("Sheet1", description);
        }

        [Fact]
        public void SC032_RowDeleted_GetDescription()
        {
            var change = new SmlChange
            {
                ChangeType = SmlChangeType.RowDeleted,
                SheetName = "Sheet1",
                RowIndex = 3
            };

            var description = change.GetDescription();

            Assert.Contains("Row 3", description);
            Assert.Contains("deleted", description);
        }

        [Fact]
        public void SC033_SheetRenamed_GetDescription()
        {
            var change = new SmlChange
            {
                ChangeType = SmlChangeType.SheetRenamed,
                SheetName = "NewName",
                OldSheetName = "OldName"
            };

            var description = change.GetDescription();

            Assert.Contains("OldName", description);
            Assert.Contains("NewName", description);
            Assert.Contains("renamed", description);
        }

        [Fact]
        public void SC034_ColumnInserted_GetDescription()
        {
            var change = new SmlChange
            {
                ChangeType = SmlChangeType.ColumnInserted,
                SheetName = "Sheet1",
                ColumnIndex = 3  // Column C
            };

            var description = change.GetDescription();

            Assert.Contains("Column C", description);
            Assert.Contains("inserted", description);
        }

        [Fact]
        public void SC035_ColumnDeleted_GetDescription()
        {
            var change = new SmlChange
            {
                ChangeType = SmlChangeType.ColumnDeleted,
                SheetName = "Sheet1",
                ColumnIndex = 26  // Column Z
            };

            var description = change.GetDescription();

            Assert.Contains("Column Z", description);
            Assert.Contains("deleted", description);
        }

        #endregion

        #region Phase 2: Edge Cases

        [Fact]
        public void SC036_AllRowsDeleted_HandledCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Row1",
                    ["A2"] = "Row2",
                    ["A3"] = "Row3"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>()  // Empty sheet
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings { EnableRowAlignment = true };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - Should detect all cells/rows as deleted
            Assert.True(result.RowsDeleted >= 1 || result.CellsDeleted >= 3);
        }

        [Fact]
        public void SC037_AllRowsInserted_HandledCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>()  // Empty sheet
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Row1",
                    ["A2"] = "Row2",
                    ["A3"] = "Row3"
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings { EnableRowAlignment = true };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - Should detect all cells/rows as inserted
            Assert.True(result.RowsInserted >= 1 || result.CellsAdded >= 3);
        }

        [Fact]
        public void SC038_WideSpreadsheet_RowSignatureSampling()
        {
            // Arrange - Create a wide spreadsheet to test sampling
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>()
            };
            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>()
            };

            // Add many columns to test sampling (more than RowSignatureSampleSize)
            for (int i = 0; i < 50; i++)
            {
                var colLetter = GetColumnLetter(i + 1);
                data1["Sheet1"][$"{colLetter}1"] = $"Val{i}";
                data2["Sheet1"][$"{colLetter}1"] = $"Val{i}";
            }

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings
            {
                EnableRowAlignment = true,
                RowSignatureSampleSize = 10  // Sample only 10 of the 50 columns
            };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - Should work without errors
            Assert.Equal(0, result.TotalChanges);
        }

        private static string GetColumnLetter(int columnNumber)
        {
            var result = "";
            while (columnNumber > 0)
            {
                columnNumber--;
                result = (char)('A' + columnNumber % 26) + result;
                columnNumber /= 26;
            }
            return result;
        }

        [Fact]
        public void SC039_MultipleSheetRenames_DetectedCorrectly()
        {
            // Arrange - Multiple sheets renamed
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["OldSheet1"] = new Dictionary<string, object> { ["A1"] = "Data1" },
                ["OldSheet2"] = new Dictionary<string, object> { ["A1"] = "Data2" },
                ["Unchanged"] = new Dictionary<string, object> { ["A1"] = "Same" }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["NewSheet1"] = new Dictionary<string, object> { ["A1"] = "Data1" },
                ["NewSheet2"] = new Dictionary<string, object> { ["A1"] = "Data2" },
                ["Unchanged"] = new Dictionary<string, object> { ["A1"] = "Same" }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings { EnableSheetRenameDetection = true };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Equal(2, result.SheetsRenamed);
        }

        [Fact]
        public void SC040_CombinedRowAndCellChanges()
        {
            // Arrange - Row insertion combined with cell value changes
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Header",
                    ["B1"] = "OldValue",
                    ["A2"] = "Row2"
                }
            };

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Header",
                    ["B1"] = "NewValue",  // Value changed
                    ["A2"] = "InsertedRow",  // New row
                    ["A3"] = "Row2"  // Original row 2 moved
                }
            };

            var doc1 = CreateTestWorkbook(data1);
            var doc2 = CreateTestWorkbook(data2);
            var settings = new SmlComparerSettings { EnableRowAlignment = true };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - Should detect both types of changes
            Assert.True(result.TotalChanges >= 2);
        }

        #endregion

        #region Integration Test: Full Comparison Output for Manual Review

        /// <summary>
        /// This test creates a comprehensive comparison scenario with all types of changes,
        /// and saves the original, modified, and comparison result to disk for manual review.
        ///
        /// Files are saved to: TestOutput/SmlComparer/
        /// - Original.xlsx: The original workbook
        /// - Modified.xlsx: The modified workbook with various changes
        /// - Comparison.xlsx: The marked-up comparison result
        /// - ComparisonResult.json: JSON summary of all detected changes
        /// </summary>
        [Fact]
        public void SC041_FullComparison_SaveFilesForManualReview()
        {
            // Setup output directory
            var outputDir = Path.Combine(Path.GetTempPath(), "SmlComparer_TestOutput");
            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, true);
            Directory.CreateDirectory(outputDir);

            // === CREATE ORIGINAL WORKBOOK ===
            // Has 3 sheets with various data to demonstrate all change types
            var originalData = new Dictionary<string, Dictionary<string, object>>
            {
                // Sheet 1: "Sales Data" - Will have rows inserted/deleted and values changed
                ["Sales Data"] = new Dictionary<string, object>
                {
                    // Header row
                    ["A1"] = "Product",
                    ["B1"] = "Q1 Sales",
                    ["C1"] = "Q2 Sales",
                    ["D1"] = "Total",

                    // Data rows
                    ["A2"] = "Widget A",
                    ["B2"] = 1000,
                    ["C2"] = 1200,
                    ["D2"] = new CellWithFormula { Formula = "B2+C2", Value = "2200" },

                    ["A3"] = "Widget B",  // This row will be deleted
                    ["B3"] = 800,
                    ["C3"] = 900,
                    ["D3"] = new CellWithFormula { Formula = "B3+C3", Value = "1700" },

                    ["A4"] = "Widget C",
                    ["B4"] = 1500,
                    ["C4"] = 1600,
                    ["D4"] = new CellWithFormula { Formula = "B4+C4", Value = "3100" },

                    ["A5"] = "Widget D",
                    ["B5"] = 2000,
                    ["C5"] = 2100,  // This value will change
                    ["D5"] = new CellWithFormula { Formula = "B5+C5", Value = "4100" },

                    // Summary row
                    ["A7"] = "Grand Total",
                    ["D7"] = new CellWithFormula { Formula = "SUM(D2:D5)", Value = "11100" }
                },

                // Sheet 2: "Inventory" - Will be renamed to "Stock Levels"
                ["Inventory"] = new Dictionary<string, object>
                {
                    ["A1"] = "Item",
                    ["B1"] = "Quantity",
                    ["C1"] = "Location",

                    ["A2"] = "Part X-100",
                    ["B2"] = 500,
                    ["C2"] = "Warehouse A",

                    ["A3"] = "Part X-200",
                    ["B3"] = 300,
                    ["C3"] = "Warehouse B",

                    ["A4"] = "Part X-300",
                    ["B4"] = 750,
                    ["C4"] = "Warehouse A"
                },

                // Sheet 3: "Employees" - Will be deleted entirely
                ["Employees"] = new Dictionary<string, object>
                {
                    ["A1"] = "Name",
                    ["B1"] = "Department",

                    ["A2"] = "John Smith",
                    ["B2"] = "Sales",

                    ["A3"] = "Jane Doe",
                    ["B3"] = "Engineering"
                }
            };

            // === CREATE MODIFIED WORKBOOK ===
            var modifiedData = new Dictionary<string, Dictionary<string, object>>
            {
                // Sheet 1: "Sales Data" - Modified version
                ["Sales Data"] = new Dictionary<string, object>
                {
                    // Header row (unchanged)
                    ["A1"] = "Product",
                    ["B1"] = "Q1 Sales",
                    ["C1"] = "Q2 Sales",
                    ["D1"] = "Total",

                    // Data rows
                    ["A2"] = "Widget A",  // Unchanged
                    ["B2"] = 1000,
                    ["C2"] = 1200,
                    ["D2"] = new CellWithFormula { Formula = "B2+C2", Value = "2200" },

                    // NEW ROW INSERTED: Widget A-Plus
                    ["A3"] = "Widget A-Plus",
                    ["B3"] = 1100,
                    ["C3"] = 1300,
                    ["D3"] = new CellWithFormula { Formula = "B3+C3", Value = "2400" },

                    // Widget B DELETED - Widget C moves up
                    ["A4"] = "Widget C",
                    ["B4"] = 1500,
                    ["C4"] = 1600,
                    ["D4"] = new CellWithFormula { Formula = "B4+C4", Value = "3100" },

                    // Widget D with VALUE CHANGED (C5: 2100 -> 2500)
                    ["A5"] = "Widget D",
                    ["B5"] = 2000,
                    ["C5"] = 2500,  // CHANGED from 2100
                    ["D5"] = new CellWithFormula { Formula = "B5+C5", Value = "4500" },

                    // NEW ROW INSERTED: Widget E
                    ["A6"] = "Widget E",
                    ["B6"] = 3000,
                    ["C6"] = 3200,
                    ["D6"] = new CellWithFormula { Formula = "B6+C6", Value = "6200" },

                    // Summary row (moved down, formula changed)
                    ["A8"] = "Grand Total",
                    ["D8"] = new CellWithFormula { Formula = "SUM(D2:D6)", Value = "18400" }  // Formula changed
                },

                // Sheet 2: RENAMED from "Inventory" to "Stock Levels"
                ["Stock Levels"] = new Dictionary<string, object>
                {
                    ["A1"] = "Item",
                    ["B1"] = "Quantity",
                    ["C1"] = "Location",

                    ["A2"] = "Part X-100",
                    ["B2"] = 450,  // VALUE CHANGED from 500
                    ["C2"] = "Warehouse A",

                    ["A3"] = "Part X-200",
                    ["B3"] = 300,  // Unchanged
                    ["C3"] = "Warehouse B",

                    ["A4"] = "Part X-300",
                    ["B4"] = 750,  // Unchanged
                    ["C4"] = "Warehouse A",

                    // NEW ROW: Part X-400
                    ["A5"] = "Part X-400",
                    ["B5"] = 200,
                    ["C5"] = "Warehouse C"
                },

                // Sheet 3: "Employees" DELETED

                // Sheet 4: NEW SHEET "Contractors"
                ["Contractors"] = new Dictionary<string, object>
                {
                    ["A1"] = "Name",
                    ["B1"] = "Company",
                    ["C1"] = "Rate",

                    ["A2"] = "Bob Builder",
                    ["B2"] = "BuildCo",
                    ["C2"] = 75.00,

                    ["A3"] = "Alice Coder",
                    ["B3"] = "DevShop",
                    ["C3"] = 125.00
                }
            };

            // Create the workbooks
            var originalDoc = CreateTestWorkbook(originalData);
            var modifiedDoc = CreateTestWorkbook(modifiedData);

            // Configure comparison settings
            var settings = new SmlComparerSettings
            {
                EnableRowAlignment = true,
                EnableSheetRenameDetection = true,
                SheetRenameSimilarityThreshold = 0.6,
                CompareValues = true,
                CompareFormulas = true,
                CompareFormatting = true,
                AuthorForChanges = "SmlComparer Test"
            };

            // Run comparison
            var result = SmlComparer.Compare(originalDoc, modifiedDoc, settings);
            var markedDoc = SmlComparer.ProduceMarkedWorkbook(originalDoc, modifiedDoc, settings);

            // Save files
            var originalPath = Path.Combine(outputDir, "Original.xlsx");
            var modifiedPath = Path.Combine(outputDir, "Modified.xlsx");
            var comparisonPath = Path.Combine(outputDir, "Comparison.xlsx");
            var jsonPath = Path.Combine(outputDir, "ComparisonResult.json");

            File.WriteAllBytes(originalPath, originalDoc.DocumentByteArray);
            File.WriteAllBytes(modifiedPath, modifiedDoc.DocumentByteArray);
            File.WriteAllBytes(comparisonPath, markedDoc.DocumentByteArray);
            File.WriteAllText(jsonPath, result.ToJson());

            // Also write a human-readable summary
            var summaryPath = Path.Combine(outputDir, "Summary.txt");
            var summary = new StringBuilder();
            summary.AppendLine("=== SmlComparer Test Results ===");
            summary.AppendLine();
            summary.AppendLine($"Output Directory: {outputDir}");
            summary.AppendLine();
            summary.AppendLine("=== Files Generated ===");
            summary.AppendLine($"  - Original.xlsx: The original workbook");
            summary.AppendLine($"  - Modified.xlsx: The modified workbook");
            summary.AppendLine($"  - Comparison.xlsx: Marked-up comparison (open in Excel to see highlights)");
            summary.AppendLine($"  - ComparisonResult.json: JSON export of all changes");
            summary.AppendLine();
            summary.AppendLine("=== Summary Statistics ===");
            summary.AppendLine($"  Total Changes: {result.TotalChanges}");
            summary.AppendLine($"  Value Changes: {result.ValueChanges}");
            summary.AppendLine($"  Formula Changes: {result.FormulaChanges}");
            summary.AppendLine($"  Format Changes: {result.FormatChanges}");
            summary.AppendLine($"  Cells Added: {result.CellsAdded}");
            summary.AppendLine($"  Cells Deleted: {result.CellsDeleted}");
            summary.AppendLine($"  Sheets Added: {result.SheetsAdded}");
            summary.AppendLine($"  Sheets Deleted: {result.SheetsDeleted}");
            summary.AppendLine($"  Sheets Renamed: {result.SheetsRenamed}");
            summary.AppendLine($"  Rows Inserted: {result.RowsInserted}");
            summary.AppendLine($"  Rows Deleted: {result.RowsDeleted}");
            summary.AppendLine();
            summary.AppendLine("=== Expected Changes ===");
            summary.AppendLine("  1. Sheet 'Inventory' renamed to 'Stock Levels'");
            summary.AppendLine("  2. Sheet 'Employees' deleted");
            summary.AppendLine("  3. Sheet 'Contractors' added");
            summary.AppendLine("  4. Row inserted: 'Widget A-Plus' (row 3 in Sales Data)");
            summary.AppendLine("  5. Row deleted: 'Widget B' (was row 3 in original)");
            summary.AppendLine("  6. Row inserted: 'Widget E' (row 6 in Sales Data)");
            summary.AppendLine("  7. Value changed: Sales Data!C5 (2100 -> 2500)");
            summary.AppendLine("  8. Value changed: Stock Levels!B2 (500 -> 450)");
            summary.AppendLine("  9. Row inserted: 'Part X-400' in Stock Levels");
            summary.AppendLine("  10. Formula changed: Grand Total moved and formula updated");
            summary.AppendLine();
            summary.AppendLine("=== Detailed Changes ===");
            foreach (var change in result.Changes)
            {
                summary.AppendLine($"  - {change.GetDescription()}");
            }

            File.WriteAllText(summaryPath, summary.ToString());

            // Output path to console for easy access
            Console.WriteLine($"Test output saved to: {outputDir}");
            Console.WriteLine(summary.ToString());

            // Assertions to verify the test ran correctly
            Assert.True(result.TotalChanges > 0, "Should detect changes");
            Assert.True(File.Exists(originalPath), "Original.xlsx should exist");
            Assert.True(File.Exists(modifiedPath), "Modified.xlsx should exist");
            Assert.True(File.Exists(comparisonPath), "Comparison.xlsx should exist");
            Assert.True(File.Exists(jsonPath), "ComparisonResult.json should exist");

            // Verify specific expected changes
            Assert.True(result.SheetsRenamed >= 1, "Should detect sheet rename (Inventory -> Stock Levels)");
            Assert.True(result.SheetsDeleted >= 1, "Should detect sheet deletion (Employees)");
            Assert.True(result.SheetsAdded >= 1, "Should detect sheet addition (Contractors)");
        }

        #endregion

        #region Phase 3 Tests - Named Ranges, Comments, Data Validation, Merged Cells, Hyperlinks

        [Fact]
        public void SC041_NamedRange_Added_DetectedCorrectly()
        {
            // Arrange - original has no named ranges
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Value1",
                    ["A2"] = "Value2"
                }
            };

            var doc1 = CreateTestWorkbook(data1);

            // Create doc2 with a named range
            var doc2Bytes = CreateWorkbookWithNamedRange("TestRange", "Sheet1!$A$1:$A$2");
            var doc2 = new SmlDocument("modified.xlsx", doc2Bytes);

            var settings = new SmlComparerSettings { CompareNamedRanges = true, EnableRowAlignment = false };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Contains(result.Changes, c => c.ChangeType == SmlChangeType.NamedRangeAdded);
            Assert.True(result.NamedRangesAdded >= 1);
        }

        [Fact]
        public void SC042_NamedRange_Deleted_DetectedCorrectly()
        {
            // Arrange
            var doc1Bytes = CreateWorkbookWithNamedRange("TestRange", "Sheet1!$A$1:$A$2");
            var doc1 = new SmlDocument("original.xlsx", doc1Bytes);

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Value1",
                    ["A2"] = "Value2"
                }
            };
            var doc2 = CreateTestWorkbook(data2);

            var settings = new SmlComparerSettings { CompareNamedRanges = true, EnableRowAlignment = false };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Contains(result.Changes, c => c.ChangeType == SmlChangeType.NamedRangeDeleted);
            Assert.True(result.NamedRangesDeleted >= 1);
        }

        [Fact]
        public void SC043_NamedRange_Changed_DetectedCorrectly()
        {
            // Arrange
            var doc1Bytes = CreateWorkbookWithNamedRange("TestRange", "Sheet1!$A$1:$A$2");
            var doc1 = new SmlDocument("original.xlsx", doc1Bytes);

            var doc2Bytes = CreateWorkbookWithNamedRange("TestRange", "Sheet1!$A$1:$A$5");
            var doc2 = new SmlDocument("modified.xlsx", doc2Bytes);

            var settings = new SmlComparerSettings { CompareNamedRanges = true, EnableRowAlignment = false };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Contains(result.Changes, c => c.ChangeType == SmlChangeType.NamedRangeChanged);
            Assert.True(result.NamedRangesChanged >= 1);
        }

        [Fact]
        public void SC044_MergedCells_Added_DetectedCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Header",
                    ["B1"] = "",
                    ["C1"] = ""
                }
            };
            var doc1 = CreateTestWorkbook(data1);

            // Create doc2 with merged cells
            var doc2Bytes = CreateWorkbookWithMergedCells("A1:C1");
            var doc2 = new SmlDocument("modified.xlsx", doc2Bytes);

            var settings = new SmlComparerSettings { CompareMergedCells = true, EnableRowAlignment = false };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Contains(result.Changes, c => c.ChangeType == SmlChangeType.MergedCellAdded);
            Assert.True(result.MergedCellsAdded >= 1);
        }

        [Fact]
        public void SC045_MergedCells_Deleted_DetectedCorrectly()
        {
            // Arrange
            var doc1Bytes = CreateWorkbookWithMergedCells("A1:C1");
            var doc1 = new SmlDocument("original.xlsx", doc1Bytes);

            var data2 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Header",
                    ["B1"] = "",
                    ["C1"] = ""
                }
            };
            var doc2 = CreateTestWorkbook(data2);

            var settings = new SmlComparerSettings { CompareMergedCells = true, EnableRowAlignment = false };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Contains(result.Changes, c => c.ChangeType == SmlChangeType.MergedCellDeleted);
            Assert.True(result.MergedCellsDeleted >= 1);
        }

        [Fact]
        public void SC046_Hyperlink_Added_DetectedCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Click here"
                }
            };
            var doc1 = CreateTestWorkbook(data1);

            var doc2Bytes = CreateWorkbookWithHyperlink("A1", "https://example.com");
            var doc2 = new SmlDocument("modified.xlsx", doc2Bytes);

            var settings = new SmlComparerSettings { CompareHyperlinks = true, EnableRowAlignment = false };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Contains(result.Changes, c => c.ChangeType == SmlChangeType.HyperlinkAdded);
            Assert.True(result.HyperlinksAdded >= 1);
        }

        [Fact]
        public void SC047_Hyperlink_Changed_DetectedCorrectly()
        {
            // Arrange
            var doc1Bytes = CreateWorkbookWithHyperlink("A1", "https://old-example.com");
            var doc1 = new SmlDocument("original.xlsx", doc1Bytes);

            var doc2Bytes = CreateWorkbookWithHyperlink("A1", "https://new-example.com");
            var doc2 = new SmlDocument("modified.xlsx", doc2Bytes);

            var settings = new SmlComparerSettings { CompareHyperlinks = true, EnableRowAlignment = false };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Contains(result.Changes, c => c.ChangeType == SmlChangeType.HyperlinkChanged);
            Assert.True(result.HyperlinksChanged >= 1);
        }

        [Fact]
        public void SC048_DataValidation_Added_DetectedCorrectly()
        {
            // Arrange
            var data1 = new Dictionary<string, Dictionary<string, object>>
            {
                ["Sheet1"] = new Dictionary<string, object>
                {
                    ["A1"] = "Status"
                }
            };
            var doc1 = CreateTestWorkbook(data1);

            var doc2Bytes = CreateWorkbookWithDataValidation("A2", new[] { "Active", "Inactive", "Pending" });
            var doc2 = new SmlDocument("modified.xlsx", doc2Bytes);

            var settings = new SmlComparerSettings { CompareDataValidation = true, EnableRowAlignment = false };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.Contains(result.Changes, c => c.ChangeType == SmlChangeType.DataValidationAdded);
            Assert.True(result.DataValidationsAdded >= 1);
        }

        [Fact]
        public void SC049_Phase3_Statistics_CorrectlySummarized()
        {
            // Arrange - Create workbooks with multiple Phase 3 features
            var doc1Bytes = CreateWorkbookWithNamedRange("Range1", "Sheet1!$A$1");
            var doc1 = new SmlDocument("original.xlsx", doc1Bytes);

            // Modified version: different named range, added merged cells
            var doc2Bytes = CreateWorkbookWithPhase3Features();
            var doc2 = new SmlDocument("modified.xlsx", doc2Bytes);

            var settings = new SmlComparerSettings
            {
                CompareNamedRanges = true,
                CompareMergedCells = true,
                CompareHyperlinks = true,
                EnableRowAlignment = false
            };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - should detect various Phase 3 changes
            Assert.True(result.TotalChanges > 0, "Should detect Phase 3 changes");
        }

        [Fact]
        public void SC050_Phase3_Features_DisabledByDefault()
        {
            // Arrange
            var doc1Bytes = CreateWorkbookWithNamedRange("Range1", "Sheet1!$A$1");
            var doc1 = new SmlDocument("original.xlsx", doc1Bytes);

            var doc2Bytes = CreateWorkbookWithNamedRange("Range1", "Sheet1!$A$1:$A$5");
            var doc2 = new SmlDocument("modified.xlsx", doc2Bytes);

            // Settings with Phase 3 features disabled
            var settings = new SmlComparerSettings
            {
                CompareNamedRanges = false,
                CompareMergedCells = false,
                CompareHyperlinks = false,
                CompareDataValidation = false,
                CompareComments = false,
                EnableRowAlignment = false
            };

            // Act
            var result = SmlComparer.Compare(doc1, doc2, settings);

            // Assert - should not detect named range changes when disabled
            Assert.DoesNotContain(result.Changes, c => c.ChangeType == SmlChangeType.NamedRangeChanged);
        }

        #region Phase 3 Test Helpers

        private static byte[] CreateWorkbookWithNamedRange(string name, string reference)
        {
            using var ms = new MemoryStream();
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                // Add worksheet
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(
                    new DocumentFormat.OpenXml.Spreadsheet.SheetData(
                        new DocumentFormat.OpenXml.Spreadsheet.Row(
                            new DocumentFormat.OpenXml.Spreadsheet.Cell
                            {
                                CellReference = "A1",
                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Value1"),
                                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                            },
                            new DocumentFormat.OpenXml.Spreadsheet.Cell
                            {
                                CellReference = "A2",
                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Value2"),
                                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                            }
                        ) { RowIndex = 1 }
                    )
                );

                // Add sheets
                workbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets(
                    new DocumentFormat.OpenXml.Spreadsheet.Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    }
                );

                // Add defined names
                workbookPart.Workbook.DefinedNames = new DocumentFormat.OpenXml.Spreadsheet.DefinedNames(
                    new DocumentFormat.OpenXml.Spreadsheet.DefinedName(reference) { Name = name }
                );
            }
            return ms.ToArray();
        }

        private static byte[] CreateWorkbookWithMergedCells(string mergeRange)
        {
            using var ms = new MemoryStream();
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(
                    new DocumentFormat.OpenXml.Spreadsheet.SheetData(
                        new DocumentFormat.OpenXml.Spreadsheet.Row(
                            new DocumentFormat.OpenXml.Spreadsheet.Cell
                            {
                                CellReference = "A1",
                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Header"),
                                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                            }
                        ) { RowIndex = 1 }
                    ),
                    new DocumentFormat.OpenXml.Spreadsheet.MergeCells(
                        new DocumentFormat.OpenXml.Spreadsheet.MergeCell { Reference = mergeRange }
                    )
                );

                workbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets(
                    new DocumentFormat.OpenXml.Spreadsheet.Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    }
                );
            }
            return ms.ToArray();
        }

        private static byte[] CreateWorkbookWithHyperlink(string cellRef, string url)
        {
            using var ms = new MemoryStream();
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                // Add hyperlink relationship
                var hyperlinkRel = worksheetPart.AddHyperlinkRelationship(new Uri(url), true);

                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(
                    new DocumentFormat.OpenXml.Spreadsheet.SheetData(
                        new DocumentFormat.OpenXml.Spreadsheet.Row(
                            new DocumentFormat.OpenXml.Spreadsheet.Cell
                            {
                                CellReference = cellRef,
                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Click here"),
                                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                            }
                        ) { RowIndex = 1 }
                    ),
                    new DocumentFormat.OpenXml.Spreadsheet.Hyperlinks(
                        new DocumentFormat.OpenXml.Spreadsheet.Hyperlink
                        {
                            Reference = cellRef,
                            Id = hyperlinkRel.Id
                        }
                    )
                );

                workbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets(
                    new DocumentFormat.OpenXml.Spreadsheet.Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    }
                );
            }
            return ms.ToArray();
        }

        private static byte[] CreateWorkbookWithDataValidation(string cellRef, string[] listItems)
        {
            using var ms = new MemoryStream();
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(
                    new DocumentFormat.OpenXml.Spreadsheet.SheetData(
                        new DocumentFormat.OpenXml.Spreadsheet.Row(
                            new DocumentFormat.OpenXml.Spreadsheet.Cell
                            {
                                CellReference = "A1",
                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Status"),
                                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                            }
                        ) { RowIndex = 1 }
                    ),
                    new DocumentFormat.OpenXml.Spreadsheet.DataValidations(
                        new DocumentFormat.OpenXml.Spreadsheet.DataValidation
                        {
                            Type = DocumentFormat.OpenXml.Spreadsheet.DataValidationValues.List,
                            AllowBlank = true,
                            ShowInputMessage = true,
                            ShowErrorMessage = true,
                            SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue>(
                                new[] { new DocumentFormat.OpenXml.StringValue(cellRef) }
                            ),
                            Formula1 = new DocumentFormat.OpenXml.Spreadsheet.Formula1($"\"{string.Join(",", listItems)}\"")
                        }
                    ) { Count = 1 }
                );

                workbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets(
                    new DocumentFormat.OpenXml.Spreadsheet.Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    }
                );
            }
            return ms.ToArray();
        }

        private static byte[] CreateWorkbookWithPhase3Features()
        {
            using var ms = new MemoryStream();
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                // Add hyperlink
                var hyperlinkRel = worksheetPart.AddHyperlinkRelationship(new Uri("https://example.com"), true);

                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(
                    new DocumentFormat.OpenXml.Spreadsheet.SheetData(
                        new DocumentFormat.OpenXml.Spreadsheet.Row(
                            new DocumentFormat.OpenXml.Spreadsheet.Cell
                            {
                                CellReference = "A1",
                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Header"),
                                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                            },
                            new DocumentFormat.OpenXml.Spreadsheet.Cell
                            {
                                CellReference = "B1",
                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Link"),
                                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                            }
                        ) { RowIndex = 1 }
                    ),
                    new DocumentFormat.OpenXml.Spreadsheet.MergeCells(
                        new DocumentFormat.OpenXml.Spreadsheet.MergeCell { Reference = "A1:A2" }
                    ),
                    new DocumentFormat.OpenXml.Spreadsheet.Hyperlinks(
                        new DocumentFormat.OpenXml.Spreadsheet.Hyperlink
                        {
                            Reference = "B1",
                            Id = hyperlinkRel.Id
                        }
                    )
                );

                workbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets(
                    new DocumentFormat.OpenXml.Spreadsheet.Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    }
                );

                // Different named range
                workbookPart.Workbook.DefinedNames = new DocumentFormat.OpenXml.Spreadsheet.DefinedNames(
                    new DocumentFormat.OpenXml.Spreadsheet.DefinedName("Sheet1!$A$1:$B$2") { Name = "Range2" }
                );
            }
            return ms.ToArray();
        }

        #endregion

        #endregion
    }
}

#endif
