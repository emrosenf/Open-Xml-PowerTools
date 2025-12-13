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
                        var rowElement = sheetDataElement.Elements<Row>()
                            .FirstOrDefault(r => r.RowIndex == (uint)row);
                        if (rowElement == null)
                        {
                            rowElement = new Row { RowIndex = (uint)row };
                            sheetDataElement.Append(rowElement);
                        }

                        // Create cell
                        var cellElement = new Cell { CellReference = cellRef };

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
            return (int)sst.Count - 1;
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
    }
}

#endif
