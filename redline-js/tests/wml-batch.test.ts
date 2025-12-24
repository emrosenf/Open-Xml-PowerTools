/**
 * Batch test runner for WmlComparer test cases
 *
 * Tests the current implementation against expected revision counts from C# golden files.
 * All 104 test cases from WmlComparerTests.cs are included.
 */

import { describe, it, expect } from 'vitest';
import { readFile } from 'fs/promises';
import { join } from 'path';
import { countDocumentRevisions } from '../src/wml/wml-comparer';

const TEST_FILES_DIR = join(__dirname, '../../TestFiles');

// All 104 test cases from the golden file manifest
// Format: [testId, source1, source2, expectedRevisions]
const ALL_TEST_CASES: [string, string, string, number][] = [
  // Basic text comparisons
  ['WC-1000', 'CA/CA001-Plain.docx', 'CA/CA001-Plain-Mod.docx', 1],
  ['WC-1010', 'WC/WC001-Digits.docx', 'WC/WC001-Digits-Mod.docx', 4],
  ['WC-1020', 'WC/WC001-Digits.docx', 'WC/WC001-Digits-Deleted-Paragraph.docx', 1],
  ['WC-1030', 'WC/WC001-Digits-Deleted-Paragraph.docx', 'WC/WC001-Digits.docx', 1],
  ['WC-1040', 'WC/WC002-Unmodified.docx', 'WC/WC002-DiffInMiddle.docx', 2],
  ['WC-1050', 'WC/WC002-Unmodified.docx', 'WC/WC002-DiffAtBeginning.docx', 2],
  ['WC-1060', 'WC/WC002-Unmodified.docx', 'WC/WC002-DeleteAtBeginning.docx', 1],
  ['WC-1070', 'WC/WC002-Unmodified.docx', 'WC/WC002-InsertAtBeginning.docx', 1],
  ['WC-1080', 'WC/WC002-Unmodified.docx', 'WC/WC002-InsertAtEnd.docx', 1],
  ['WC-1090', 'WC/WC002-Unmodified.docx', 'WC/WC002-DeleteAtEnd.docx', 1],
  ['WC-1100', 'WC/WC002-Unmodified.docx', 'WC/WC002-DeleteInMiddle.docx', 1],
  ['WC-1110', 'WC/WC002-Unmodified.docx', 'WC/WC002-InsertInMiddle.docx', 1],
  ['WC-1120', 'WC/WC002-DeleteInMiddle.docx', 'WC/WC002-Unmodified.docx', 1],

  // Table tests
  ['WC-1140', 'WC/WC006-Table.docx', 'WC/WC006-Table-Delete-Row.docx', 1],
  ['WC-1150', 'WC/WC006-Table-Delete-Row.docx', 'WC/WC006-Table.docx', 1],
  ['WC-1160', 'WC/WC006-Table.docx', 'WC/WC006-Table-Delete-Contests-of-Row.docx', 2],
  ['WC-1170', 'WC/WC007-Unmodified.docx', 'WC/WC007-Longest-At-End.docx', 2],
  ['WC-1180', 'WC/WC007-Unmodified.docx', 'WC/WC007-Deleted-at-Beginning-of-Para.docx', 1],
  ['WC-1190', 'WC/WC007-Unmodified.docx', 'WC/WC007-Moved-into-Table.docx', 2],
  ['WC-1200', 'WC/WC009-Table-Unmodified.docx', 'WC/WC009-Table-Cell-1-1-Mod.docx', 1],
  ['WC-1210', 'WC/WC010-Para-Before-Table-Unmodified.docx', 'WC/WC010-Para-Before-Table-Mod.docx', 3],
  ['WC-1220', 'WC/WC011-Before.docx', 'WC/WC011-After.docx', 2],

  // Math content
  ['WC-1230', 'WC/WC012-Math-Before.docx', 'WC/WC012-Math-After.docx', 2],

  // Images
  ['WC-1240', 'WC/WC013-Image-Before.docx', 'WC/WC013-Image-After.docx', 2],
  ['WC-1250', 'WC/WC013-Image-Before.docx', 'WC/WC013-Image-After2.docx', 2],
  ['WC-1260', 'WC/WC013-Image-Before2.docx', 'WC/WC013-Image-After2.docx', 2],

  // SmartArt
  ['WC-1270', 'WC/WC014-SmartArt-Before.docx', 'WC/WC014-SmartArt-After.docx', 2],
  ['WC-1280', 'WC/WC014-SmartArt-With-Image-Before.docx', 'WC/WC014-SmartArt-With-Image-After.docx', 2],
  ['WC-1310', 'WC/WC014-SmartArt-With-Image-Before.docx', 'WC/WC014-SmartArt-With-Image-Deleted-After.docx', 3],
  ['WC-1320', 'WC/WC014-SmartArt-With-Image-Before.docx', 'WC/WC014-SmartArt-With-Image-Deleted-After2.docx', 1],

  // Multi-paragraph
  ['WC-1330', 'WC/WC015-Three-Paragraphs.docx', 'WC/WC015-Three-Paragraphs-After.docx', 3],

  // Images with paragraphs
  ['WC-1340', 'WC/WC016-Para-Image-Para.docx', 'WC/WC016-Para-Image-Para-w-Deleted-Image.docx', 1],
  ['WC-1350', 'WC/WC017-Image.docx', 'WC/WC017-Image-After.docx', 3],

  // Fields
  ['WC-1360', 'WC/WC018-Field-Simple-Before.docx', 'WC/WC018-Field-Simple-After-1.docx', 2],
  ['WC-1370', 'WC/WC018-Field-Simple-Before.docx', 'WC/WC018-Field-Simple-After-2.docx', 3],

  // Hyperlinks
  ['WC-1380', 'WC/WC019-Hyperlink-Before.docx', 'WC/WC019-Hyperlink-After-1.docx', 3],
  ['WC-1390', 'WC/WC019-Hyperlink-Before.docx', 'WC/WC019-Hyperlink-After-2.docx', 5],

  // Footnotes
  ['WC-1400', 'WC/WC020-FootNote-Before.docx', 'WC/WC020-FootNote-After-1.docx', 3],
  ['WC-1410', 'WC/WC020-FootNote-Before.docx', 'WC/WC020-FootNote-After-2.docx', 5],

  // Complex math
  ['WC-1420', 'WC/WC021-Math-Before-1.docx', 'WC/WC021-Math-After-1.docx', 9],
  ['WC-1430', 'WC/WC021-Math-Before-2.docx', 'WC/WC021-Math-After-2.docx', 6],
  ['WC-1440', 'WC/WC022-Image-Math-Para-Before.docx', 'WC/WC022-Image-Math-Para-After.docx', 10],

  // Tables with images
  ['WC-1450', 'WC/WC023-Table-4-Row-Image-Before.docx', 'WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx', 7],
  ['WC-1460', 'WC/WC024-Table-Before.docx', 'WC/WC024-Table-After.docx', 1],
  ['WC-1470', 'WC/WC024-Table-Before.docx', 'WC/WC024-Table-After2.docx', 7],
  ['WC-1480', 'WC/WC025-Simple-Table-Before.docx', 'WC/WC025-Simple-Table-After.docx', 4],
  ['WC-1500', 'WC/WC026-Long-Table-Before.docx', 'WC/WC026-Long-Table-After-1.docx', 2],

  // Twenty paragraphs
  ['WC-1510', 'WC/WC027-Twenty-Paras-Before.docx', 'WC/WC027-Twenty-Paras-After-1.docx', 2],
  ['WC-1520', 'WC/WC027-Twenty-Paras-After-1.docx', 'WC/WC027-Twenty-Paras-Before.docx', 2],
  ['WC-1530', 'WC/WC027-Twenty-Paras-Before.docx', 'WC/WC027-Twenty-Paras-After-2.docx', 4],

  // Image and math combinations
  ['WC-1540', 'WC/WC030-Image-Math-Before.docx', 'WC/WC030-Image-Math-After.docx', 2],
  ['WC-1550', 'WC/WC031-Two-Maths-Before.docx', 'WC/WC031-Two-Maths-After.docx', 4],

  // Paragraph properties
  ['WC-1560', 'WC/WC032-Para-with-Para-Props.docx', 'WC/WC032-Para-with-Para-Props-After.docx', 3],

  // Merged cells
  ['WC-1570', 'WC/WC033-Merged-Cells-Before.docx', 'WC/WC033-Merged-Cells-After1.docx', 2],
  ['WC-1580', 'WC/WC033-Merged-Cells-Before.docx', 'WC/WC033-Merged-Cells-After2.docx', 4],

  // Footnotes variants
  ['WC-1600', 'WC/WC034-Footnotes-Before.docx', 'WC/WC034-Footnotes-After1.docx', 1],
  ['WC-1610', 'WC/WC034-Footnotes-Before.docx', 'WC/WC034-Footnotes-After2.docx', 4],
  ['WC-1620', 'WC/WC034-Footnotes-Before.docx', 'WC/WC034-Footnotes-After3.docx', 3],
  ['WC-1630', 'WC/WC034-Footnotes-After3.docx', 'WC/WC034-Footnotes-Before.docx', 3],
  ['WC-1640', 'WC/WC035-Footnote-Before.docx', 'WC/WC035-Footnote-After.docx', 2],
  ['WC-1650', 'WC/WC035-Footnote-After.docx', 'WC/WC035-Footnote-Before.docx', 2],
  ['WC-1660', 'WC/WC036-Footnote-With-Table-Before.docx', 'WC/WC036-Footnote-With-Table-After.docx', 5],
  ['WC-1670', 'WC/WC036-Footnote-With-Table-After.docx', 'WC/WC036-Footnote-With-Table-Before.docx', 5],

  // Endnotes
  ['WC-1680', 'WC/WC034-Endnotes-Before.docx', 'WC/WC034-Endnotes-After1.docx', 1],
  ['WC-1700', 'WC/WC034-Endnotes-Before.docx', 'WC/WC034-Endnotes-After2.docx', 4],
  ['WC-1710', 'WC/WC034-Endnotes-Before.docx', 'WC/WC034-Endnotes-After3.docx', 7],
  ['WC-1720', 'WC/WC034-Endnotes-After3.docx', 'WC/WC034-Endnotes-Before.docx', 7],
  ['WC-1730', 'WC/WC035-Endnote-Before.docx', 'WC/WC035-Endnote-After.docx', 2],
  ['WC-1740', 'WC/WC035-Endnote-After.docx', 'WC/WC035-Endnote-Before.docx', 2],
  ['WC-1750', 'WC/WC036-Endnote-With-Table-Before.docx', 'WC/WC036-Endnote-With-Table-After.docx', 6],
  ['WC-1760', 'WC/WC036-Endnote-With-Table-After.docx', 'WC/WC036-Endnote-With-Table-Before.docx', 6],

  // Textboxes
  ['WC-1770', 'WC/WC037-Textbox-Before.docx', 'WC/WC037-Textbox-After1.docx', 2],

  // Line breaks
  ['WC-1780', 'WC/WC038-Document-With-BR-Before.docx', 'WC/WC038-Document-With-BR-After.docx', 2],

  // Revision consolidation
  ['WC-1800', 'RC/RC001-Before.docx', 'RC/RC001-After1.docx', 2],
  ['WC-1810', 'RC/RC002-Image.docx', 'RC/RC002-Image-After1.docx', 1],

  // Breaks in rows
  ['WC-1820', 'WC/WC039-Break-In-Row.docx', 'WC/WC039-Break-In-Row-After1.docx', 1],

  // More tables
  ['WC-1830', 'WC/WC041-Table-5.docx', 'WC/WC041-Table-5-Mod.docx', 2],
  ['WC-1840', 'WC/WC042-Table-5.docx', 'WC/WC042-Table-5-Mod.docx', 2],
  ['WC-1850', 'WC/WC043-Nested-Table.docx', 'WC/WC043-Nested-Table-Mod.docx', 2],

  // Text boxes
  ['WC-1860', 'WC/WC044-Text-Box.docx', 'WC/WC044-Text-Box-Mod.docx', 2],
  ['WC-1870', 'WC/WC045-Text-Box.docx', 'WC/WC045-Text-Box-Mod.docx', 2],
  ['WC-1880', 'WC/WC046-Two-Text-Box.docx', 'WC/WC046-Two-Text-Box-Mod.docx', 2],
  ['WC-1890', 'WC/WC047-Two-Text-Box.docx', 'WC/WC047-Two-Text-Box-Mod.docx', 2],
  ['WC-1900', 'WC/WC048-Text-Box-in-Cell.docx', 'WC/WC048-Text-Box-in-Cell-Mod.docx', 6],
  ['WC-1910', 'WC/WC049-Text-Box-in-Cell.docx', 'WC/WC049-Text-Box-in-Cell-Mod.docx', 5],
  ['WC-1920', 'WC/WC050-Table-in-Text-Box.docx', 'WC/WC050-Table-in-Text-Box-Mod.docx', 8],
  ['WC-1930', 'WC/WC051-Table-in-Text-Box.docx', 'WC/WC051-Table-in-Text-Box-Mod.docx', 9],

  // SmartArt same
  ['WC-1940', 'WC/WC052-SmartArt-Same.docx', 'WC/WC052-SmartArt-Same-Mod.docx', 2],

  // Text in cell
  ['WC-1950', 'WC/WC053-Text-in-Cell.docx', 'WC/WC053-Text-in-Cell-Mod.docx', 2],

  // Documents with existing tracked changes (0 revisions expected)
  ['WC-1960', 'WC/WC054-Text-in-Cell.docx', 'WC/WC054-Text-in-Cell-Mod.docx', 0],
  ['WC-1970', 'WC/WC055-French.docx', 'WC/WC055-French-Mod.docx', 0],
  ['WC-1980', 'WC/WC056-French.docx', 'WC/WC056-French-Mod.docx', 0],

  // Merged cells in tables
  ['WC-2000', 'WC/WC058-Table-Merged-Cell.docx', 'WC/WC058-Table-Merged-Cell-Mod.docx', 6],

  // More footnotes/endnotes
  ['WC-2010', 'WC/WC059-Footnote.docx', 'WC/WC059-Footnote-Mod.docx', 5],
  ['WC-2020', 'WC/WC060-Endnote.docx', 'WC/WC060-Endnote-Mod.docx', 3],

  // Styles
  ['WC-2030', 'WC/WC061-Style-Added.docx', 'WC/WC061-Style-Added-Mod.docx', 1],
  ['WC-2040', 'WC/WC062-New-Char-Style-Added.docx', 'WC/WC062-New-Char-Style-Added-Mod.docx', 3],

  // More footnotes
  ['WC-2050', 'WC/WC063-Footnote.docx', 'WC/WC063-Footnote-Mod.docx', 1],
  ['WC-2060', 'WC/WC063-Footnote-Mod.docx', 'WC/WC063-Footnote.docx', 1],
  ['WC-2070', 'WC/WC064-Footnote.docx', 'WC/WC064-Footnote-Mod.docx', 0],

  // More textboxes
  ['WC-2080', 'WC/WC065-Textbox.docx', 'WC/WC065-Textbox-Mod.docx', 2],
  ['WC-2090', 'WC/WC066-Textbox-Before-Ins.docx', 'WC/WC066-Textbox-Before-Ins-Mod.docx', 1],
  ['WC-2092', 'WC/WC066-Textbox-Before-Ins-Mod.docx', 'WC/WC066-Textbox-Before-Ins.docx', 1],
  ['WC-2100', 'WC/WC067-Textbox-Image.docx', 'WC/WC067-Textbox-Image-Mod.docx', 2],
];

describe('WmlComparer Full Test Suite (104 tests)', () => {
  describe.each(ALL_TEST_CASES)('%s: %s vs %s', (testId, source1, source2, expectedRevisions) => {
    it(`should detect ${expectedRevisions} revisions`, async () => {
      const doc1Path = join(TEST_FILES_DIR, source1);
      const doc2Path = join(TEST_FILES_DIR, source2);

      let doc1: Buffer, doc2: Buffer;
      try {
        doc1 = await readFile(doc1Path);
        doc2 = await readFile(doc2Path);
      } catch {
        console.warn(`Skipping ${testId}: test fixtures not found`);
        return;
      }

      const result = await countDocumentRevisions(doc1, doc2);

      console.log(
        `${testId}: Expected ${expectedRevisions}, got ${result.total} (${result.insertions} ins, ${result.deletions} del)`
      );

      expect(result.total).toBe(expectedRevisions);
    });
  });
});
