/**
 * Batch test runner for WmlComparer test cases
 *
 * Tests the current implementation against expected revision counts
 */

import { describe, it, expect } from 'vitest';
import { readFile } from 'fs/promises';
import { join } from 'path';
import { countDocumentRevisions } from '../src/wml/wml-comparer';

const TEST_FILES_DIR = join(__dirname, '../../TestFiles');

// Test cases: [testId, source1, source2, expectedRevisions]
const TEST_CASES: [string, string, string, number][] = [
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
  ['WC-1330', 'WC/WC015-Three-Paragraphs.docx', 'WC/WC015-Three-Paragraphs-After.docx', 3],
];

describe('WmlComparer Batch Tests', () => {
  describe.each(TEST_CASES)('%s: %s vs %s', (testId, source1, source2, expectedRevisions) => {
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
