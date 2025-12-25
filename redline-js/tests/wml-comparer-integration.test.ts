/**
 * Integration tests for WmlComparer
 *
 * Tests the end-to-end comparison of Word documents.
 */

import { describe, it, expect } from 'vitest';
import { readFile } from 'fs/promises';
import { join } from 'path';
import { compareDocuments, countDocumentRevisions, buildChangeList, WmlChangeType } from '../src/wml/wml-comparer';

const TEST_FILES_DIR = join(__dirname, '../../TestFiles');

describe('WmlComparer Integration', () => {
  it('compares identical documents with no changes', async () => {
    const docPath = join(TEST_FILES_DIR, 'CA/CA001-Plain.docx');
    let data: Buffer;

    try {
      data = await readFile(docPath);
    } catch {
      console.warn('Skipping: test fixture not found');
      return;
    }

    const result = await compareDocuments(data, data);

    expect(result.insertions).toBe(0);
    expect(result.deletions).toBe(0);
    expect(result.revisionCount).toBe(0);
    expect(result.document).toBeInstanceOf(Buffer);
  });

  it('detects changes between different documents', async () => {
    const doc1Path = join(TEST_FILES_DIR, 'WC/WC001-Digits.docx');
    const doc2Path = join(TEST_FILES_DIR, 'WC/WC001-Digits-Mod.docx');

    let doc1: Buffer, doc2: Buffer;

    try {
      doc1 = await readFile(doc1Path);
      doc2 = await readFile(doc2Path);
    } catch {
      console.warn('Skipping: test fixtures not found');
      return;
    }

    const result = await compareDocuments(doc1, doc2);

    // These documents should have differences
    expect(result.revisionCount).toBeGreaterThan(0);
    expect(result.document).toBeInstanceOf(Buffer);

    console.log(
      `Detected ${result.insertions} insertions, ${result.deletions} deletions, ${result.revisionCount} total`
    );
  });

  it('countDocumentRevisions returns quick count', async () => {
    const doc1Path = join(TEST_FILES_DIR, 'WC/WC001-Digits.docx');
    const doc2Path = join(TEST_FILES_DIR, 'WC/WC001-Digits-Mod.docx');

    let doc1: Buffer, doc2: Buffer;

    try {
      doc1 = await readFile(doc1Path);
      doc2 = await readFile(doc2Path);
    } catch {
      console.warn('Skipping: test fixtures not found');
      return;
    }

    const counts = await countDocumentRevisions(doc1, doc2);

    expect(counts.total).toBeGreaterThan(0);
    console.log(`Quick count: ${counts.insertions} ins, ${counts.deletions} del, ${counts.total} total`);
  });

  it('uses custom author and date', async () => {
    const docPath = join(TEST_FILES_DIR, 'CA/CA001-Plain.docx');
    let data: Buffer;

    try {
      data = await readFile(docPath);
    } catch {
      console.warn('Skipping: test fixture not found');
      return;
    }

    const result = await compareDocuments(data, data, {
      author: 'Custom Author',
      dateTime: new Date('2024-06-15T12:00:00Z'),
    });

    expect(result.document).toBeInstanceOf(Buffer);
  });
});

describe('WmlComparer Test Cases', () => {
  it('WC-1000: handles basic text comparison', async () => {
    const doc1Path = join(TEST_FILES_DIR, 'WC/WC001-Digits.docx');
    const doc2Path = join(TEST_FILES_DIR, 'WC/WC001-Digits-Mod.docx');

    let doc1: Buffer, doc2: Buffer;
    try {
      doc1 = await readFile(doc1Path);
      doc2 = await readFile(doc2Path);
    } catch {
      console.warn('Skipping: test fixtures not found');
      return;
    }

    const result = await compareDocuments(doc1, doc2);

    expect(result.revisionCount).toBe(4);
    expect(result.insertions).toBe(2);
    expect(result.deletions).toBe(2);
  });
});

describe('WmlComparer Changes Array', () => {
  it('returns changes array with individual change details', async () => {
    const doc1Path = join(TEST_FILES_DIR, 'WC/WC001-Digits.docx');
    const doc2Path = join(TEST_FILES_DIR, 'WC/WC001-Digits-Mod.docx');

    let doc1: Buffer, doc2: Buffer;
    try {
      doc1 = await readFile(doc1Path);
      doc2 = await readFile(doc2Path);
    } catch {
      console.warn('Skipping: test fixtures not found');
      return;
    }

    const result = await compareDocuments(doc1, doc2);

    expect(result.changes).toBeDefined();
    expect(Array.isArray(result.changes)).toBe(true);
    expect(result.changes.length).toBeGreaterThan(0);

    for (const change of result.changes) {
      expect(change.changeType).toBeDefined();
      expect(change.revisionId).toBeDefined();
    }
  });

  it('buildChangeList creates UI-friendly change items', async () => {
    const doc1Path = join(TEST_FILES_DIR, 'WC/WC001-Digits.docx');
    const doc2Path = join(TEST_FILES_DIR, 'WC/WC001-Digits-Mod.docx');

    let doc1: Buffer, doc2: Buffer;
    try {
      doc1 = await readFile(doc1Path);
      doc2 = await readFile(doc2Path);
    } catch {
      console.warn('Skipping: test fixtures not found');
      return;
    }

    const result = await compareDocuments(doc1, doc2);
    const changeList = buildChangeList(result);

    expect(changeList.length).toBeGreaterThan(0);

    for (const item of changeList) {
      expect(item.id).toBeDefined();
      expect(item.id).toMatch(/^change-\d+$/);
      expect(item.changeType).toBeDefined();
      expect(item.summary).toBeDefined();
      expect(item.anchor).toBeDefined();
    }
  });

  it('buildChangeList merges adjacent delete+insert into replacements', async () => {
    const doc1Path = join(TEST_FILES_DIR, 'WC/WC002-Unmodified.docx');
    const doc2Path = join(TEST_FILES_DIR, 'WC/WC002-DiffInMiddle.docx');

    let doc1: Buffer, doc2: Buffer;
    try {
      doc1 = await readFile(doc1Path);
      doc2 = await readFile(doc2Path);
    } catch {
      console.warn('Skipping: test fixtures not found');
      return;
    }

    const result = await compareDocuments(doc1, doc2);
    const changeList = buildChangeList(result, { mergeReplacements: true });

    const replacements = changeList.filter(c => c.changeType === WmlChangeType.TextReplaced);
    console.log(`Found ${replacements.length} replacements, ${changeList.length} total items`);
  });
});
