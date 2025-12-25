/**
 * Regression tests for WmlComparer bug fixes
 *
 * These tests ensure that fixed issues don't reappear:
 * 1. Embedded images are preserved (not converted to DRAWING_ text tokens)
 * 2. All rId references in output are valid (exist in relationships file)
 * 3. No orphan sectPr elements with invalid header/footer references
 */

import { describe, it, expect } from 'vitest';
import { readFile } from 'fs/promises';
import { join } from 'path';
import JSZip from 'jszip';
import { compareDocuments } from '../src/wml/wml-comparer';

const TEST_FILES_DIR = join(__dirname, '../../TestFiles');

async function loadTestFile(relativePath: string): Promise<Buffer | null> {
  try {
    return await readFile(join(TEST_FILES_DIR, relativePath));
  } catch {
    return null;
  }
}

async function extractDocumentXml(docxBuffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(docxBuffer);
  const file = zip.file('word/document.xml');
  if (!file) throw new Error('document.xml not found');
  return file.async('string');
}

async function extractRelationships(docxBuffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(docxBuffer);
  const file = zip.file('word/_rels/document.xml.rels');
  if (!file) throw new Error('document.xml.rels not found');
  return file.async('string');
}

function extractRIdReferences(xml: string): string[] {
  const matches = xml.match(/r:id="(rId\d+)"|r:embed="(rId\d+)"/g) || [];
  return matches.map((m) => {
    const match = m.match(/rId\d+/);
    return match ? match[0] : '';
  }).filter(Boolean);
}

function extractDefinedRIds(relsXml: string): string[] {
  const matches = relsXml.match(/Id="(rId\d+)"/g) || [];
  return matches.map((m) => {
    const match = m.match(/rId\d+/);
    return match ? match[0] : '';
  }).filter(Boolean);
}

describe('WmlComparer Regression Tests', () => {
  describe('Image Preservation', () => {
    it('preserves drawing elements when comparing documents with images', async () => {
      const doc1 = await loadTestFile('WC/WC013-Image-Before.docx');
      const doc2 = await loadTestFile('WC/WC013-Image-After.docx');

      if (!doc1 || !doc2) {
        console.warn('Skipping: test fixtures not found');
        return;
      }

      const result = await compareDocuments(doc1, doc2);
      const documentXml = await extractDocumentXml(result.document);

      const hasDrawingElements = documentXml.includes('<w:drawing>');
      const hasDrawingTokens = />\s*DRAWING_[^<]+</.test(documentXml);

      expect(hasDrawingElements).toBe(true);
      expect(hasDrawingTokens).toBe(false);
    });

    it('does not leave DRAWING_ tokens as text in output', async () => {
      const doc1 = await loadTestFile('WC/WC016-Para-Image-Para.docx');
      const doc2 = await loadTestFile('WC/WC016-Para-Image-Para-w-Deleted-Image.docx');

      if (!doc1 || !doc2) {
        console.warn('Skipping: test fixtures not found');
        return;
      }

      const result = await compareDocuments(doc1, doc2);
      const documentXml = await extractDocumentXml(result.document);

      const drawingTokenMatches = documentXml.match(/DRAWING_\w+/g) || [];
      const textDrawingTokens = drawingTokenMatches.filter((token) => {
        const regex = new RegExp(`<w:t[^>]*>[^<]*${token}[^<]*</w:t>`);
        return regex.test(documentXml);
      });

      expect(textDrawingTokens.length).toBe(0);
    });

    it('preserves images in table comparisons', async () => {
      const doc1 = await loadTestFile('WC/WC023-Table-4-Row-Image-Before.docx');
      const doc2 = await loadTestFile('WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx');

      if (!doc1 || !doc2) {
        console.warn('Skipping: test fixtures not found');
        return;
      }

      const result = await compareDocuments(doc1, doc2);
      const documentXml = await extractDocumentXml(result.document);

      const hasDrawingTokensAsText = />\s*DRAWING_[^<]+</.test(documentXml);
      expect(hasDrawingTokensAsText).toBe(false);
    });
  });

  describe('Relationship ID Validity', () => {
    it('all rId references exist in relationships file', async () => {
      const doc1 = await loadTestFile('WC/WC013-Image-Before.docx');
      const doc2 = await loadTestFile('WC/WC013-Image-After.docx');

      if (!doc1 || !doc2) {
        console.warn('Skipping: test fixtures not found');
        return;
      }

      const result = await compareDocuments(doc1, doc2);
      const documentXml = await extractDocumentXml(result.document);
      const relsXml = await extractRelationships(result.document);

      const referencedRIds = extractRIdReferences(documentXml);
      const definedRIds = new Set(extractDefinedRIds(relsXml));

      const orphanRIds = referencedRIds.filter((rId) => !definedRIds.has(rId));

      expect(orphanRIds).toEqual([]);
    });

    it('no orphan rId references after comparing documents with headers/footers', async () => {
      const doc1 = await loadTestFile('WC/WC002-Unmodified.docx');
      const doc2 = await loadTestFile('WC/WC002-DiffInMiddle.docx');

      if (!doc1 || !doc2) {
        console.warn('Skipping: test fixtures not found');
        return;
      }

      const result = await compareDocuments(doc1, doc2);
      const documentXml = await extractDocumentXml(result.document);
      const relsXml = await extractRelationships(result.document);

      const referencedRIds = extractRIdReferences(documentXml);
      const definedRIds = new Set(extractDefinedRIds(relsXml));

      const orphanRIds = referencedRIds.filter((rId) => !definedRIds.has(rId));

      if (orphanRIds.length > 0) {
        console.error('Orphan rId references found:', orphanRIds);
      }

      expect(orphanRIds).toEqual([]);
    });

    it('sectPr in deleted content does not cause orphan rIds', async () => {
      const doc1 = await loadTestFile('WC/WC001-Digits.docx');
      const doc2 = await loadTestFile('WC/WC001-Digits-Deleted-Paragraph.docx');

      if (!doc1 || !doc2) {
        console.warn('Skipping: test fixtures not found');
        return;
      }

      const result = await compareDocuments(doc1, doc2);
      const documentXml = await extractDocumentXml(result.document);
      const relsXml = await extractRelationships(result.document);

      const referencedRIds = extractRIdReferences(documentXml);
      const definedRIds = new Set(extractDefinedRIds(relsXml));

      const orphanRIds = referencedRIds.filter((rId) => !definedRIds.has(rId));

      expect(orphanRIds).toEqual([]);
    });
  });

  describe('Document Validity', () => {
    it('output document can be parsed as valid ZIP', async () => {
      const doc1 = await loadTestFile('WC/WC001-Digits.docx');
      const doc2 = await loadTestFile('WC/WC001-Digits-Mod.docx');

      if (!doc1 || !doc2) {
        console.warn('Skipping: test fixtures not found');
        return;
      }

      const result = await compareDocuments(doc1, doc2);

      const zip = await JSZip.loadAsync(result.document);
      const documentXml = zip.file('word/document.xml');
      const relsFile = zip.file('word/_rels/document.xml.rels');
      const contentTypes = zip.file('[Content_Types].xml');

      expect(documentXml).not.toBeNull();
      expect(relsFile).not.toBeNull();
      expect(contentTypes).not.toBeNull();
    });

    it('document.xml contains valid XML structure', async () => {
      const doc1 = await loadTestFile('WC/WC013-Image-Before.docx');
      const doc2 = await loadTestFile('WC/WC013-Image-After2.docx');

      if (!doc1 || !doc2) {
        console.warn('Skipping: test fixtures not found');
        return;
      }

      const result = await compareDocuments(doc1, doc2);
      const documentXml = await extractDocumentXml(result.document);

      expect(documentXml).toContain('<?xml');
      expect(documentXml).toContain('<w:document');
      expect(documentXml).toContain('<w:body>');
      expect(documentXml).toContain('</w:body>');
      expect(documentXml).toContain('</w:document>');
    });
  });
});
