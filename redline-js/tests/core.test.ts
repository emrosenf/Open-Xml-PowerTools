/**
 * Core infrastructure tests
 *
 * Verifies that XML parsing, ZIP handling, and document loading work correctly.
 */

import { describe, it, expect } from 'vitest';
import { readFile } from 'fs/promises';
import { join } from 'path';

import {
  parseXml,
  buildXml,
  getTagName,
  getChildren,
  getTextContent,
  openPackage,
  getPartAsXml,
  hashString,
  W,
} from '../src/core';

import {
  loadWordDocument,
  extractText,
  extractParagraphs,
  getDocumentBody,
} from '../src/wml/document';

const TEST_FILES_DIR = join(__dirname, '../../TestFiles');

describe('XML Utilities', () => {
  it('parses simple XML', () => {
    const xml = '<root><child>text</child></root>';
    const nodes = parseXml(xml);

    expect(nodes).toHaveLength(1);
    expect(getTagName(nodes[0])).toBe('root');
  });

  it('round-trips XML', () => {
    const xml = '<root><child attr="value">text</child></root>';
    const nodes = parseXml(xml);
    const rebuilt = buildXml(nodes);

    expect(rebuilt).toContain('<root>');
    expect(rebuilt).toContain('<child');
    expect(rebuilt).toContain('attr="value"');
    expect(rebuilt).toContain('text');
  });

  it('extracts text content', () => {
    const xml = '<root><a>Hello</a><b> World</b></root>';
    const nodes = parseXml(xml);
    const text = getTextContent(nodes[0]);

    expect(text).toBe('Hello World');
  });
});

describe('Hash Utilities', () => {
  it('produces consistent hashes', () => {
    const hash1 = hashString('test content');
    const hash2 = hashString('test content');

    expect(hash1).toBe(hash2);
    expect(hash1).toHaveLength(64); // SHA-256 produces 64 hex chars
  });

  it('produces different hashes for different content', () => {
    const hash1 = hashString('content a');
    const hash2 = hashString('content b');

    expect(hash1).not.toBe(hash2);
  });
});

describe('Namespace Constants', () => {
  it('has correct Word namespace', () => {
    expect(W).toBe('http://schemas.openxmlformats.org/wordprocessingml/2006/main');
  });
});

describe('Package Handling', () => {
  it('opens a valid Word document', async () => {
    const docPath = join(TEST_FILES_DIR, 'CA/CA001-Plain.docx');
    let data: Buffer;

    try {
      data = await readFile(docPath);
    } catch {
      console.warn('Skipping: test fixture not found');
      return;
    }

    const pkg = await openPackage(data);

    expect(pkg.fileType).toBe('word');
    expect(pkg.contentTypes.overrides.size).toBeGreaterThan(0);
  });

  it('reads document.xml from package', async () => {
    const docPath = join(TEST_FILES_DIR, 'CA/CA001-Plain.docx');
    let data: Buffer;

    try {
      data = await readFile(docPath);
    } catch {
      console.warn('Skipping: test fixture not found');
      return;
    }

    const pkg = await openPackage(data);
    const docXml = await getPartAsXml(pkg, 'word/document.xml');

    expect(docXml).not.toBeNull();
    expect(docXml!.length).toBeGreaterThan(0);
  });
});

describe('Word Document Loading', () => {
  it('loads a Word document', async () => {
    const docPath = join(TEST_FILES_DIR, 'CA/CA001-Plain.docx');
    let data: Buffer;

    try {
      data = await readFile(docPath);
    } catch {
      console.warn('Skipping: test fixture not found');
      return;
    }

    const doc = await loadWordDocument(data);

    expect(doc.mainDocument).toBeDefined();
    expect(doc.mainDocument.length).toBeGreaterThan(0);
  });

  it('extracts document body', async () => {
    const docPath = join(TEST_FILES_DIR, 'CA/CA001-Plain.docx');
    let data: Buffer;

    try {
      data = await readFile(docPath);
    } catch {
      console.warn('Skipping: test fixture not found');
      return;
    }

    const doc = await loadWordDocument(data);
    const body = getDocumentBody(doc);

    expect(body).not.toBeNull();
    expect(getTagName(body!)).toBe('w:body');
  });

  it('extracts text from document', async () => {
    const docPath = join(TEST_FILES_DIR, 'CA/CA001-Plain.docx');
    let data: Buffer;

    try {
      data = await readFile(docPath);
    } catch {
      console.warn('Skipping: test fixture not found');
      return;
    }

    const doc = await loadWordDocument(data);
    const text = extractText(doc);

    expect(text.length).toBeGreaterThan(0);
    console.log('Extracted text:', text.substring(0, 100));
  });

  it('extracts paragraphs from document', async () => {
    const docPath = join(TEST_FILES_DIR, 'CA/CA001-Plain.docx');
    let data: Buffer;

    try {
      data = await readFile(docPath);
    } catch {
      console.warn('Skipping: test fixture not found');
      return;
    }

    const doc = await loadWordDocument(data);
    const paragraphs = extractParagraphs(doc);

    expect(paragraphs.length).toBeGreaterThan(0);
    console.log(`Found ${paragraphs.length} paragraphs`);
  });
});
