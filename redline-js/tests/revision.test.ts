/**
 * Tests for revision markup generation
 */

import { describe, it, expect, beforeEach } from 'vitest';
import {
  createInsertion,
  createDeletion,
  createRun,
  createParagraph,
  createRunPropertyChange,
  createParagraphPropertyChange,
  wrapWithRevision,
  isRevisionElement,
  isInsertion,
  isDeletion,
  countRevisions,
  resetRevisionIdCounter,
  type RevisionSettings,
} from '../src/wml/revision';
import { buildXml, getTagName, getChildren, getAttributes } from '../src/core/xml';

const TEST_SETTINGS: RevisionSettings = {
  author: 'Test Author',
  dateTime: '2024-01-15T10:30:00Z',
};

describe('Revision Markup', () => {
  beforeEach(() => {
    resetRevisionIdCounter();
  });

  describe('createRun', () => {
    it('creates a basic run with text', () => {
      const run = createRun('Hello');

      expect(getTagName(run)).toBe('w:r');
      const children = getChildren(run);
      expect(children.length).toBe(1);
      expect(getTagName(children[0])).toBe('w:t');
    });

    it('preserves whitespace for text with leading/trailing spaces', () => {
      const run = createRun(' Hello ');

      const children = getChildren(run);
      const textNode = children[0];
      const attrs = textNode[':@'] as Record<string, string>;
      expect(attrs?.['@_xml:space']).toBe('preserve');
    });

    it('includes run properties when provided', () => {
      const props = { 'w:rPr': [{ 'w:b': [] }] };
      const run = createRun('Bold', props);

      const children = getChildren(run);
      expect(children.length).toBe(2);
      expect(getTagName(children[0])).toBe('w:rPr');
      expect(getTagName(children[1])).toBe('w:t');
    });
  });

  describe('createParagraph', () => {
    it('creates a paragraph with runs', () => {
      const run1 = createRun('Hello');
      const run2 = createRun('World');
      const para = createParagraph([run1, run2]);

      expect(getTagName(para)).toBe('w:p');
      const children = getChildren(para);
      expect(children.length).toBe(2);
    });

    it('includes paragraph properties when provided', () => {
      const props = { 'w:pPr': [{ 'w:jc': [], ':@': { '@_w:val': 'center' } }] };
      const run = createRun('Centered');
      const para = createParagraph([run], props);

      const children = getChildren(para);
      expect(children.length).toBe(2);
      expect(getTagName(children[0])).toBe('w:pPr');
    });
  });

  describe('createInsertion', () => {
    it('creates a w:ins element', () => {
      const run = createRun('Inserted text');
      const ins = createInsertion(run, TEST_SETTINGS);

      expect(getTagName(ins)).toBe('w:ins');
    });

    it('includes required attributes', () => {
      const run = createRun('Text');
      const ins = createInsertion(run, TEST_SETTINGS);

      const attrs = ins[':@'] as Record<string, string>;
      expect(attrs['@_w:author']).toBe('Test Author');
      expect(attrs['@_w:date']).toBe('2024-01-15T10:30:00Z');
      expect(attrs['@_w:id']).toBeDefined();
    });

    it('wraps content as children', () => {
      const run = createRun('Text');
      const ins = createInsertion(run, TEST_SETTINGS);

      const children = getChildren(ins);
      expect(children.length).toBe(1);
      expect(getTagName(children[0])).toBe('w:r');
    });

    it('handles multiple content items', () => {
      const run1 = createRun('One');
      const run2 = createRun('Two');
      const ins = createInsertion([run1, run2], TEST_SETTINGS);

      const children = getChildren(ins);
      expect(children.length).toBe(2);
    });

    it('assigns unique IDs', () => {
      const ins1 = createInsertion(createRun('A'), TEST_SETTINGS);
      const ins2 = createInsertion(createRun('B'), TEST_SETTINGS);

      const id1 = (ins1[':@'] as Record<string, string>)['@_w:id'];
      const id2 = (ins2[':@'] as Record<string, string>)['@_w:id'];

      expect(id1).not.toBe(id2);
    });
  });

  describe('createDeletion', () => {
    it('creates a w:del element', () => {
      const run = createRun('Deleted text');
      const del = createDeletion(run, TEST_SETTINGS);

      expect(getTagName(del)).toBe('w:del');
    });

    it('includes required attributes', () => {
      const run = createRun('Text');
      const del = createDeletion(run, TEST_SETTINGS);

      const attrs = del[':@'] as Record<string, string>;
      expect(attrs['@_w:author']).toBe('Test Author');
      expect(attrs['@_w:date']).toBe('2024-01-15T10:30:00Z');
      expect(attrs['@_w:id']).toBeDefined();
    });

    it('converts w:t to w:delText', () => {
      const run = createRun('Deleted');
      const del = createDeletion(run, TEST_SETTINGS);

      const xml = buildXml(del);
      expect(xml).toContain('w:delText');
      expect(xml).not.toContain('<w:t>');
    });
  });

  describe('createRunPropertyChange', () => {
    it('creates a w:rPrChange element', () => {
      const oldProps = { 'w:rPr': [{ 'w:b': [] }] };
      const change = createRunPropertyChange(oldProps, TEST_SETTINGS);

      expect(getTagName(change)).toBe('w:rPrChange');
    });

    it('includes tracking attributes', () => {
      const oldProps = { 'w:rPr': [] };
      const change = createRunPropertyChange(oldProps, TEST_SETTINGS);

      const attrs = change[':@'] as Record<string, string>;
      expect(attrs['@_w:author']).toBe('Test Author');
      expect(attrs['@_w:id']).toBeDefined();
    });
  });

  describe('createParagraphPropertyChange', () => {
    it('creates a w:pPrChange element', () => {
      const oldProps = { 'w:pPr': [{ 'w:jc': [], ':@': { '@_w:val': 'left' } }] };
      const change = createParagraphPropertyChange(oldProps, TEST_SETTINGS);

      expect(getTagName(change)).toBe('w:pPrChange');
    });
  });

  describe('wrapWithRevision', () => {
    it('creates insertion for inserted status', () => {
      const run = createRun('Text');
      const result = wrapWithRevision(run, 'inserted', TEST_SETTINGS);

      expect(getTagName(result)).toBe('w:ins');
    });

    it('creates deletion for deleted status', () => {
      const run = createRun('Text');
      const result = wrapWithRevision(run, 'deleted', TEST_SETTINGS);

      expect(getTagName(result)).toBe('w:del');
    });
  });

  describe('revision detection', () => {
    it('isRevisionElement identifies w:ins', () => {
      const ins = createInsertion(createRun('Text'), TEST_SETTINGS);
      expect(isRevisionElement(ins)).toBe(true);
    });

    it('isRevisionElement identifies w:del', () => {
      const del = createDeletion(createRun('Text'), TEST_SETTINGS);
      expect(isRevisionElement(del)).toBe(true);
    });

    it('isRevisionElement returns false for other elements', () => {
      const run = createRun('Text');
      expect(isRevisionElement(run)).toBe(false);
    });

    it('isInsertion works correctly', () => {
      const ins = createInsertion(createRun('Text'), TEST_SETTINGS);
      const del = createDeletion(createRun('Text'), TEST_SETTINGS);

      expect(isInsertion(ins)).toBe(true);
      expect(isInsertion(del)).toBe(false);
    });

    it('isDeletion works correctly', () => {
      const ins = createInsertion(createRun('Text'), TEST_SETTINGS);
      const del = createDeletion(createRun('Text'), TEST_SETTINGS);

      expect(isDeletion(ins)).toBe(false);
      expect(isDeletion(del)).toBe(true);
    });
  });

  describe('countRevisions', () => {
    it('counts insertions and deletions', () => {
      const ins1 = createInsertion(createRun('A'), TEST_SETTINGS);
      const ins2 = createInsertion(createRun('B'), TEST_SETTINGS);
      const del = createDeletion(createRun('C'), TEST_SETTINGS);

      const para = createParagraph([ins1, ins2, del]);
      const counts = countRevisions(para);

      expect(counts.insertions).toBe(2);
      expect(counts.deletions).toBe(1);
      expect(counts.total).toBe(3);
    });

    it('handles empty content', () => {
      const para = createParagraph([createRun('Plain text')]);
      const counts = countRevisions(para);

      expect(counts.insertions).toBe(0);
      expect(counts.deletions).toBe(0);
      expect(counts.total).toBe(0);
    });

    it('handles array input', () => {
      const ins = createInsertion(createRun('A'), TEST_SETTINGS);
      const del = createDeletion(createRun('B'), TEST_SETTINGS);

      const counts = countRevisions([ins, del]);

      expect(counts.insertions).toBe(1);
      expect(counts.deletions).toBe(1);
    });
  });
});

describe('XML Output', () => {
  beforeEach(() => {
    resetRevisionIdCounter();
  });

  it('produces valid insertion XML', () => {
    const run = createRun('New text');
    const ins = createInsertion(run, TEST_SETTINGS);
    const xml = buildXml(ins);

    expect(xml).toContain('<w:ins');
    expect(xml).toContain('w:author="Test Author"');
    expect(xml).toContain('w:id="1"');
    expect(xml).toContain('<w:r>');
    expect(xml).toContain('<w:t>New text</w:t>');
  });

  it('produces valid deletion XML', () => {
    const run = createRun('Old text');
    const del = createDeletion(run, TEST_SETTINGS);
    const xml = buildXml(del);

    expect(xml).toContain('<w:del');
    expect(xml).toContain('w:author="Test Author"');
    expect(xml).toContain('<w:r>');
    expect(xml).toContain('<w:delText>Old text</w:delText>');
  });
});
