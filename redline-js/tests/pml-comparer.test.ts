import { beforeAll, describe, expect, it } from 'vitest';
import { comparePresentations, buildChangeList } from '../src/pml/pml-comparer';
import { PmlChangeType, TextChangeType, type PmlChange, PmlComparisonResult } from '../src/pml/types';
import {
  loadDocument,
  loadGoldenFile,
  loadManifest,
  skipIfMissingFixtures,
  type PmlTestCase,
  type TestManifest,
} from './setup';

const defaultSettings = {};

describe('PmlComparer', () => {
  let manifest: TestManifest | null = null;

  beforeAll(async () => {
    try {
      manifest = await loadManifest();
    } catch {
      manifest = null;
    }
  });

  it('loads the golden manifest', () => {
    if (!manifest) {
      return;
    }
    expect(Array.isArray(manifest.pmlTests)).toBe(true);
  });

  describe('Golden File Validation', () => {
    it('matches golden outputs', async () => {
      if (!manifest) {
        return;
      }

      for (const testCase of manifest.pmlTests as PmlTestCase[]) {
        if (await skipIfMissingFixtures(testCase.source1, testCase.source2)) {
          continue;
        }

        const doc1 = await loadDocument(testCase.source1);
        const doc2 = await loadDocument(testCase.source2);
        const result = await comparePresentations(doc1, doc2, defaultSettings);
        const actual = JSON.parse(result.toJson());

        const goldenPath = testCase.outputFile ?? `pml/${testCase.testId}.json`;
        const expectedBuffer = await loadGoldenFile(goldenPath);
        const expected = JSON.parse(expectedBuffer.toString());

        expect(actual).toEqual(expected);
      }
    });
  });

  describe('buildChangeList', () => {
    it('generates preview text for text changes', () => {
      const result = new PmlComparisonResult();
      result.changes.push({
        changeType: PmlChangeType.TextChanged,
        slideIndex: 1,
        shapeName: 'Title',
        shapeId: 'shape-1',
        textChanges: [
          { type: TextChangeType.Delete, paragraphIndex: 0, runIndex: 0, oldText: 'Hello' },
          { type: TextChangeType.Insert, paragraphIndex: 0, runIndex: 1, newText: 'World' },
        ],
      });

      const items = buildChangeList(result);

      expect(items).toHaveLength(1);
      expect(items[0].previewText).toBe('-"Hello" +"World"');
      expect(items[0].wordCount).toEqual({ deleted: 1, inserted: 1 });
    });

    it('generates preview text for replace changes', () => {
      const result = new PmlComparisonResult();
      result.changes.push({
        changeType: PmlChangeType.TextChanged,
        slideIndex: 1,
        shapeName: 'Content',
        shapeId: 'shape-2',
        textChanges: [
          { type: TextChangeType.Replace, paragraphIndex: 0, runIndex: 0, oldText: 'old text', newText: 'new text' },
        ],
      });

      const items = buildChangeList(result);

      expect(items).toHaveLength(1);
      expect(items[0].previewText).toBe('"old text" → "new text"');
      expect(items[0].wordCount).toEqual({ deleted: 2, inserted: 2 });
    });

    it('generates preview text from oldValue/newValue', () => {
      const result = new PmlComparisonResult();
      result.changes.push({
        changeType: PmlChangeType.SlideNotesChanged,
        slideIndex: 1,
        oldValue: 'Original notes text',
        newValue: 'Updated notes',
      });

      const items = buildChangeList(result);

      expect(items).toHaveLength(1);
      expect(items[0].previewText).toBe('"Original notes text" → "Updated notes"');
      expect(items[0].wordCount).toEqual({ deleted: 3, inserted: 2 });
    });

    it('truncates long preview text', () => {
      const longText = 'This is a very long text that should be truncated because it exceeds the maximum length allowed';
      const result = new PmlComparisonResult();
      result.changes.push({
        changeType: PmlChangeType.TextChanged,
        slideIndex: 1,
        shapeName: 'Body',
        newValue: longText,
      });

      const items = buildChangeList(result, { maxPreviewLength: 50 });

      expect(items).toHaveLength(1);
      expect(items[0].previewText!.length).toBeLessThanOrEqual(50);
      expect(items[0].previewText).toContain('...');
    });

    it('generates anchors for navigation', () => {
      const result = new PmlComparisonResult();
      result.changes.push({
        changeType: PmlChangeType.ShapeInserted,
        slideIndex: 3,
        shapeName: 'New Shape',
        shapeId: 'shape-42',
      });

      const items = buildChangeList(result);

      expect(items).toHaveLength(1);
      expect(items[0].anchor).toBe('slide-3#shape-shape-42');
    });

    it('groups changes by slide when option is set', () => {
      const result = new PmlComparisonResult();
      result.changes.push(
        { changeType: PmlChangeType.TextChanged, slideIndex: 1, shapeName: 'Title' },
        { changeType: PmlChangeType.ShapeInserted, slideIndex: 1, shapeName: 'Box' },
        { changeType: PmlChangeType.TextChanged, slideIndex: 2, shapeName: 'Title' },
      );

      const items = buildChangeList(result, { groupBySlide: true });

      const slideHeaders = items.filter(i => i.id.startsWith('pml-slide-'));
      expect(slideHeaders.length).toBeGreaterThanOrEqual(2);
    });

    it('handles changes without text content', () => {
      const result = new PmlComparisonResult();
      result.changes.push({
        changeType: PmlChangeType.ShapeMoved,
        slideIndex: 1,
        shapeName: 'Image',
        shapeId: 'shape-1',
        oldX: 100,
        oldY: 100,
        newX: 200,
        newY: 200,
      });

      const items = buildChangeList(result);

      expect(items).toHaveLength(1);
      expect(items[0].previewText).toBeUndefined();
      expect(items[0].wordCount).toBeUndefined();
    });
  });
});
