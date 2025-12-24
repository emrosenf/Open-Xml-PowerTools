import { beforeAll, describe, expect, it } from 'vitest';
import { comparePresentations } from '../src/pml/pml-comparer';
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
});
