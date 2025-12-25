/**
 * Tests for LCS (Longest Common Subsequence) algorithm
 */

import { describe, it, expect } from 'vitest';
import {
  findLongestMatch,
  computeCorrelation,
  flattenCorrelation,
  diffText,
  CorrelationStatus,
  type Hashable,
} from '../src/core/lcs';

// Helper to create hashable items from strings
function h(s: string): Hashable {
  return { hash: s };
}

function hs(strings: string[]): Hashable[] {
  return strings.map(h);
}

describe('findLongestMatch', () => {
  it('finds exact match', () => {
    const items1 = hs(['a', 'b', 'c']);
    const items2 = hs(['a', 'b', 'c']);

    const match = findLongestMatch(items1, items2);

    expect(match).toEqual({ i1: 0, i2: 0, length: 3 });
  });

  it('finds match at different positions', () => {
    const items1 = hs(['x', 'a', 'b', 'c', 'y']);
    const items2 = hs(['z', 'a', 'b', 'c', 'w']);

    const match = findLongestMatch(items1, items2);

    expect(match).toEqual({ i1: 1, i2: 1, length: 3 });
  });

  it('returns null for no common content', () => {
    const items1 = hs(['a', 'b', 'c']);
    const items2 = hs(['x', 'y', 'z']);

    const match = findLongestMatch(items1, items2);

    expect(match).toBeNull();
  });

  it('returns null for empty arrays', () => {
    const match1 = findLongestMatch([], []);
    const match2 = findLongestMatch(hs(['a']), []);
    const match3 = findLongestMatch([], hs(['a']));

    expect(match1).toBeNull();
    expect(match2).toBeNull();
    expect(match3).toBeNull();
  });

  it('respects minMatchLength', () => {
    const items1 = hs(['a', 'b', 'c']);
    const items2 = hs(['x', 'a', 'y']);

    const match1 = findLongestMatch(items1, items2, { minMatchLength: 1 });
    const match2 = findLongestMatch(items1, items2, { minMatchLength: 2 });

    expect(match1).toEqual({ i1: 0, i2: 1, length: 1 });
    expect(match2).toBeNull();
  });

  it('respects detailThreshold', () => {
    const items1 = hs(['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j']); // 10 items
    const items2 = hs(['x', 'a', 'y', 'z', 'w', 'v', 'u', 't', 's', 'r']); // 10 items, 1 match

    const match1 = findLongestMatch(items1, items2, { detailThreshold: 0.05 }); // 5% - should match
    const match2 = findLongestMatch(items1, items2, { detailThreshold: 0.15 }); // 15% - should not match

    expect(match1).not.toBeNull();
    expect(match2).toBeNull();
  });

  it('respects shouldSkipAsAnchor', () => {
    const items1 = hs(['skip', 'a', 'b']);
    const items2 = hs(['skip', 'a', 'b']);

    const match = findLongestMatch(items1, items2, {
      shouldSkipAsAnchor: (item) => item.hash === 'skip',
    });

    // Should skip the 'skip' item and start from 'a'
    expect(match).toEqual({ i1: 1, i2: 1, length: 2 });
  });
});

describe('computeCorrelation', () => {
  it('handles identical arrays', () => {
    const items1 = hs(['a', 'b', 'c']);
    const items2 = hs(['a', 'b', 'c']);

    const result = computeCorrelation(items1, items2);

    expect(result).toHaveLength(1);
    expect(result[0].status).toBe(CorrelationStatus.Equal);
    expect(result[0].items1?.map((i) => i.hash)).toEqual(['a', 'b', 'c']);
    expect(result[0].items2?.map((i) => i.hash)).toEqual(['a', 'b', 'c']);
  });

  it('handles completely different arrays', () => {
    const items1 = hs(['a', 'b']);
    const items2 = hs(['x', 'y']);

    const result = computeCorrelation(items1, items2);

    expect(result).toHaveLength(2);
    expect(result[0].status).toBe(CorrelationStatus.Deleted);
    expect(result[1].status).toBe(CorrelationStatus.Inserted);
  });

  it('handles empty arrays', () => {
    expect(computeCorrelation([], [])).toEqual([]);

    const inserted = computeCorrelation([], hs(['a']));
    expect(inserted).toHaveLength(1);
    expect(inserted[0].status).toBe(CorrelationStatus.Inserted);

    const deleted = computeCorrelation(hs(['a']), []);
    expect(deleted).toHaveLength(1);
    expect(deleted[0].status).toBe(CorrelationStatus.Deleted);
  });

  it('handles insertion at beginning', () => {
    const items1 = hs(['a', 'b']);
    const items2 = hs(['x', 'a', 'b']);

    const result = computeCorrelation(items1, items2);

    expect(result).toHaveLength(2);
    expect(result[0].status).toBe(CorrelationStatus.Inserted);
    expect(result[0].items2?.map((i) => i.hash)).toEqual(['x']);
    expect(result[1].status).toBe(CorrelationStatus.Equal);
    expect(result[1].items1?.map((i) => i.hash)).toEqual(['a', 'b']);
  });

  it('handles deletion at beginning', () => {
    const items1 = hs(['x', 'a', 'b']);
    const items2 = hs(['a', 'b']);

    const result = computeCorrelation(items1, items2);

    expect(result).toHaveLength(2);
    expect(result[0].status).toBe(CorrelationStatus.Deleted);
    expect(result[0].items1?.map((i) => i.hash)).toEqual(['x']);
    expect(result[1].status).toBe(CorrelationStatus.Equal);
  });

  it('handles insertion at end', () => {
    const items1 = hs(['a', 'b']);
    const items2 = hs(['a', 'b', 'x']);

    const result = computeCorrelation(items1, items2);

    expect(result).toHaveLength(2);
    expect(result[0].status).toBe(CorrelationStatus.Equal);
    expect(result[1].status).toBe(CorrelationStatus.Inserted);
    expect(result[1].items2?.map((i) => i.hash)).toEqual(['x']);
  });

  it('handles deletion at end', () => {
    const items1 = hs(['a', 'b', 'x']);
    const items2 = hs(['a', 'b']);

    const result = computeCorrelation(items1, items2);

    expect(result).toHaveLength(2);
    expect(result[0].status).toBe(CorrelationStatus.Equal);
    expect(result[1].status).toBe(CorrelationStatus.Deleted);
    expect(result[1].items1?.map((i) => i.hash)).toEqual(['x']);
  });

  it('handles replacement in middle', () => {
    const items1 = hs(['a', 'b', 'c', 'd']);
    const items2 = hs(['a', 'x', 'y', 'd']);

    const result = computeCorrelation(items1, items2);

    // Should find 'a' match, then 'd' match
    // Result: Equal(a), Deleted(b,c), Inserted(x,y), Equal(d)
    const statuses = result.map((r) => r.status);
    expect(statuses).toContain(CorrelationStatus.Equal);
    expect(statuses).toContain(CorrelationStatus.Deleted);
    expect(statuses).toContain(CorrelationStatus.Inserted);
  });

  it('handles complex diff', () => {
    const items1 = hs(['The', 'quick', 'brown', 'fox']);
    const items2 = hs(['The', 'slow', 'brown', 'dog']);

    const result = computeCorrelation(items1, items2);

    // Should find matches for 'The' and 'brown'
    const equalParts = result.filter((r) => r.status === CorrelationStatus.Equal);
    expect(equalParts.length).toBeGreaterThanOrEqual(1);
  });
});

describe('flattenCorrelation', () => {
  it('merges adjacent sequences of same status', () => {
    const sequences = [
      { status: CorrelationStatus.Deleted, items1: hs(['a']), items2: null },
      { status: CorrelationStatus.Deleted, items1: hs(['b']), items2: null },
      { status: CorrelationStatus.Equal, items1: hs(['c']), items2: hs(['c']) },
    ];

    const result = flattenCorrelation(sequences);

    expect(result).toHaveLength(2);
    expect(result[0].status).toBe(CorrelationStatus.Deleted);
    expect(result[0].items1?.map((i) => i.hash)).toEqual(['a', 'b']);
    expect(result[1].status).toBe(CorrelationStatus.Equal);
  });

  it('handles empty input', () => {
    expect(flattenCorrelation([])).toEqual([]);
  });

  it('handles single item', () => {
    const sequences = [{ status: CorrelationStatus.Equal, items1: hs(['a']), items2: hs(['a']) }];

    const result = flattenCorrelation(sequences);

    expect(result).toHaveLength(1);
  });
});

describe('diffText', () => {
  it('diffs identical text', () => {
    const result = diffText('hello', 'hello');

    expect(result).toHaveLength(1);
    expect(result[0]).toEqual({ type: 'equal', value: 'hello' });
  });

  it('diffs completely different text', () => {
    const result = diffText('abc', 'xyz');

    expect(result.some((r) => r.type === 'delete')).toBe(true);
    expect(result.some((r) => r.type === 'insert')).toBe(true);
  });

  it('diffs text with insertion', () => {
    const result = diffText('ac', 'abc');

    // Should show 'a' equal, 'b' inserted, 'c' equal
    const types = result.map((r) => r.type);
    expect(types).toContain('equal');
    expect(types).toContain('insert');
  });

  it('diffs text with deletion', () => {
    const result = diffText('abc', 'ac');

    const types = result.map((r) => r.type);
    expect(types).toContain('equal');
    expect(types).toContain('delete');
  });

  it('diffs by words when pattern provided', () => {
    const result = diffText('The quick fox', 'The slow fox', /\s+/);

    expect(result).toEqual([
      { type: 'equal', value: 'The' },
      { type: 'delete', value: 'quick' },
      { type: 'insert', value: 'slow' },
      { type: 'equal', value: 'fox' },
    ]);
  });

  it('handles empty strings', () => {
    expect(diffText('', '')).toEqual([]);

    const insert = diffText('', 'abc');
    expect(insert).toHaveLength(1);
    expect(insert[0].type).toBe('insert');

    const del = diffText('abc', '');
    expect(del).toHaveLength(1);
    expect(del[0].type).toBe('delete');
  });
});
