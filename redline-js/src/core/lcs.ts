/**
 * Longest Common Subsequence (LCS) algorithm for document comparison
 *
 * This is a port of the LCS algorithm from Open-Xml-PowerTools WmlComparer.
 * The algorithm finds the longest common contiguous subsequence between two
 * arrays of comparison units, then recursively processes the non-matching
 * portions.
 *
 * Key insight: This is NOT the classic LCS that finds non-contiguous matches.
 * This finds the longest contiguous matching run, which is simpler but requires
 * recursive application to find all matches.
 */

/**
 * Correlation status indicating how content relates between documents
 */
export enum CorrelationStatus {
  /** Content appears in both documents at this position */
  Equal = 'Equal',
  /** Content was deleted from the original document */
  Deleted = 'Deleted',
  /** Content was inserted in the modified document */
  Inserted = 'Inserted',
  /** Not yet determined - needs further processing */
  Unknown = 'Unknown',
}

/**
 * Interface for items that can be compared using LCS
 * Items must have a hash property for equality comparison
 */
export interface Hashable {
  hash: string;
}

/**
 * A correlated sequence showing how portions of two arrays relate
 */
export interface CorrelatedSequence<T extends Hashable> {
  status: CorrelationStatus;
  /** Items from the first (original) array, null if inserted */
  items1: T[] | null;
  /** Items from the second (modified) array, null if deleted */
  items2: T[] | null;
}

/**
 * Settings for the LCS algorithm
 */
export interface LcsSettings {
  /**
   * Minimum length for a match to be considered valid.
   * Helps avoid matching insignificant content like single spaces.
   * Default: 1
   */
  minMatchLength?: number;

  /**
   * Threshold (0-1) for minimum match ratio relative to total length.
   * If the longest match is less than this ratio of the max length,
   * it's considered not a real match. Default: 0.0 (any match accepted)
   */
  detailThreshold?: number;

  /**
   * Optional predicate to skip certain items from being match anchors.
   * Useful for skipping paragraph marks, whitespace-only content, etc.
   */
  shouldSkipAsAnchor?: (item: Hashable) => boolean;
}

const DEFAULT_SETTINGS: Required<LcsSettings> = {
  minMatchLength: 1,
  detailThreshold: 0.0,
  shouldSkipAsAnchor: () => false,
};

/**
 * Find the longest common contiguous subsequence between two arrays.
 *
 * This is an O(n*m) algorithm where n and m are the lengths of the input arrays.
 * It finds the longest run of consecutive items with matching hashes.
 *
 * @returns Object with start indices and length, or null if no match found
 */
export function findLongestMatch<T extends Hashable>(
  items1: T[],
  items2: T[],
  settings: LcsSettings = {}
): { i1: number; i2: number; length: number } | null {
  const opts = { ...DEFAULT_SETTINGS, ...settings };

  let bestLength = 0;
  let bestI1 = -1;
  let bestI2 = -1;

  // Optimization: don't search positions where we can't possibly find
  // a longer match than what we already have
  for (let i1 = 0; i1 < items1.length - bestLength; i1++) {
    for (let i2 = 0; i2 < items2.length - bestLength; i2++) {
      // Count consecutive matches starting at this position
      let matchLength = 0;
      let curI1 = i1;
      let curI2 = i2;

      while (
        curI1 < items1.length &&
        curI2 < items2.length &&
        items1[curI1].hash === items2[curI2].hash
      ) {
        matchLength++;
        curI1++;
        curI2++;
      }

      if (matchLength > bestLength) {
        bestLength = matchLength;
        bestI1 = i1;
        bestI2 = i2;
      }
    }
  }

  // Apply minimum length filter
  if (bestLength < opts.minMatchLength) {
    return null;
  }

  // Skip matches that start with items that shouldn't be anchors
  while (bestLength > 0 && opts.shouldSkipAsAnchor(items1[bestI1])) {
    bestI1++;
    bestI2++;
    bestLength--;
  }

  if (bestLength === 0) {
    return null;
  }

  // Apply detail threshold filter
  if (opts.detailThreshold > 0) {
    const maxLen = Math.max(items1.length, items2.length);
    if (bestLength / maxLen < opts.detailThreshold) {
      return null;
    }
  }

  return { i1: bestI1, i2: bestI2, length: bestLength };
}

/**
 * Compute the LCS-based correlation between two arrays.
 *
 * This recursively finds matches and builds a list of correlated sequences
 * showing which parts are equal, deleted, or inserted.
 *
 * @param items1 The original array
 * @param items2 The modified array
 * @param settings Optional settings for match thresholds
 * @returns Array of correlated sequences in order
 */
export function computeCorrelation<T extends Hashable>(
  items1: T[],
  items2: T[],
  settings: LcsSettings = {}
): CorrelatedSequence<T>[] {
  // Handle empty array cases
  if (items1.length === 0 && items2.length === 0) {
    return [];
  }

  if (items1.length === 0) {
    return [
      {
        status: CorrelationStatus.Inserted,
        items1: null,
        items2: items2,
      },
    ];
  }

  if (items2.length === 0) {
    return [
      {
        status: CorrelationStatus.Deleted,
        items1: items1,
        items2: null,
      },
    ];
  }

  // Find longest match
  const match = findLongestMatch(items1, items2, settings);

  // No match found - everything is different
  if (!match) {
    return [
      {
        status: CorrelationStatus.Deleted,
        items1: items1,
        items2: null,
      },
      {
        status: CorrelationStatus.Inserted,
        items1: null,
        items2: items2,
      },
    ];
  }

  const result: CorrelatedSequence<T>[] = [];

  // Process items before the match
  if (match.i1 > 0 || match.i2 > 0) {
    const before = computeCorrelation(
      items1.slice(0, match.i1),
      items2.slice(0, match.i2),
      settings
    );
    result.push(...before);
  }

  // Add the matching portion
  result.push({
    status: CorrelationStatus.Equal,
    items1: items1.slice(match.i1, match.i1 + match.length),
    items2: items2.slice(match.i2, match.i2 + match.length),
  });

  // Process items after the match
  const afterI1 = match.i1 + match.length;
  const afterI2 = match.i2 + match.length;
  if (afterI1 < items1.length || afterI2 < items2.length) {
    const after = computeCorrelation(
      items1.slice(afterI1),
      items2.slice(afterI2),
      settings
    );
    result.push(...after);
  }

  return result;
}

/**
 * Flatten a list of correlated sequences, merging adjacent sequences
 * of the same status.
 */
export function flattenCorrelation<T extends Hashable>(
  sequences: CorrelatedSequence<T>[]
): CorrelatedSequence<T>[] {
  if (sequences.length === 0) return [];

  const result: CorrelatedSequence<T>[] = [];
  let current = { ...sequences[0] };

  for (let i = 1; i < sequences.length; i++) {
    const next = sequences[i];

    if (next.status === current.status) {
      // Merge adjacent sequences of same status
      if (current.items1 && next.items1) {
        current.items1 = [...current.items1, ...next.items1];
      }
      if (current.items2 && next.items2) {
        current.items2 = [...current.items2, ...next.items2];
      }
    } else {
      result.push(current);
      current = { ...next };
    }
  }

  result.push(current);
  return result;
}

/**
 * Simple diff result for text comparison
 */
export interface DiffResult {
  type: 'equal' | 'insert' | 'delete';
  value: string;
}

/**
 * Simple text diff using LCS algorithm.
 * Useful for character-by-character or word-by-word comparison.
 *
 * @param text1 Original text
 * @param text2 Modified text
 * @param splitPattern How to split text into units (default: by character)
 * @returns Array of diff results
 */
export function diffText(
  text1: string,
  text2: string,
  splitPattern: RegExp | string = ''
): DiffResult[] {
  // Split texts into units
  const units1 =
    splitPattern === ''
      ? text1.split('')
      : text1.split(splitPattern).filter((s) => s.length > 0);
  const units2 =
    splitPattern === ''
      ? text2.split('')
      : text2.split(splitPattern).filter((s) => s.length > 0);

  // Create hashable items (for simple strings, the hash is the string itself)
  const items1: Hashable[] = units1.map((s) => ({ hash: s }));
  const items2: Hashable[] = units2.map((s) => ({ hash: s }));

  // Compute correlation
  const correlation = computeCorrelation(items1, items2);

  // Convert to diff results
  const results: DiffResult[] = [];

  for (const seq of correlation) {
    if (seq.status === CorrelationStatus.Equal && seq.items1) {
      results.push({
        type: 'equal',
        value: seq.items1.map((i) => i.hash).join(splitPattern === '' ? '' : ' '),
      });
    } else if (seq.status === CorrelationStatus.Deleted && seq.items1) {
      results.push({
        type: 'delete',
        value: seq.items1.map((i) => i.hash).join(splitPattern === '' ? '' : ' '),
      });
    } else if (seq.status === CorrelationStatus.Inserted && seq.items2) {
      results.push({
        type: 'insert',
        value: seq.items2.map((i) => i.hash).join(splitPattern === '' ? '' : ' '),
      });
    }
  }

  return results;
}
