import {
  type SlideSignature,
  type ShapeMatch,
  ShapeMatchType,
  ShapeMatchMethod,
  type ShapeSignature,
  type PmlComparerSettings,
  PmlShapeType,
} from './types';

export function matchShapes(
  slide1: SlideSignature,
  slide2: SlideSignature,
  settings: Required<PmlComparerSettings>
): ShapeMatch[] {
  const matches: ShapeMatch[] = [];
  const used1 = new Set<string>();
  const used2 = new Set<string>();

  matchByPlaceholder(slide1, slide2, matches, used1, used2);
  matchByNameAndType(slide1, slide2, matches, used1, used2);
  matchByNameOnly(slide1, slide2, matches, used1, used2);

  if (settings.enableFuzzyShapeMatching) {
    fuzzyMatch(slide1, slide2, matches, used1, used2, settings);
  }

  addUnmatched(slide1, slide2, matches, used1, used2);

  return matches;
}

function getShapeKey(shape: ShapeSignature): string {
  return `${shape.id}:${shape.name ?? ''}`;
}

function matchByPlaceholder(
  slide1: SlideSignature,
  slide2: SlideSignature,
  matches: ShapeMatch[],
  used1: Set<string>,
  used2: Set<string>
): void {
  const placeholders1 = slide1.shapes.filter((s) => s.placeholder);
  for (const shape1 of placeholders1) {
    const key1 = getShapeKey(shape1);
    if (used1.has(key1)) continue;

    const match = slide2.shapes.find(
      (s2) =>
        s2.placeholder &&
        !used2.has(getShapeKey(s2)) &&
        s2.placeholder?.type === shape1.placeholder?.type &&
        s2.placeholder?.index === shape1.placeholder?.index
    );

    if (match) {
      matches.push({
        matchType: ShapeMatchType.Matched,
        oldShape: shape1,
        newShape: match,
        score: 1.0,
        method: ShapeMatchMethod.Placeholder,
      });
      used1.add(key1);
      used2.add(getShapeKey(match));
    }
  }
}

function matchByNameAndType(
  slide1: SlideSignature,
  slide2: SlideSignature,
  matches: ShapeMatch[],
  used1: Set<string>,
  used2: Set<string>
): void {
  for (const shape1 of slide1.shapes) {
    const key1 = getShapeKey(shape1);
    if (used1.has(key1)) continue;
    if (!shape1.name) continue;

    const match = slide2.shapes.find(
      (s2) =>
        !used2.has(getShapeKey(s2)) &&
        s2.name === shape1.name &&
        s2.type === shape1.type
    );

    if (match) {
      matches.push({
        matchType: ShapeMatchType.Matched,
        oldShape: shape1,
        newShape: match,
        score: 0.95,
        method: ShapeMatchMethod.NameAndType,
      });
      used1.add(key1);
      used2.add(getShapeKey(match));
    }
  }
}

function matchByNameOnly(
  slide1: SlideSignature,
  slide2: SlideSignature,
  matches: ShapeMatch[],
  used1: Set<string>,
  used2: Set<string>
): void {
  for (const shape1 of slide1.shapes) {
    const key1 = getShapeKey(shape1);
    if (used1.has(key1)) continue;
    if (!shape1.name) continue;

    const match = slide2.shapes.find(
      (s2) => !used2.has(getShapeKey(s2)) && s2.name === shape1.name
    );

    if (match) {
      matches.push({
        matchType: ShapeMatchType.Matched,
        oldShape: shape1,
        newShape: match,
        score: 0.8,
        method: ShapeMatchMethod.NameOnly,
      });
      used1.add(key1);
      used2.add(getShapeKey(match));
    }
  }
}

function fuzzyMatch(
  slide1: SlideSignature,
  slide2: SlideSignature,
  matches: ShapeMatch[],
  used1: Set<string>,
  used2: Set<string>,
  settings: Required<PmlComparerSettings>
): void {
  const remaining1 = slide1.shapes.filter((s) => !used1.has(getShapeKey(s)));
  const remaining2 = slide2.shapes.filter((s) => !used2.has(getShapeKey(s)));

  for (const shape1 of remaining1) {
    const key1 = getShapeKey(shape1);
    if (used1.has(key1)) continue;

    let bestScore = 0;
    let bestMatch: ShapeSignature | undefined;

    for (const shape2 of remaining2) {
      const key2 = getShapeKey(shape2);
      if (used2.has(key2)) continue;

      const score = computeShapeMatchScore(shape1, shape2, settings);
      if (score > bestScore && score >= settings.shapeSimilarityThreshold) {
        bestScore = score;
        bestMatch = shape2;
      }
    }

    if (bestMatch) {
      matches.push({
        matchType: ShapeMatchType.Matched,
        oldShape: shape1,
        newShape: bestMatch,
        score: bestScore,
        method: ShapeMatchMethod.Fuzzy,
      });
      used1.add(key1);
      used2.add(getShapeKey(bestMatch));
    }
  }
}

function addUnmatched(
  slide1: SlideSignature,
  slide2: SlideSignature,
  matches: ShapeMatch[],
  used1: Set<string>,
  used2: Set<string>
): void {
  for (const shape of slide1.shapes) {
    if (used1.has(getShapeKey(shape))) continue;
    matches.push({
      matchType: ShapeMatchType.Deleted,
      oldShape: shape,
      score: 0,
    });
  }

  for (const shape of slide2.shapes) {
    if (used2.has(getShapeKey(shape))) continue;
    matches.push({
      matchType: ShapeMatchType.Inserted,
      newShape: shape,
      score: 0,
    });
  }
}

function computeShapeMatchScore(
  s1: ShapeSignature,
  s2: ShapeSignature,
  settings: Required<PmlComparerSettings>
): number {
  if (s1.type !== s2.type) return 0;

  let score = 0.2;

  if (s1.transform && s2.transform) {
    if (isNear(s1.transform, s2.transform, settings.positionTolerance)) {
      score += 0.3;
    } else {
      const dx = s1.transform.x - s2.transform.x;
      const dy = s1.transform.y - s2.transform.y;
      const distance = Math.sqrt(dx * dx + dy * dy);
      if (distance < settings.positionTolerance * 5) {
        score += 0.1;
      }
    }
  }

  if (s1.type === PmlShapeType.Picture) {
    if (s1.imageHash && s1.imageHash === s2.imageHash) {
      score += 0.5;
    }
  } else if (s1.textBody && s2.textBody) {
    if (s1.textBody.plainText === s2.textBody.plainText) {
      score += 0.5;
    } else {
      const textSim = computeTextSimilarity(s1.textBody.plainText, s2.textBody.plainText);
      score += textSim * 0.5;
    }
  } else if (s1.contentHash === s2.contentHash) {
    score += 0.5;
  }

  return score;
}

function computeTextSimilarity(s1: string, s2: string): number {
  if (!s1 && !s2) return 1;
  if (!s1 || !s2) return 0;
  if (s1 === s2) return 1;

  const maxLen = Math.max(s1.length, s2.length);
  if (maxLen === 0) return 1;

  const distance = levenshteinDistance(s1, s2);
  return 1.0 - distance / maxLen;
}

function levenshteinDistance(s1: string, s2: string): number {
  const m = s1.length;
  const n = s2.length;
  const d: number[][] = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));

  for (let i = 0; i <= m; i += 1) d[i][0] = i;
  for (let j = 0; j <= n; j += 1) d[0][j] = j;

  for (let j = 1; j <= n; j += 1) {
    for (let i = 1; i <= m; i += 1) {
      const cost = s1[i - 1] === s2[j - 1] ? 0 : 1;
      d[i][j] = Math.min(
        d[i - 1][j] + 1,
        d[i][j - 1] + 1,
        d[i - 1][j - 1] + cost
      );
    }
  }

  return d[m][n];
}

function isNear(
  t1: { x: number; y: number },
  t2: { x: number; y: number },
  tolerance: number
): boolean {
  return Math.abs(t1.x - t2.x) <= tolerance && Math.abs(t1.y - t2.y) <= tolerance;
}
