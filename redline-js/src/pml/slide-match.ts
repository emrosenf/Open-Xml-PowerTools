import {
  type PresentationSignature,
  type SlideMatch,
  SlideMatchType,
  type SlideSignature,
  type PmlComparerSettings,
} from './types';
import { hashString } from '../core/hash';

export function matchSlides(
  sig1: PresentationSignature,
  sig2: PresentationSignature,
  settings: Required<PmlComparerSettings>
): SlideMatch[] {
  const matches: SlideMatch[] = [];
  const used1 = new Set<number>();
  const used2 = new Set<number>();

  matchByTitleText(sig1, sig2, matches, used1, used2);
  matchByFingerprint(sig1, sig2, matches, used1, used2);

  if (settings.useSlideAlignmentLCS) {
    matchBySimilarity(sig1, sig2, matches, used1, used2, settings);
  } else {
    matchByPosition(sig1, sig2, matches, used1, used2);
  }

  addUnmatched(sig1, sig2, matches, used1, used2);

  return matches
    .sort((a, b) => (a.newIndex ?? Number.MAX_SAFE_INTEGER) - (b.newIndex ?? Number.MAX_SAFE_INTEGER))
    .sort((a, b) => (a.oldIndex ?? Number.MAX_SAFE_INTEGER) - (b.oldIndex ?? Number.MAX_SAFE_INTEGER));
}

function matchByTitleText(
  sig1: PresentationSignature,
  sig2: PresentationSignature,
  matches: SlideMatch[],
  used1: Set<number>,
  used2: Set<number>
): void {
  for (const slide1 of sig1.slides) {
    if (used1.has(slide1.index)) continue;
    if (!slide1.titleText) continue;

    const match = sig2.slides.find(
      (s2) => !used2.has(s2.index) && s2.titleText === slide1.titleText
    );

    if (match) {
      matches.push({
        matchType: SlideMatchType.Matched,
        oldIndex: slide1.index,
        newIndex: match.index,
        oldSlide: slide1,
        newSlide: match,
        similarity: 1.0,
      });
      used1.add(slide1.index);
      used2.add(match.index);
    }
  }
}

function matchByFingerprint(
  sig1: PresentationSignature,
  sig2: PresentationSignature,
  matches: SlideMatch[],
  used1: Set<number>,
  used2: Set<number>
): void {
  const fingerprints1 = new Map<number, string>();
  const fingerprints2 = new Map<number, string>();

  for (const slide of sig1.slides) {
    if (!used1.has(slide.index)) {
      fingerprints1.set(slide.index, computeFingerprint(slide));
    }
  }

  for (const slide of sig2.slides) {
    if (!used2.has(slide.index)) {
      fingerprints2.set(slide.index, computeFingerprint(slide));
    }
  }

  for (const slide1 of sig1.slides) {
    if (used1.has(slide1.index)) continue;
    const fp1 = fingerprints1.get(slide1.index);
    if (!fp1) continue;

    const match = sig2.slides.find((s2) => {
      if (used2.has(s2.index)) return false;
      const fp2 = fingerprints2.get(s2.index);
      return fp2 === fp1;
    });

    if (match) {
      matches.push({
        matchType: SlideMatchType.Matched,
        oldIndex: slide1.index,
        newIndex: match.index,
        oldSlide: slide1,
        newSlide: match,
        similarity: 1.0,
      });
      used1.add(slide1.index);
      used2.add(match.index);
    }
  }
}

function matchBySimilarity(
  sig1: PresentationSignature,
  sig2: PresentationSignature,
  matches: SlideMatch[],
  used1: Set<number>,
  used2: Set<number>,
  settings: Required<PmlComparerSettings>
): void {
  const remaining1 = sig1.slides.filter((s) => !used1.has(s.index));
  const remaining2 = sig2.slides.filter((s) => !used2.has(s.index));
  if (remaining1.length === 0 || remaining2.length === 0) return;

  const similarities: number[][] = Array.from({ length: remaining1.length }, () =>
    Array.from({ length: remaining2.length }, () => 0)
  );

  for (let i = 0; i < remaining1.length; i += 1) {
    for (let j = 0; j < remaining2.length; j += 1) {
      similarities[i][j] = computeSlideSimilarity(remaining1[i], remaining2[j]);
    }
  }

  const matched1 = new Set<number>();
  const matched2 = new Set<number>();

  while (matched1.size < remaining1.length && matched2.size < remaining2.length) {
    let bestSim = 0;
    let bestI = -1;
    let bestJ = -1;

    for (let i = 0; i < remaining1.length; i += 1) {
      if (matched1.has(i)) continue;
      for (let j = 0; j < remaining2.length; j += 1) {
        if (matched2.has(j)) continue;
        if (similarities[i][j] > bestSim) {
          bestSim = similarities[i][j];
          bestI = i;
          bestJ = j;
        }
      }
    }

    if (bestI < 0 || bestSim < settings.slideSimilarityThreshold) {
      break;
    }

    matches.push({
      matchType: SlideMatchType.Matched,
      oldIndex: remaining1[bestI].index,
      newIndex: remaining2[bestJ].index,
      oldSlide: remaining1[bestI],
      newSlide: remaining2[bestJ],
      similarity: bestSim,
    });
    used1.add(remaining1[bestI].index);
    used2.add(remaining2[bestJ].index);
    matched1.add(bestI);
    matched2.add(bestJ);
  }
}

function matchByPosition(
  sig1: PresentationSignature,
  sig2: PresentationSignature,
  matches: SlideMatch[],
  used1: Set<number>,
  used2: Set<number>
): void {
  const remaining1 = sig1.slides.filter((s) => !used1.has(s.index)).sort((a, b) => a.index - b.index);
  const remaining2 = sig2.slides.filter((s) => !used2.has(s.index)).sort((a, b) => a.index - b.index);

  const count = Math.min(remaining1.length, remaining2.length);
  for (let i = 0; i < count; i += 1) {
    matches.push({
      matchType: SlideMatchType.Matched,
      oldIndex: remaining1[i].index,
      newIndex: remaining2[i].index,
      oldSlide: remaining1[i],
      newSlide: remaining2[i],
      similarity: computeSlideSimilarity(remaining1[i], remaining2[i]),
    });
    used1.add(remaining1[i].index);
    used2.add(remaining2[i].index);
  }
}

function addUnmatched(
  sig1: PresentationSignature,
  sig2: PresentationSignature,
  matches: SlideMatch[],
  used1: Set<number>,
  used2: Set<number>
): void {
  for (const slide of sig1.slides) {
    if (used1.has(slide.index)) continue;
    matches.push({
      matchType: SlideMatchType.Deleted,
      oldIndex: slide.index,
      oldSlide: slide,
      similarity: 0,
    });
  }

  for (const slide of sig2.slides) {
    if (used2.has(slide.index)) continue;
    matches.push({
      matchType: SlideMatchType.Inserted,
      newIndex: slide.index,
      newSlide: slide,
      similarity: 0,
    });
  }
}

function computeSlideSimilarity(s1: SlideSignature, s2: SlideSignature): number {
  let score = 0;
  let maxScore = 0;

  if (s1.titleText || s2.titleText) {
    maxScore += 3;
    if (s1.titleText && s1.titleText === s2.titleText) {
      score += 3;
    } else if (s1.titleText && s2.titleText) {
      const similarity = computeTextSimilarity(s1.titleText, s2.titleText);
      score += similarity * 2;
    }
  }

  maxScore += 1;
  if (s1.contentHash && s1.contentHash === s2.contentHash) {
    score += 1;
  }

  maxScore += 1;
  const count1 = s1.shapes.length;
  const count2 = s2.shapes.length;
  if (count1 === count2) {
    score += 1;
  } else if (Math.abs(count1 - count2) <= 2) {
    score += 0.5;
  }

  maxScore += 1;
  const types1 = s1.shapes.map((s) => s.type);
  const types2 = s2.shapes.map((s) => s.type);
  const commonTypes = types1.filter((t) => types2.includes(t)).length;
  const totalTypes = Math.max(types1.length, types2.length);
  if (totalTypes > 0) {
    score += commonTypes / totalTypes;
  }

  maxScore += 2;
  const names1 = new Set(s1.shapes.map((s) => s.name).filter(Boolean) as string[]);
  const names2 = new Set(s2.shapes.map((s) => s.name).filter(Boolean) as string[]);
  const commonNames = [...names1].filter((n) => names2.has(n)).length;
  const totalNames = Math.max(names1.size, names2.size);
  if (totalNames > 0) {
    score += (2 * commonNames) / totalNames;
  }

  return maxScore > 0 ? score / maxScore : 0;
}

function computeTextSimilarity(s1: string, s2: string): number {
  if (!s1 || !s2) return 0;
  if (s1 === s2) return 1;

  const words1 = new Set(s1.toLowerCase().split(' ').filter(Boolean));
  const words2 = new Set(s2.toLowerCase().split(' ').filter(Boolean));
  const intersection = [...words1].filter((w) => words2.has(w)).length;
  const union = new Set([...words1, ...words2]).size;
  return union > 0 ? intersection / union : 0;
}

function computeFingerprint(slide: SlideSignature): string {
  const parts: string[] = [];
  parts.push(slide.titleText ?? '');
  parts.push('|');
  const shapes = [...slide.shapes].sort((a, b) => a.zOrder - b.zOrder);
  for (const shape of shapes) {
    parts.push(shape.name ?? '');
    parts.push(':');
    parts.push(String(shape.type));
    parts.push(':');
    parts.push(shape.textBody?.plainText ?? '');
    parts.push('|');
  }
  return hashString(parts.join(''));
}
