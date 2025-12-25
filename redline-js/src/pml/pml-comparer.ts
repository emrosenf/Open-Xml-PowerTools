import { openPackage, clonePackage, savePackage } from '../core/package';
import {
  PmlChangeType,
  TextChangeType,
  type PmlComparerSettings,
  type PmlComparisonResult,
  type PmlChange,
  type PmlChangeListItem,
  type PmlChangeListOptions,
  type PmlWordCount,
  describeChange,
} from './types';
import { canonicalizePresentation } from './canonicalize';
import { matchSlides } from './slide-match';
import { computeDiff } from './diff';
import { renderMarkedPresentation } from './markup';

export async function comparePresentations(
  older: Buffer | Uint8Array | ArrayBuffer,
  newer: Buffer | Uint8Array | ArrayBuffer,
  settings: PmlComparerSettings = {}
): Promise<PmlComparisonResult> {
  if (!older) {
    throw new Error('Older presentation is required');
  }
  if (!newer) {
    throw new Error('Newer presentation is required');
  }

  const normalized = normalizeSettings(settings);
  log(normalized, 'PmlComparer.Compare: Starting comparison');

  const olderPkg = await openPackage(older);
  const newerPkg = await openPackage(newer);

  const sig1 = await canonicalizePresentation(olderPkg, normalized);
  const sig2 = await canonicalizePresentation(newerPkg, normalized);

  log(normalized, `Canonicalized older: ${sig1.slides.length} slides`);
  log(normalized, `Canonicalized newer: ${sig2.slides.length} slides`);

  const slideMatches = matchSlides(sig1, sig2, normalized);
  const matchedCount = slideMatches.filter((m) => m.matchType === 0).length;
  log(normalized, `Matched ${matchedCount} slides`);

  const result = computeDiff(sig1, sig2, slideMatches, normalized);
  log(normalized, `Found ${result.totalChanges} changes`);

  return result;
}

export async function produceMarkedPresentation(
  older: Buffer | Uint8Array | ArrayBuffer,
  newer: Buffer | Uint8Array | ArrayBuffer,
  settings: PmlComparerSettings = {}
): Promise<Buffer> {
  const normalized = normalizeSettings(settings);
  const result = await comparePresentations(older, newer, normalized);
  const newerPkg = await openPackage(newer);
  const markedPkg = await clonePackage(newerPkg);

  const marked = await renderMarkedPresentation(markedPkg, result, normalized);
  if (marked) {
    return marked;
  }

  return savePackage(markedPkg);
}

export async function canonicalizePresentationDocument(
  doc: Buffer | Uint8Array | ArrayBuffer,
  settings: PmlComparerSettings = {}
): Promise<object> {
  const normalized = normalizeSettings(settings);
  const pkg = await openPackage(doc);
  return canonicalizePresentation(pkg, normalized);
}

export function buildChangeList(
  result: PmlComparisonResult,
  options: PmlChangeListOptions = {}
): PmlChangeListItem[] {
  const maxPreviewLength = options.maxPreviewLength ?? 100;
  
  const items = result.changes.map((change, index) => {
    const anchor = buildAnchor(change.slideIndex, change.shapeId);
    const wordCount = computeWordCount(change);
    const previewText = buildPreviewText(change, maxPreviewLength);
    
    return {
      id: `pml-change-${index + 1}`,
      changeType: change.changeType,
      slideIndex: change.slideIndex,
      shapeName: change.shapeName,
      shapeId: change.shapeId,
      summary: describeChange(change),
      previewText,
      wordCount,
      details: {
        oldValue: change.oldValue,
        newValue: change.newValue,
        oldSlideIndex: change.oldSlideIndex,
        textChanges: change.textChanges,
        matchConfidence: change.matchConfidence,
      },
      anchor,
    };
  });

  if (options.groupBySlide) {
    return groupBySlide(items);
  }

  return items;
}

function groupBySlide(items: PmlChangeListItem[]): PmlChangeListItem[] {
  const grouped: PmlChangeListItem[] = [];
  const bySlide = new Map<number, PmlChangeListItem[]>();

  for (const item of items) {
    if (item.slideIndex === undefined) {
      grouped.push(item);
      continue;
    }
    if (!bySlide.has(item.slideIndex)) {
      bySlide.set(item.slideIndex, []);
    }
    bySlide.get(item.slideIndex)!.push(item);
  }

  for (const [slideIndex, slideItems] of bySlide) {
    grouped.push({
      id: `pml-slide-${slideIndex}`,
      changeType: PmlChangeType.SlideMoved,
      slideIndex,
      summary: `Slide ${slideIndex} changes (${slideItems.length})`,
      anchor: buildAnchor(slideIndex),
    });
    grouped.push(...slideItems);
  }

  return grouped;
}

function buildAnchor(slideIndex?: number, shapeId?: string): string | undefined {
  if (!slideIndex) return undefined;
  if (!shapeId) return `slide-${slideIndex}`;
  return `slide-${slideIndex}#shape-${shapeId}`;
}

function normalizeSettings(settings: PmlComparerSettings): Required<PmlComparerSettings> {
  return {
    compareSlideStructure: settings.compareSlideStructure ?? true,
    compareShapeStructure: settings.compareShapeStructure ?? true,
    compareTextContent: settings.compareTextContent ?? true,
    compareTextFormatting: settings.compareTextFormatting ?? true,
    compareShapeTransforms: settings.compareShapeTransforms ?? true,
    compareShapeStyles: settings.compareShapeStyles ?? false,
    compareImageContent: settings.compareImageContent ?? true,
    compareCharts: settings.compareCharts ?? true,
    compareTables: settings.compareTables ?? true,
    compareNotes: settings.compareNotes ?? false,
    compareTransitions: settings.compareTransitions ?? false,
    enableFuzzyShapeMatching: settings.enableFuzzyShapeMatching ?? true,
    slideSimilarityThreshold: settings.slideSimilarityThreshold ?? 0.4,
    shapeSimilarityThreshold: settings.shapeSimilarityThreshold ?? 0.7,
    positionTolerance: settings.positionTolerance ?? 91440,
    useSlideAlignmentLCS: settings.useSlideAlignmentLCS ?? true,
    authorForChanges: settings.authorForChanges ?? 'Open-Xml-PowerTools',
    addSummarySlide: settings.addSummarySlide ?? true,
    addNotesAnnotations: settings.addNotesAnnotations ?? true,
    insertedColor: settings.insertedColor ?? '00AA00',
    deletedColor: settings.deletedColor ?? 'FF0000',
    modifiedColor: settings.modifiedColor ?? 'FFA500',
    movedColor: settings.movedColor ?? '0000FF',
    formattingColor: settings.formattingColor ?? '9932CC',
    logCallback: settings.logCallback,
  };
}

function log(settings: Required<PmlComparerSettings>, message: string): void {
  if (settings.logCallback) {
    settings.logCallback(message);
  }
}

function countWords(text: string | undefined): number {
  if (!text) return 0;
  return text.trim().split(/\s+/).filter(Boolean).length;
}

function computeWordCount(change: PmlChange): PmlWordCount | undefined {
  if (change.textChanges && change.textChanges.length > 0) {
    let deleted = 0;
    let inserted = 0;
    for (const tc of change.textChanges) {
      if (tc.type === TextChangeType.Delete || tc.type === TextChangeType.Replace) {
        deleted += countWords(tc.oldText);
      }
      if (tc.type === TextChangeType.Insert || tc.type === TextChangeType.Replace) {
        inserted += countWords(tc.newText);
      }
    }
    if (deleted > 0 || inserted > 0) {
      return { deleted, inserted };
    }
  }

  if (change.oldValue || change.newValue) {
    const deleted = countWords(change.oldValue);
    const inserted = countWords(change.newValue);
    if (deleted > 0 || inserted > 0) {
      return { deleted, inserted };
    }
  }

  return undefined;
}

function buildPreviewText(change: PmlChange, maxLength: number): string | undefined {
  if (change.textChanges && change.textChanges.length > 0) {
    const parts: string[] = [];
    for (const tc of change.textChanges) {
      if (tc.type === TextChangeType.Delete && tc.oldText) {
        parts.push(`-"${tc.oldText}"`);
      } else if (tc.type === TextChangeType.Insert && tc.newText) {
        parts.push(`+"${tc.newText}"`);
      } else if (tc.type === TextChangeType.Replace) {
        if (tc.oldText && tc.newText) {
          parts.push(`"${tc.oldText}" → "${tc.newText}"`);
        }
      }
    }
    const preview = parts.join(' ');
    return truncate(preview, maxLength);
  }

  if (change.oldValue && change.newValue) {
    return truncate(`"${change.oldValue}" → "${change.newValue}"`, maxLength);
  }
  if (change.oldValue) {
    return truncate(`-"${change.oldValue}"`, maxLength);
  }
  if (change.newValue) {
    return truncate(`+"${change.newValue}"`, maxLength);
  }

  return undefined;
}

function truncate(text: string, maxLength: number): string {
  if (text.length <= maxLength) return text;
  return text.slice(0, maxLength - 3) + '...';
}
