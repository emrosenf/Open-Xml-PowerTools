import { openPackage, clonePackage, savePackage } from '../core/package';
import {
  PmlChangeType,
  type PmlComparerSettings,
  type PmlComparisonResult,
  type PmlChangeListItem,
  type PmlChangeListOptions,
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
  const items = result.changes.map((change, index) => {
    const anchor = buildAnchor(change.slideIndex, change.shapeId);
    return {
      id: `pml-change-${index + 1}`,
      changeType: change.changeType,
      slideIndex: change.slideIndex,
      shapeName: change.shapeName,
      shapeId: change.shapeId,
      summary: describeChange(change),
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
