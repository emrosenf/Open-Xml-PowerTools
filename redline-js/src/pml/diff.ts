import {
  type PmlComparerSettings,
  type PresentationSignature,
  type SlideMatch,
  SlideMatchType,
  type SlideSignature,
  type ShapeSignature,
  type ShapeMatch,
  ShapeMatchType,
  PmlChangeType,
  PmlShapeType,
  type TextBodySignature,
  type RunPropertiesSignature,
  PmlComparisonResult,
} from './types';
import { matchShapes } from './shape-match';

export function computeDiff(
  sig1: PresentationSignature,
  sig2: PresentationSignature,
  slideMatches: SlideMatch[],
  settings: Required<PmlComparerSettings>
): PmlComparisonResult {
  const result = new PmlComparisonResult();

  if (sig1.slideCx !== sig2.slideCx || sig1.slideCy !== sig2.slideCy) {
    result.changes.push({
      changeType: PmlChangeType.SlideSizeChanged,
      oldValue: `${sig1.slideCx}x${sig1.slideCy}`,
      newValue: `${sig2.slideCx}x${sig2.slideCy}`,
    });
  }

  for (const slideMatch of slideMatches) {
    if (slideMatch.matchType === SlideMatchType.Inserted) {
      if (settings.compareSlideStructure) {
        result.changes.push({
          changeType: PmlChangeType.SlideInserted,
          slideIndex: slideMatch.newIndex,
        });
      }
      continue;
    }

    if (slideMatch.matchType === SlideMatchType.Deleted) {
      if (settings.compareSlideStructure) {
        result.changes.push({
          changeType: PmlChangeType.SlideDeleted,
          oldSlideIndex: slideMatch.oldIndex,
        });
      }
      continue;
    }

    if (slideMatch.matchType === SlideMatchType.Matched) {
      if (settings.compareSlideStructure && slideMatch.oldIndex !== slideMatch.newIndex) {
        result.changes.push({
          changeType: PmlChangeType.SlideMoved,
          slideIndex: slideMatch.newIndex,
          oldSlideIndex: slideMatch.oldIndex,
        });
      }

      if (slideMatch.oldSlide && slideMatch.newSlide && slideMatch.newIndex) {
        compareSlideContents(
          slideMatch.oldSlide,
          slideMatch.newSlide,
          slideMatch.newIndex,
          settings,
          result
        );
      }
    }
  }

  return result;
}

function compareSlideContents(
  slide1: SlideSignature,
  slide2: SlideSignature,
  slideIndex: number,
  settings: Required<PmlComparerSettings>,
  result: PmlComparisonResult
): void {
  if (slide1.layoutHash !== slide2.layoutHash) {
    result.changes.push({
      changeType: PmlChangeType.SlideLayoutChanged,
      slideIndex,
    });
  }

  if (slide1.backgroundHash !== slide2.backgroundHash) {
    result.changes.push({
      changeType: PmlChangeType.SlideBackgroundChanged,
      slideIndex,
    });
  }

  if (settings.compareNotes && slide1.notesText !== slide2.notesText) {
    result.changes.push({
      changeType: PmlChangeType.SlideNotesChanged,
      slideIndex,
      oldValue: slide1.notesText,
      newValue: slide2.notesText,
    });
  }

  if (settings.compareShapeStructure) {
    const shapeMatches = matchShapes(slide1, slide2, settings);
    for (const shapeMatch of shapeMatches) {
      if (shapeMatch.matchType === ShapeMatchType.Inserted || shapeMatch.matchType === ShapeMatchType.Deleted) {
        const shape = shapeMatch.matchType === ShapeMatchType.Inserted ? shapeMatch.newShape : shapeMatch.oldShape;
        if (!shape) continue;
        result.changes.push({
          changeType:
            shapeMatch.matchType === ShapeMatchType.Inserted
              ? PmlChangeType.ShapeInserted
              : PmlChangeType.ShapeDeleted,
          slideIndex,
          shapeName: shape.name,
          shapeId: shape.id.toString(),
          matchConfidence: shapeMatch.score,
        });
        continue;
      }

      if (shapeMatch.matchType === ShapeMatchType.Matched) {
        if (!shapeMatch.oldShape || !shapeMatch.newShape) continue;
        compareMatchedShapes(
          shapeMatch.oldShape,
          shapeMatch.newShape,
          slideIndex,
          shapeMatch,
          settings,
          result
        );
      }
    }
  }
}

function compareMatchedShapes(
  shape1: ShapeSignature,
  shape2: ShapeSignature,
  slideIndex: number,
  match: ShapeMatch,
  settings: Required<PmlComparerSettings>,
  result: PmlComparisonResult
): void {
  if (settings.compareShapeTransforms && shape1.transform && shape2.transform) {
    const t1 = shape1.transform;
    const t2 = shape2.transform;

    if (!isNear(t1, t2, settings.positionTolerance)) {
      result.changes.push({
        changeType: PmlChangeType.ShapeMoved,
        slideIndex,
        shapeName: shape2.name,
        shapeId: shape2.id.toString(),
        oldX: t1.x,
        oldY: t1.y,
        newX: t2.x,
        newY: t2.y,
        matchConfidence: match.score,
      });
    }

    if (!isSameSize(t1, t2, settings.positionTolerance)) {
      result.changes.push({
        changeType: PmlChangeType.ShapeResized,
        slideIndex,
        shapeName: shape2.name,
        shapeId: shape2.id.toString(),
        oldCx: t1.cx,
        oldCy: t1.cy,
        newCx: t2.cx,
        newCy: t2.cy,
        matchConfidence: match.score,
      });
    }

    if (t1.rotation !== t2.rotation) {
      result.changes.push({
        changeType: PmlChangeType.ShapeRotated,
        slideIndex,
        shapeName: shape2.name,
        shapeId: shape2.id.toString(),
        oldValue: String(t1.rotation),
        newValue: String(t2.rotation),
        matchConfidence: match.score,
      });
    }
  }

  if (shape1.zOrder !== shape2.zOrder) {
    result.changes.push({
      changeType: PmlChangeType.ShapeZOrderChanged,
      slideIndex,
      shapeName: shape2.name,
      shapeId: shape2.id.toString(),
      oldValue: String(shape1.zOrder),
      newValue: String(shape2.zOrder),
    });
  }

  switch (shape1.type) {
    case PmlShapeType.TextBox:
    case PmlShapeType.AutoShape:
      if (settings.compareTextContent) {
        compareTextContent(shape1, shape2, slideIndex, settings, result);
      }
      break;
    case PmlShapeType.Picture:
      if (settings.compareImageContent && shape1.imageHash !== shape2.imageHash) {
        result.changes.push({
          changeType: PmlChangeType.ImageReplaced,
          slideIndex,
          shapeName: shape2.name,
          shapeId: shape2.id.toString(),
        });
      }
      break;
    case PmlShapeType.Table:
      if (settings.compareTables && shape1.tableHash !== shape2.tableHash) {
        result.changes.push({
          changeType: PmlChangeType.TableContentChanged,
          slideIndex,
          shapeName: shape2.name,
          shapeId: shape2.id.toString(),
        });
      }
      break;
    case PmlShapeType.Chart:
      if (settings.compareCharts && shape1.chartHash !== shape2.chartHash) {
        result.changes.push({
          changeType: PmlChangeType.ChartDataChanged,
          slideIndex,
          shapeName: shape2.name,
          shapeId: shape2.id.toString(),
        });
      }
      break;
    default:
      break;
  }
}

function compareTextContent(
  shape1: ShapeSignature,
  shape2: ShapeSignature,
  slideIndex: number,
  settings: Required<PmlComparerSettings>,
  result: PmlComparisonResult
): void {
  const text1 = shape1.textBody;
  const text2 = shape2.textBody;

  if (!text1 && !text2) return;

  if (!text1 || !text2) {
    result.changes.push({
      changeType: PmlChangeType.TextChanged,
      slideIndex,
      shapeName: shape2.name,
      shapeId: shape2.id.toString(),
      oldValue: text1?.plainText ?? '',
      newValue: text2?.plainText ?? '',
    });
    return;
  }

  if (text1.plainText !== text2.plainText) {
    result.changes.push({
      changeType: PmlChangeType.TextChanged,
      slideIndex,
      shapeName: shape2.name,
      shapeId: shape2.id.toString(),
      oldValue: text1.plainText,
      newValue: text2.plainText,
    });
  } else if (settings.compareTextFormatting) {
    if (hasFormattingChanges(text1, text2)) {
      result.changes.push({
        changeType: PmlChangeType.TextFormattingChanged,
        slideIndex,
        shapeName: shape2.name,
        shapeId: shape2.id.toString(),
      });
    }
  }
}

function hasFormattingChanges(text1: TextBodySignature, text2: TextBodySignature): boolean {
  if (text1.paragraphs.length !== text2.paragraphs.length) return true;

  for (let i = 0; i < text1.paragraphs.length; i += 1) {
    const p1 = text1.paragraphs[i];
    const p2 = text2.paragraphs[i];

    if (p1.alignment !== p2.alignment || p1.hasBullet !== p2.hasBullet) {
      return true;
    }

    if (p1.runs.length !== p2.runs.length) return true;

    for (let j = 0; j < p1.runs.length; j += 1) {
      const r1 = p1.runs[j];
      const r2 = p2.runs[j];

      if (r1.properties && r2.properties) {
        if (!runPropertiesEqual(r1.properties, r2.properties)) return true;
      } else if (r1.properties || r2.properties) {
        return true;
      }
    }
  }

  return false;
}

function runPropertiesEqual(a: RunPropertiesSignature, b: RunPropertiesSignature): boolean {
  return (
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.strikethrough === b.strikethrough &&
    a.fontName === b.fontName &&
    a.fontSize === b.fontSize &&
    a.fontColor === b.fontColor
  );
}

function isNear(
  t1: { x: number; y: number },
  t2: { x: number; y: number },
  tolerance: number
): boolean {
  return Math.abs(t1.x - t2.x) <= tolerance && Math.abs(t1.y - t2.y) <= tolerance;
}

function isSameSize(
  t1: { cx: number; cy: number },
  t2: { cx: number; cy: number },
  tolerance: number
): boolean {
  return Math.abs(t1.cx - t2.cx) <= tolerance && Math.abs(t1.cy - t2.cy) <= tolerance;
}

