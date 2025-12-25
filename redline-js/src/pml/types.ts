export interface PmlComparerSettings {
  compareSlideStructure?: boolean;
  compareShapeStructure?: boolean;
  compareTextContent?: boolean;
  compareTextFormatting?: boolean;
  compareShapeTransforms?: boolean;
  compareShapeStyles?: boolean;
  compareImageContent?: boolean;
  compareCharts?: boolean;
  compareTables?: boolean;
  compareNotes?: boolean;
  compareTransitions?: boolean;
  enableFuzzyShapeMatching?: boolean;
  slideSimilarityThreshold?: number;
  shapeSimilarityThreshold?: number;
  positionTolerance?: number;
  useSlideAlignmentLCS?: boolean;
  authorForChanges?: string;
  addSummarySlide?: boolean;
  addNotesAnnotations?: boolean;
  insertedColor?: string;
  deletedColor?: string;
  modifiedColor?: string;
  movedColor?: string;
  formattingColor?: string;
  logCallback?: (message: string) => void;
}

export enum PmlChangeType {
  SlideSizeChanged,
  ThemeChanged,
  SlideInserted,
  SlideDeleted,
  SlideMoved,
  SlideLayoutChanged,
  SlideBackgroundChanged,
  SlideTransitionChanged,
  SlideNotesChanged,
  ShapeInserted,
  ShapeDeleted,
  ShapeMoved,
  ShapeResized,
  ShapeRotated,
  ShapeZOrderChanged,
  ShapeTypeChanged,
  TextChanged,
  TextFormattingChanged,
  ImageReplaced,
  TableContentChanged,
  TableStructureChanged,
  ChartDataChanged,
  ChartFormatChanged,
  ShapeFillChanged,
  ShapeLineChanged,
  ShapeEffectsChanged,
  GroupMembershipChanged,
}

export enum TextChangeType {
  Insert,
  Delete,
  Replace,
  FormatOnly,
}

/**
 * Word count statistics for a PML change.
 */
export interface PmlWordCount {
  /** Number of words deleted */
  deleted: number;
  /** Number of words inserted */
  inserted: number;
}

export interface PmlTextChange {
  type: TextChangeType;
  paragraphIndex: number;
  runIndex: number;
  oldText?: string;
  newText?: string;
}

export interface PmlChange {
  changeType: PmlChangeType;
  slideIndex?: number;
  oldSlideIndex?: number;
  shapeName?: string;
  shapeId?: string;
  oldValue?: string;
  newValue?: string;
  oldX?: number;
  oldY?: number;
  oldCx?: number;
  oldCy?: number;
  newX?: number;
  newY?: number;
  newCx?: number;
  newCy?: number;
  textChanges?: PmlTextChange[];
  matchConfidence?: number;
}

export class PmlComparisonResult {
  public readonly changes: PmlChange[] = [];

  get totalChanges(): number {
    return this.changes.length;
  }

  get slidesInserted(): number {
    return this.count(PmlChangeType.SlideInserted);
  }

  get slidesDeleted(): number {
    return this.count(PmlChangeType.SlideDeleted);
  }

  get slidesMoved(): number {
    return this.count(PmlChangeType.SlideMoved);
  }

  get shapesInserted(): number {
    return this.count(PmlChangeType.ShapeInserted);
  }

  get shapesDeleted(): number {
    return this.count(PmlChangeType.ShapeDeleted);
  }

  get shapesMoved(): number {
    return this.count(PmlChangeType.ShapeMoved);
  }

  get shapesResized(): number {
    return this.count(PmlChangeType.ShapeResized);
  }

  get textChanges(): number {
    return this.count(PmlChangeType.TextChanged);
  }

  get formattingChanges(): number {
    return this.count(PmlChangeType.TextFormattingChanged);
  }

  get imagesReplaced(): number {
    return this.count(PmlChangeType.ImageReplaced);
  }

  getChangesBySlide(slideIndex: number): PmlChange[] {
    return this.changes.filter((c) => c.slideIndex === slideIndex);
  }

  getChangesByType(type: PmlChangeType): PmlChange[] {
    return this.changes.filter((c) => c.changeType === type);
  }

  getChangesByShape(shapeName: string): PmlChange[] {
    return this.changes.filter((c) => c.shapeName === shapeName);
  }

  toJson(): string {
    const payload = {
      Summary: {
        TotalChanges: this.totalChanges,
        SlidesInserted: this.slidesInserted,
        SlidesDeleted: this.slidesDeleted,
        SlidesMoved: this.slidesMoved,
        ShapesInserted: this.shapesInserted,
        ShapesDeleted: this.shapesDeleted,
        ShapesMoved: this.shapesMoved,
        ShapesResized: this.shapesResized,
        TextChanges: this.textChanges,
        FormattingChanges: this.formattingChanges,
        ImagesReplaced: this.imagesReplaced,
      },
      Changes: this.changes.map((change) => ({
        ChangeType: change.changeType,
        SlideIndex: change.slideIndex ?? null,
        OldSlideIndex: change.oldSlideIndex ?? null,
        ShapeName: change.shapeName ?? null,
        OldValue: change.oldValue ?? null,
        NewValue: change.newValue ?? null,
        Description: describeChange(change),
      })),
    };

    return JSON.stringify(payload, null, 2);
  }

  private count(type: PmlChangeType): number {
    return this.changes.filter((c) => c.changeType === type).length;
  }
}

export enum PmlShapeType {
  Unknown,
  TextBox,
  AutoShape,
  Picture,
  Table,
  Chart,
  SmartArt,
  Group,
  Connector,
  OleObject,
  Media,
}

export interface PlaceholderInfo {
  type: string;
  index?: number;
}

export interface TransformSignature {
  x: number;
  y: number;
  cx: number;
  cy: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
}

export interface RunPropertiesSignature {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
}

export interface RunSignature {
  text: string;
  properties?: RunPropertiesSignature;
  contentHash: string;
}

export interface ParagraphSignature {
  runs: RunSignature[];
  plainText: string;
  alignment?: string;
  hasBullet: boolean;
}

export interface TextBodySignature {
  paragraphs: ParagraphSignature[];
  plainText: string;
}

export interface ShapeSignature {
  name?: string;
  id: number;
  type: PmlShapeType;
  placeholder?: PlaceholderInfo;
  transform?: TransformSignature;
  zOrder: number;
  geometryHash?: string;
  textBody?: TextBodySignature;
  imageHash?: string;
  tableHash?: string;
  chartHash?: string;
  children?: ShapeSignature[];
  contentHash: string;
}

export interface SlideSignature {
  index: number;
  relationshipId: string;
  layoutRelationshipId?: string;
  layoutHash?: string;
  shapes: ShapeSignature[];
  notesText?: string;
  titleText?: string;
  contentHash?: string;
  backgroundHash?: string;
}

export interface PresentationSignature {
  slideCx: number;
  slideCy: number;
  slides: SlideSignature[];
  themeHash?: string;
}

export enum SlideMatchType {
  Matched,
  Inserted,
  Deleted,
}

export interface SlideMatch {
  matchType: SlideMatchType;
  oldIndex?: number;
  newIndex?: number;
  oldSlide?: SlideSignature;
  newSlide?: SlideSignature;
  similarity: number;
}

export enum ShapeMatchType {
  Matched,
  Inserted,
  Deleted,
}

export enum ShapeMatchMethod {
  Placeholder,
  NameAndType,
  NameOnly,
  Fuzzy,
}

export interface ShapeMatch {
  matchType: ShapeMatchType;
  oldShape?: ShapeSignature;
  newShape?: ShapeSignature;
  score: number;
  method?: ShapeMatchMethod;
}

export interface PmlChangeListItem {
  id: string;
  changeType: PmlChangeType;
  slideIndex?: number;
  shapeName?: string;
  shapeId?: string;
  summary: string;
  previewText?: string;
  wordCount?: PmlWordCount;
  count?: number;
  details?: {
    oldValue?: string;
    newValue?: string;
    oldSlideIndex?: number;
    textChanges?: PmlTextChange[];
    matchConfidence?: number;
  };
  anchor?: string;
}

export interface PmlChangeListOptions {
  groupBySlide?: boolean;
  maxPreviewLength?: number;
}

export function describeChange(change: PmlChange): string {
  switch (change.changeType) {
    case PmlChangeType.SlideInserted:
      return `Slide ${change.slideIndex} inserted`;
    case PmlChangeType.SlideDeleted:
      return `Slide ${change.oldSlideIndex} deleted`;
    case PmlChangeType.SlideMoved:
      return `Slide moved from position ${change.oldSlideIndex} to ${change.slideIndex}`;
    case PmlChangeType.SlideLayoutChanged:
      return `Slide ${change.slideIndex} layout changed`;
    case PmlChangeType.SlideBackgroundChanged:
      return `Slide ${change.slideIndex} background changed`;
    case PmlChangeType.SlideNotesChanged:
      return `Slide ${change.slideIndex} notes changed`;
    case PmlChangeType.ShapeInserted:
      return `Shape '${change.shapeName}' inserted on slide ${change.slideIndex}`;
    case PmlChangeType.ShapeDeleted:
      return `Shape '${change.shapeName}' deleted from slide ${change.slideIndex}`;
    case PmlChangeType.ShapeMoved:
      return `Shape '${change.shapeName}' moved on slide ${change.slideIndex}`;
    case PmlChangeType.ShapeResized:
      return `Shape '${change.shapeName}' resized on slide ${change.slideIndex}`;
    case PmlChangeType.ShapeRotated:
      return `Shape '${change.shapeName}' rotated on slide ${change.slideIndex}`;
    case PmlChangeType.ShapeZOrderChanged:
      return `Shape '${change.shapeName}' z-order changed on slide ${change.slideIndex}`;
    case PmlChangeType.TextChanged:
      return `Text changed in '${change.shapeName}' on slide ${change.slideIndex}`;
    case PmlChangeType.TextFormattingChanged:
      return `Text formatting changed in '${change.shapeName}' on slide ${change.slideIndex}`;
    case PmlChangeType.ImageReplaced:
      return `Image replaced in '${change.shapeName}' on slide ${change.slideIndex}`;
    case PmlChangeType.TableContentChanged:
      return `Table content changed in '${change.shapeName}' on slide ${change.slideIndex}`;
    case PmlChangeType.ChartDataChanged:
      return `Chart data changed in '${change.shapeName}' on slide ${change.slideIndex}`;
    default:
      return `${PmlChangeType[change.changeType]} on slide ${change.slideIndex}`;
  }
}
