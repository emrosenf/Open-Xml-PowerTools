import {
  getPartAsXml,
  getPartAsString,
  getRelationships,
  type OoxmlPackage,
} from '../core/package';
import {
  getAttribute,
  getChildren,
  getTagName,
  getTextContent,
  buildXml,
  type XmlNode,
} from '../core/xml';
import { hashBytes, hashString } from '../core/hash';
import {
  type PmlComparerSettings,
  type PresentationSignature,
  type SlideSignature,
  type ShapeSignature,
  type TextBodySignature,
  type ParagraphSignature,
  type RunSignature,
  type RunPropertiesSignature,
  type TransformSignature,
  type PlaceholderInfo,
  PmlShapeType,
} from './types';

export async function canonicalizePresentation(
  pkg: OoxmlPackage,
  settings: Required<PmlComparerSettings>
): Promise<PresentationSignature> {
  const presentationXml = await getPartAsXml(pkg, 'ppt/presentation.xml');
  const signature: PresentationSignature = {
    slideCx: 0,
    slideCy: 0,
    slides: [],
  };

  if (!presentationXml) {
    return signature;
  }

  const presentationRoot = findFirstElement(presentationXml, 'p:presentation');
  if (!presentationRoot) {
    return signature;
  }

  const slideSize = findChild(presentationRoot, 'p:sldSz');
  if (slideSize) {
    signature.slideCx = parseLong(getAttribute(slideSize, 'cx'));
    signature.slideCy = parseLong(getAttribute(slideSize, 'cy'));
  }

  const slideList = findChild(presentationRoot, 'p:sldIdLst');
  if (!slideList) {
    return signature;
  }

  const rels = await getRelationships(pkg, 'ppt/presentation.xml');
  const slideIds = getChildren(slideList).filter((node) => getTagName(node) === 'p:sldId');

  let slideIndex = 1;
  for (const slideId of slideIds) {
    const relId = getAttribute(slideId, 'r:id');
    if (!relId) {
      slideIndex += 1;
      continue;
    }

    const rel = rels.find((r) => r.id === relId);
    if (!rel) {
      slideIndex += 1;
      continue;
    }

    const slidePath = resolveRelationshipTarget('ppt/presentation.xml', rel.target);
    if (!slidePath) {
      slideIndex += 1;
      continue;
    }

    const slideSignature = await canonicalizeSlide(pkg, slidePath, slideIndex, relId, settings);
    signature.slides.push(slideSignature);
    slideIndex += 1;
  }

  return signature;
}

async function canonicalizeSlide(
  pkg: OoxmlPackage,
  slidePath: string,
  index: number,
  relId: string,
  settings: Required<PmlComparerSettings>
): Promise<SlideSignature> {
  const slideXml = await getPartAsXml(pkg, slidePath);
  const signature: SlideSignature = {
    index,
    relationshipId: relId,
    shapes: [],
  };

  if (!slideXml) {
    return signature;
  }

  const slideRoot = findFirstElement(slideXml, 'p:sld');
  if (!slideRoot) {
    return signature;
  }

  const slideRels = await getRelationships(pkg, slidePath);
  const layoutRel = slideRels.find((rel) => rel.type.includes('slideLayout'));
  if (layoutRel) {
    signature.layoutRelationshipId = layoutRel.id;
    const layoutPath = resolveRelationshipTarget(slidePath, layoutRel.target);
    if (layoutPath) {
      const layoutXml = await getPartAsXml(pkg, layoutPath);
      const layoutRoot = layoutXml ? findFirstElement(layoutXml, 'p:sldLayout') : null;
      const layoutType = layoutRoot ? getAttribute(layoutRoot, 'type') || 'custom' : 'custom';
      signature.layoutHash = hashString(layoutType);
    }
  }

  const commonSlide = findChild(slideRoot, 'p:cSld');
  if (!commonSlide) {
    return signature;
  }

  const background = findChild(commonSlide, 'p:bg');
  if (background) {
    signature.backgroundHash = hashString(buildXml(background));
  }

  const shapeTree = findChild(commonSlide, 'p:spTree');
  if (!shapeTree) {
    return signature;
  }

  let zOrder = 0;
  for (const element of getChildren(shapeTree)) {
    const tag = getTagName(element);
    if (!tag) continue;
    if (!isShapeElement(tag)) continue;

    const shapeSignature = await canonicalizeShape(pkg, element, slidePath, zOrder, settings);
    if (shapeSignature) {
      signature.shapes.push(shapeSignature);
      if (shapeSignature.placeholder?.type === 'title' || shapeSignature.placeholder?.type === 'ctrTitle') {
        signature.titleText = shapeSignature.textBody?.plainText;
      }
      zOrder += 1;
    }
  }

  if (settings.compareNotes) {
    signature.notesText = await extractNotesText(pkg, slidePath);
  }

  const contentBuilder: string[] = [];
  contentBuilder.push(signature.titleText ?? '');
  for (const shape of signature.shapes) {
    contentBuilder.push('|');
    contentBuilder.push(shape.name ?? '');
    contentBuilder.push(':');
    contentBuilder.push(String(shape.type));
    contentBuilder.push(':');
    contentBuilder.push(shape.textBody?.plainText ?? '');
  }
  signature.contentHash = hashString(contentBuilder.join(''));

  return signature;
}

async function canonicalizeShape(
  pkg: OoxmlPackage,
  element: XmlNode,
  slidePath: string,
  zOrder: number,
  settings: Required<PmlComparerSettings>
): Promise<ShapeSignature | null> {
  const tag = getTagName(element);
  if (!tag) return null;

  const signature: ShapeSignature = {
    id: 0,
    type: PmlShapeType.Unknown,
    zOrder,
    contentHash: '',
  };

  if (tag === 'p:sp') {
    signature.type = PmlShapeType.AutoShape;
  } else if (tag === 'p:pic') {
    signature.type = PmlShapeType.Picture;
  } else if (tag === 'p:graphicFrame') {
    const graphic = findChild(element, 'a:graphic');
    const graphicData = graphic ? findChild(graphic, 'a:graphicData') : null;
    const uri = graphicData ? getAttribute(graphicData, 'uri') : undefined;
    if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
      signature.type = PmlShapeType.Table;
    } else if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/chart') {
      signature.type = PmlShapeType.Chart;
    } else if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/diagram') {
      signature.type = PmlShapeType.SmartArt;
    } else {
      signature.type = PmlShapeType.OleObject;
    }
  } else if (tag === 'p:grpSp') {
    signature.type = PmlShapeType.Group;
  } else if (tag === 'p:cxnSp') {
    signature.type = PmlShapeType.Connector;
  }

  const nvProps =
    findChild(element, 'p:nvSpPr') ||
    findChild(element, 'p:nvPicPr') ||
    findChild(element, 'p:nvGraphicFramePr') ||
    findChild(element, 'p:nvGrpSpPr') ||
    findChild(element, 'p:nvCxnSpPr');

  if (nvProps) {
    const cNvPr = findChild(nvProps, 'p:cNvPr');
    if (cNvPr) {
      signature.name = getAttribute(cNvPr, 'name') ?? '';
      signature.id = parseInt(getAttribute(cNvPr, 'id') ?? '0', 10);
    }

    const nvPr = findChild(nvProps, 'p:nvPr');
    const ph = nvPr ? findChild(nvPr, 'p:ph') : null;
    if (ph) {
      signature.placeholder = {
        type: getAttribute(ph, 'type') ?? 'body',
        index: parseInt(getAttribute(ph, 'idx') ?? '', 10) || undefined,
      } as PlaceholderInfo;
    }
  }

  const shapeProps = findChild(element, 'p:spPr') || findChild(element, 'p:grpSpPr');
  if (shapeProps) {
    const xfrm = findChild(shapeProps, 'a:xfrm');
    if (xfrm) {
      signature.transform = extractTransform(xfrm);
    }

    const prstGeom = findChild(shapeProps, 'a:prstGeom');
    const custGeom = findChild(shapeProps, 'a:custGeom');
    if (prstGeom) {
      signature.geometryHash = getAttribute(prstGeom, 'prst') ?? undefined;
    } else if (custGeom) {
      signature.geometryHash = hashString(buildXml(custGeom));
    }
  }

  if (tag === 'p:grpSp') {
    const grpProps = findChild(element, 'p:grpSpPr');
    const grpXfrm = grpProps ? findChild(grpProps, 'a:xfrm') : null;
    if (grpXfrm && !signature.transform) {
      signature.transform = extractTransform(grpXfrm);
    }
  }

  const textBodyNode = findChild(element, 'p:txBody');
  if (textBodyNode) {
    signature.textBody = extractTextBody(textBodyNode);
    if (signature.type === PmlShapeType.AutoShape && signature.textBody.plainText) {
      signature.type = PmlShapeType.TextBox;
    }
  }

  if (signature.type === PmlShapeType.Picture && settings.compareImageContent) {
    signature.imageHash = await extractImageHash(pkg, element, slidePath);
  }

  if (signature.type === PmlShapeType.Table && settings.compareTables) {
    signature.tableHash = extractTableHash(element);
  }

  if (signature.type === PmlShapeType.Chart && settings.compareCharts) {
    signature.chartHash = await extractChartHash(pkg, element, slidePath);
  }

  if (signature.type === PmlShapeType.Group) {
    const children: ShapeSignature[] = [];
    let childOrder = 0;
    for (const child of getChildren(element)) {
      const childTag = getTagName(child);
      if (!childTag || !isShapeElement(childTag)) continue;
      const childSig = await canonicalizeShape(pkg, child, slidePath, childOrder, settings);
      if (childSig) {
        children.push(childSig);
        childOrder += 1;
      }
    }
    signature.children = children;
  }

  const contentParts: string[] = [];
  contentParts.push(String(signature.type));
  contentParts.push('|');
  contentParts.push(signature.textBody?.plainText ?? '');
  contentParts.push('|');
  contentParts.push(signature.imageHash ?? '');
  contentParts.push('|');
  contentParts.push(signature.tableHash ?? '');
  contentParts.push('|');
  contentParts.push(signature.chartHash ?? '');
  signature.contentHash = hashString(contentParts.join(''));

  return signature;
}

function extractTransform(node: XmlNode): TransformSignature {
  const off = findChild(node, 'a:off');
  const ext = findChild(node, 'a:ext');

  return {
    x: parseLong(off ? getAttribute(off, 'x') : undefined),
    y: parseLong(off ? getAttribute(off, 'y') : undefined),
    cx: parseLong(ext ? getAttribute(ext, 'cx') : undefined),
    cy: parseLong(ext ? getAttribute(ext, 'cy') : undefined),
    rotation: parseInt(getAttribute(node, 'rot') ?? '0', 10) || 0,
    flipH: parseBoolean(getAttribute(node, 'flipH')),
    flipV: parseBoolean(getAttribute(node, 'flipV')),
  };
}

function extractTextBody(node: XmlNode): TextBodySignature {
  const paragraphs: ParagraphSignature[] = [];
  const paragraphText: string[] = [];

  for (const para of getChildren(node)) {
    if (getTagName(para) !== 'a:p') continue;
    const paraSig: ParagraphSignature = {
      runs: [],
      plainText: '',
      hasBullet: false,
    };

    const pPr = findChild(para, 'a:pPr');
    if (pPr) {
      paraSig.alignment = getAttribute(pPr, 'algn') ?? undefined;
      paraSig.hasBullet = Boolean(findChild(pPr, 'a:buChar') || findChild(pPr, 'a:buAutoNum'));
    }

    const runText: string[] = [];
    for (const run of getChildren(para)) {
      const runTag = getTagName(run);
      if (runTag === 'a:r') {
        const t = findChild(run, 'a:t');
        const text = t ? getTextContent(t) : '';
        runText.push(text);

        const rPr = findChild(run, 'a:rPr');
        const runSig: RunSignature = {
          text,
          properties: rPr ? extractRunProperties(rPr) : undefined,
          contentHash: hashString(text),
        };
        paraSig.runs.push(runSig);
      } else if (runTag === 'a:fld') {
        const t = findChild(run, 'a:t');
        const text = t ? getTextContent(t) : '';
        runText.push(text);
        paraSig.runs.push({ text, contentHash: hashString(text) });
      }
    }

    paraSig.plainText = runText.join('');
    paragraphText.push(paraSig.plainText);
    paragraphs.push(paraSig);
  }

  return {
    paragraphs,
    plainText: paragraphText.join('\n'),
  };
}

function extractRunProperties(node: XmlNode): RunPropertiesSignature {
  const bold = parseBoolean(getAttribute(node, 'b'));
  const italic = parseBoolean(getAttribute(node, 'i'));
  const underline = (getAttribute(node, 'u') ?? '') !== 'none' && Boolean(getAttribute(node, 'u'));
  const strikethrough = (getAttribute(node, 'strike') ?? '') !== 'noStrike' && Boolean(getAttribute(node, 'strike'));
  const fontSize = parseInt(getAttribute(node, 'sz') ?? '', 10) || undefined;

  const latin = findChild(node, 'a:latin');
  const fontName = latin ? getAttribute(latin, 'typeface') ?? undefined : undefined;

  let fontColor: string | undefined;
  const solidFill = findChild(node, 'a:solidFill');
  if (solidFill) {
    const srgb = findChild(solidFill, 'a:srgbClr');
    fontColor = srgb ? getAttribute(srgb, 'val') ?? undefined : undefined;
  }

  return {
    bold,
    italic,
    underline,
    strikethrough,
    fontName,
    fontSize,
    fontColor,
  };
}

async function extractImageHash(pkg: OoxmlPackage, element: XmlNode, slidePath: string): Promise<string | undefined> {
  const blipFill = findChild(element, 'p:blipFill');
  const blip = blipFill ? findChild(blipFill, 'a:blip') : null;
  const embed = blip ? getAttribute(blip, 'r:embed') : undefined;
  if (!embed) return undefined;

  const rels = await getRelationships(pkg, slidePath);
  const rel = rels.find((r) => r.id === embed);
  if (!rel) return undefined;

  const imagePath = resolveRelationshipTarget(slidePath, rel.target);
  if (!imagePath) return undefined;

  const file = pkg.zip.file(imagePath);
  if (!file) return undefined;

  const bytes = await file.async('uint8array');
  return hashBytes(bytes);
}

function extractTableHash(element: XmlNode): string | undefined {
  const graphic = findChild(element, 'a:graphic');
  const graphicData = graphic ? findChild(graphic, 'a:graphicData') : null;
  const table = graphicData ? findChild(graphicData, 'a:tbl') : null;
  if (!table) return undefined;

  const content: string[] = [];
  for (const row of getChildren(table)) {
    if (getTagName(row) !== 'a:tr') continue;
    for (const cell of getChildren(row)) {
      if (getTagName(cell) !== 'a:tc') continue;
      const txBody = findChild(cell, 'a:txBody');
      if (txBody) {
        const text = extractTextBody(txBody);
        content.push(text.plainText);
        content.push('|');
      }
    }
    content.push('||');
  }

  return hashString(content.join(''));
}

async function extractChartHash(pkg: OoxmlPackage, element: XmlNode, slidePath: string): Promise<string | undefined> {
  const graphic = findChild(element, 'a:graphic');
  const graphicData = graphic ? findChild(graphic, 'a:graphicData') : null;
  const chartRef = graphicData ? findChild(graphicData, 'c:chart') : null;
  const relId = chartRef ? getAttribute(chartRef, 'r:id') : undefined;
  if (!relId) return undefined;

  const rels = await getRelationships(pkg, slidePath);
  const rel = rels.find((r) => r.id === relId);
  if (!rel) return undefined;

  const chartPath = resolveRelationshipTarget(slidePath, rel.target);
  if (!chartPath) return undefined;

  const chartXml = await getPartAsString(pkg, chartPath);
  if (!chartXml) return undefined;

  return hashString(chartXml);
}

async function extractNotesText(pkg: OoxmlPackage, slidePath: string): Promise<string | undefined> {
  const rels = await getRelationships(pkg, slidePath);
  const notesRel = rels.find((rel) => rel.type.includes('notesSlide'));
  if (!notesRel) return undefined;

  const notesPath = resolveRelationshipTarget(slidePath, notesRel.target);
  if (!notesPath) return undefined;

  const notesXml = await getPartAsXml(pkg, notesPath);
  if (!notesXml) return undefined;

  const notesRoot = findFirstElement(notesXml, 'p:notes');
  const spTree = notesRoot ? findChild(findChild(notesRoot, 'p:cSld') ?? notesRoot, 'p:spTree') : null;
  if (!spTree) return undefined;

  const textParts: string[] = [];
  for (const sp of getChildren(spTree)) {
    if (getTagName(sp) !== 'p:sp') continue;
    const txBody = findChild(sp, 'p:txBody');
    if (!txBody) continue;
    const text = extractTextBody(txBody);
    if (text.plainText) {
      textParts.push(text.plainText);
    }
  }

  return textParts.join('\n');
}

function isShapeElement(tag: string): boolean {
  return tag === 'p:sp' || tag === 'p:pic' || tag === 'p:graphicFrame' || tag === 'p:grpSp' || tag === 'p:cxnSp';
}

function findChild(node: XmlNode, tagName: string): XmlNode | null {
  for (const child of getChildren(node)) {
    if (getTagName(child) === tagName) {
      return child;
    }
  }
  return null;
}

function findFirstElement(nodes: XmlNode[], tagName: string): XmlNode | null {
  for (const node of nodes) {
    if (getTagName(node) === tagName) {
      return node;
    }
    const found = findInChildren(node, tagName);
    if (found) return found;
  }
  return null;
}

function findInChildren(node: XmlNode, tagName: string): XmlNode | null {
  for (const child of getChildren(node)) {
    if (getTagName(child) === tagName) {
      return child;
    }
    const found = findInChildren(child, tagName);
    if (found) return found;
  }
  return null;
}

function resolveRelationshipTarget(basePath: string, target: string): string | null {
  if (!target) return null;
  if (target.startsWith('/')) return target.slice(1);

  const baseParts = basePath.split('/');
  baseParts.pop();

  const targetParts = target.split('/');
  for (const part of targetParts) {
    if (part === '..') {
      baseParts.pop();
    } else if (part !== '.') {
      baseParts.push(part);
    }
  }

  return baseParts.join('/');
}

function parseLong(value?: string): number {
  if (!value) return 0;
  const parsed = parseInt(value, 10);
  return Number.isNaN(parsed) ? 0 : parsed;
}

function parseBoolean(value?: string): boolean {
  if (!value) return false;
  return value === '1' || value === 'true';
}
