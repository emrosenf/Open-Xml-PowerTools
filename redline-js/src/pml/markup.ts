import {
  type OoxmlPackage,
  getPartAsXml,
  setPartFromXml,
  savePackage,
  listParts,
  getRelationships,
} from '../core/package';
import {
  type XmlNode,
  getTagName,
  getChildren,
  getAttribute,
  findByTagName,
  createNode,
} from '../core/xml';
import {
  PmlChangeType,
  type PmlComparerSettings,
  type PmlComparisonResult,
  type PmlChange,
  describeChange,
} from './types';

const PRESENTATION_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const NOTES_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide';
const NOTES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const RELS_XMLNS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const PML_XMLNS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A_XMLNS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_XMLNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export async function renderMarkedPresentation(
  pkg: OoxmlPackage,
  result: PmlComparisonResult,
  settings: Required<PmlComparerSettings>
): Promise<Buffer | null> {
  if (result.totalChanges === 0) {
    return null;
  }

  const presentationXml = await getPartAsXml(pkg, 'ppt/presentation.xml');
  if (!presentationXml) {
    return null;
  }

  const presentationRoot = findFirstElement(presentationXml, 'p:presentation');
  if (!presentationRoot) {
    return null;
  }

  const slideIdList = findChild(presentationRoot, 'p:sldIdLst');
  if (!slideIdList) {
    return null;
  }

  const rels = await getRelationships(pkg, 'ppt/presentation.xml');
  const changesBySlide = groupChangesBySlide(result.changes);
  const slideIds = getChildren(slideIdList).filter((node) => getTagName(node) === 'p:sldId');

  for (let i = 0; i < slideIds.length; i += 1) {
    const slideIndex = i + 1;
    const slideChanges = changesBySlide.get(slideIndex);
    if (!slideChanges) continue;

    const slideId = slideIds[i];
    const relId = getAttribute(slideId, 'r:id');
    if (!relId) continue;

    const rel = rels.find((entry) => entry.id === relId);
    if (!rel) continue;

    const slidePath = resolveRelationshipTarget('ppt/presentation.xml', rel.target);
    if (!slidePath) continue;

    try {
      const slideXml = await getPartAsXml(pkg, slidePath);
      if (!slideXml) continue;

      const slideRoot = findFirstElement(slideXml, 'p:sld');
      const commonSlide = slideRoot ? findChild(slideRoot, 'p:cSld') : null;
      const spTree = commonSlide ? findChild(commonSlide, 'p:spTree') : null;
      if (!spTree) continue;

      let nextId = getNextShapeId(spTree);
      addChangeOverlays(spTree, slideChanges, settings, () => nextId++);
      setPartFromXml(pkg, slidePath, slideXml);

      if (settings.addNotesAnnotations) {
        await addNotesAnnotations(pkg, slidePath, slideChanges);
      }
    } catch {
    }
  }

  if (settings.addSummarySlide && result.totalChanges > 0) {
    await addSummarySlide(pkg, result);
  }

  return savePackage(pkg);
}

function groupChangesBySlide(changes: PmlChange[]): Map<number, PmlChange[]> {
  const grouped = new Map<number, PmlChange[]>();

  for (const change of changes) {
    if (!change.slideIndex) continue;
    if (!grouped.has(change.slideIndex)) {
      grouped.set(change.slideIndex, []);
    }
    grouped.get(change.slideIndex)!.push(change);
  }

  return grouped;
}

function addChangeOverlays(
  spTree: XmlNode,
  changes: PmlChange[],
  settings: Required<PmlComparerSettings>,
  nextId: () => number
): void {
  const labels = getChildren(spTree);

  for (const change of changes) {
    switch (change.changeType) {
      case PmlChangeType.ShapeInserted:
        labels.push(createChangeLabel(change, 'NEW', settings.insertedColor, nextId()));
        break;
      case PmlChangeType.ShapeMoved:
        labels.push(createChangeLabel(change, 'MOVED', settings.movedColor, nextId()));
        break;
      case PmlChangeType.ShapeResized:
        labels.push(createChangeLabel(change, 'RESIZED', settings.modifiedColor, nextId()));
        break;
      case PmlChangeType.TextChanged:
        labels.push(createChangeLabel(change, 'TEXT CHANGED', settings.modifiedColor, nextId()));
        break;
      case PmlChangeType.ImageReplaced:
        labels.push(createChangeLabel(change, 'IMAGE REPLACED', settings.modifiedColor, nextId()));
        break;
      case PmlChangeType.TableContentChanged:
        labels.push(createChangeLabel(change, 'TABLE CHANGED', settings.modifiedColor, nextId()));
        break;
      case PmlChangeType.ChartDataChanged:
        labels.push(createChangeLabel(change, 'CHART CHANGED', settings.modifiedColor, nextId()));
        break;
      default:
        break;
    }
  }
}

function createChangeLabel(change: PmlChange, text: string, color: string, id: number): XmlNode {
  let x = change.newX ?? change.oldX ?? 0;
  let y = change.newY ?? change.oldY ?? 0;

  if (y > 200000) {
    y -= 200000;
  }

  const fillColor = normalizeColor(color);

  return createNode('p:sp', undefined, [
    createNode('p:nvSpPr', undefined, [
      createNode('p:cNvPr', { id: String(id), name: `Change Label: ${change.shapeName ?? ''}` }),
      createNode('p:cNvSpPr', undefined, [createNode('a:spLocks', { noGrp: '1' })]),
      createNode('p:nvPr'),
    ]),
    createNode('p:spPr', undefined, [
      createNode('a:xfrm', undefined, [
        createNode('a:off', { x: String(x), y: String(y) }),
        createNode('a:ext', { cx: '1500000', cy: '300000' }),
      ]),
      createNode('a:prstGeom', { prst: 'rect' }, [createNode('a:avLst')]),
      createNode('a:solidFill', undefined, [createNode('a:srgbClr', { val: fillColor })]),
      createNode('a:ln', { w: '12700' }, [
        createNode('a:solidFill', undefined, [createNode('a:srgbClr', { val: '000000' })]),
      ]),
    ]),
    createNode('p:txBody', undefined, [
      createNode('a:bodyPr', { wrap: 'square', rtlCol: '0', anchor: 'ctr' }),
      createNode('a:lstStyle'),
      createNode('a:p', undefined, [
        createNode('a:pPr', { algn: 'ctr' }),
        createNode('a:r', undefined, [
          createNode('a:rPr', { lang: 'en-US', sz: '1000', b: '1' }, [
            createNode('a:solidFill', undefined, [createNode('a:srgbClr', { val: 'FFFFFF' })]),
          ]),
          createNode('a:t', undefined, [text]),
        ]),
        createNode('a:endParaRPr', { lang: 'en-US' }),
      ]),
    ]),
  ]);
}

async function addNotesAnnotations(
  pkg: OoxmlPackage,
  slidePath: string,
  changes: PmlChange[]
): Promise<void> {
  const notesSlide = await getOrCreateNotesSlide(pkg, slidePath);
  if (!notesSlide) return;

  const { notesPath, notesXml } = notesSlide;
  const notesRoot = findFirstElement(notesXml, 'p:notes');
  const commonSlide = notesRoot ? findChild(notesRoot, 'p:cSld') : null;
  const spTree = commonSlide ? findChild(commonSlide, 'p:spTree') : null;
  if (!spTree) return;

  const notesShape = findNotesBodyShape(spTree);
  if (!notesShape) return;

  const txBody = findChild(notesShape, 'p:txBody');
  if (!txBody) return;

  const paragraphs = getChildren(txBody);
  paragraphs.push(createNotesParagraph(`--- Changes (${changes.length}) ---`, true));

  for (const change of changes.slice(0, 10)) {
    paragraphs.push(createNotesParagraph(`- ${describeChange(change)}`, false));
  }

  if (changes.length > 10) {
    paragraphs.push(createNotesParagraph(`... and ${changes.length - 10} more changes`, false));
  }

  setPartFromXml(pkg, notesPath, notesXml);
}

function createNotesParagraph(text: string, bold: boolean): XmlNode {
  const rPrAttrs: Record<string, string> = { lang: 'en-US' };
  if (bold) {
    rPrAttrs.b = '1';
  }

  return createNode('a:p', undefined, [
    createNode('a:r', undefined, [
      createNode('a:rPr', rPrAttrs),
      createNode('a:t', undefined, [text]),
    ]),
  ]);
}

function findNotesBodyShape(spTree: XmlNode): XmlNode | null {
  for (const node of getChildren(spTree)) {
    if (getTagName(node) !== 'p:sp') continue;

    const nvSpPr = findChild(node, 'p:nvSpPr');
    const nvPr = nvSpPr ? findChild(nvSpPr, 'p:nvPr') : null;
    const ph = nvPr ? findChild(nvPr, 'p:ph') : null;
    if (ph && getAttribute(ph, 'type') === 'body') {
      return node;
    }
  }

  return null;
}

async function getOrCreateNotesSlide(
  pkg: OoxmlPackage,
  slidePath: string
): Promise<{ notesPath: string; notesXml: XmlNode[] } | null> {
  const rels = await getRelationships(pkg, slidePath);
  const existing = rels.find((rel) => rel.type.includes('notesSlide'));

  if (existing) {
    const notesPath = resolveRelationshipTarget(slidePath, existing.target);
    if (!notesPath) return null;
    const notesXml = await getPartAsXml(pkg, notesPath);
    if (!notesXml) return null;
    return { notesPath, notesXml };
  }

  const nextIndex = getNextPartIndex(pkg, 'ppt/notesSlides/notesSlide', '.xml');
  const notesPath = `ppt/notesSlides/notesSlide${nextIndex}.xml`;
  const notesXml = [createEmptyNotesSlide()];

  setPartFromXml(pkg, notesPath, notesXml);
  await addContentTypeOverride(pkg, notesPath, NOTES_CONTENT_TYPE);
  await addRelationship(pkg, slidePath, NOTES_REL_TYPE, '../notesSlides/notesSlide' + nextIndex + '.xml');

  return { notesPath, notesXml };
}

function createEmptyNotesSlide(): XmlNode {
  return createNode('p:notes', { 'xmlns:a': A_XMLNS, 'xmlns:p': PML_XMLNS, 'xmlns:r': R_XMLNS }, [
    createNode('p:cSld', undefined, [
      createNode('p:spTree', undefined, [
        createNode('p:nvGrpSpPr', undefined, [
          createNode('p:cNvPr', { id: '1', name: '' }),
          createNode('p:cNvGrpSpPr'),
          createNode('p:nvPr'),
        ]),
        createNode('p:grpSpPr'),
        createNode('p:sp', undefined, [
          createNode('p:nvSpPr', undefined, [
            createNode('p:cNvPr', { id: '2', name: 'Notes Placeholder' }),
            createNode('p:cNvSpPr'),
            createNode('p:nvPr', undefined, [createNode('p:ph', { type: 'body', idx: '1' })]),
          ]),
          createNode('p:spPr'),
          createNode('p:txBody', undefined, [
            createNode('a:bodyPr'),
            createNode('a:lstStyle'),
            createNode('a:p', undefined, [createNode('a:endParaRPr', { lang: 'en-US' })]),
          ]),
        ]),
      ]),
    ]),
  ]);
}

async function addSummarySlide(pkg: OoxmlPackage, result: PmlComparisonResult): Promise<void> {
  const presentationXml = await getPartAsXml(pkg, 'ppt/presentation.xml');
  if (!presentationXml) return;

  const presentationRoot = findFirstElement(presentationXml, 'p:presentation');
  if (!presentationRoot) return;

  let slideIdList = findChild(presentationRoot, 'p:sldIdLst');
  if (!slideIdList) {
    slideIdList = createNode('p:sldIdLst');
    getChildren(presentationRoot).push(slideIdList);
  }

  const slideIndex = getNextPartIndex(pkg, 'ppt/slides/slide', '.xml');
  const slidePath = `ppt/slides/slide${slideIndex}.xml`;

  const slideXml = [
    createNode('p:sld', { 'xmlns:a': A_XMLNS, 'xmlns:p': PML_XMLNS, 'xmlns:r': R_XMLNS }, [
      createNode('p:cSld', undefined, [
        createNode('p:spTree', undefined, [
          createNode('p:nvGrpSpPr', undefined, [
            createNode('p:cNvPr', { id: '1', name: '' }),
            createNode('p:cNvGrpSpPr'),
            createNode('p:nvPr'),
          ]),
          createNode('p:grpSpPr'),
          createTitleShape('Comparison Summary', 2),
          createSummaryContentShape(result, 3),
        ]),
      ]),
    ]),
  ];

  setPartFromXml(pkg, slidePath, slideXml);
  await addContentTypeOverride(pkg, slidePath, SLIDE_CONTENT_TYPE);

  const rId = await addRelationship(pkg, 'ppt/presentation.xml', PRESENTATION_REL_TYPE, `slides/slide${slideIndex}.xml`);
  const slideId = createNode('p:sldId', { id: String(getNextSlideId(slideIdList)), 'r:id': rId });
  getChildren(slideIdList).push(slideId);

  setPartFromXml(pkg, 'ppt/presentation.xml', presentationXml);
}

function createTitleShape(title: string, id: number): XmlNode {
  return createNode('p:sp', undefined, [
    createNode('p:nvSpPr', undefined, [
      createNode('p:cNvPr', { id: String(id), name: 'Title' }),
      createNode('p:cNvSpPr'),
      createNode('p:nvPr'),
    ]),
    createNode('p:spPr', undefined, [
      createNode('a:xfrm', undefined, [
        createNode('a:off', { x: '457200', y: '274638' }),
        createNode('a:ext', { cx: '8229600', cy: '1143000' }),
      ]),
      createNode('a:prstGeom', { prst: 'rect' }, [createNode('a:avLst')]),
    ]),
    createNode('p:txBody', undefined, [
      createNode('a:bodyPr'),
      createNode('a:lstStyle'),
      createNode('a:p', undefined, [
        createNode('a:r', undefined, [
          createNode('a:rPr', { lang: 'en-US', sz: '4400', b: '1' }),
          createNode('a:t', undefined, [title]),
        ]),
      ]),
    ]),
  ]);
}

function createSummaryContentShape(result: PmlComparisonResult, id: number): XmlNode {
  const lines = [
    `Total Changes: ${result.totalChanges}`,
    '',
    `Slides Inserted: ${result.slidesInserted}`,
    `Slides Deleted: ${result.slidesDeleted}`,
    `Slides Moved: ${result.slidesMoved}`,
    '',
    `Shapes Inserted: ${result.shapesInserted}`,
    `Shapes Deleted: ${result.shapesDeleted}`,
    `Shapes Moved: ${result.shapesMoved}`,
    `Shapes Resized: ${result.shapesResized}`,
    '',
    `Text Changes: ${result.textChanges}`,
    `Formatting Changes: ${result.formattingChanges}`,
    `Images Replaced: ${result.imagesReplaced}`,
  ];

  const paragraphs = lines.map((line) =>
    createNode('a:p', undefined, [
      createNode('a:r', undefined, [
        createNode('a:rPr', { lang: 'en-US', sz: '2000' }),
        createNode('a:t', undefined, [line]),
      ]),
    ])
  );

  return createNode('p:sp', undefined, [
    createNode('p:nvSpPr', undefined, [
      createNode('p:cNvPr', { id: String(id), name: 'Content' }),
      createNode('p:cNvSpPr'),
      createNode('p:nvPr'),
    ]),
    createNode('p:spPr', undefined, [
      createNode('a:xfrm', undefined, [
        createNode('a:off', { x: '457200', y: '1600200' }),
        createNode('a:ext', { cx: '8229600', cy: '4525963' }),
      ]),
      createNode('a:prstGeom', { prst: 'rect' }, [createNode('a:avLst')]),
    ]),
    createNode('p:txBody', undefined, [createNode('a:bodyPr'), createNode('a:lstStyle'), ...paragraphs]),
  ]);
}

function getNextShapeId(spTree: XmlNode): number {
  let maxId = 0;
  for (const node of findByTagName(spTree, 'p:cNvPr')) {
    const id = parseInt(getAttribute(node, 'id') ?? '0', 10);
    if (id > maxId) {
      maxId = id;
    }
  }
  return maxId + 1;
}

function getNextSlideId(slideIdList: XmlNode): number {
  let maxId = 256;
  for (const node of getChildren(slideIdList)) {
    if (getTagName(node) !== 'p:sldId') continue;
    const id = parseInt(getAttribute(node, 'id') ?? '0', 10);
    if (id > maxId) {
      maxId = id;
    }
  }
  return maxId + 1;
}

async function addRelationship(
  pkg: OoxmlPackage,
  partPath: string,
  type: string,
  target: string
): Promise<string> {
  const relsPath = getRelsPath(partPath);
  const relsXml = (await getPartAsXml(pkg, relsPath)) ?? [
    createNode('Relationships', { xmlns: RELS_XMLNS }, []),
  ];

  const root = findFirstElement(relsXml, 'Relationships');
  if (!root) return 'rId1';

  let nextId = 1;
  for (const node of getChildren(root)) {
    if (getTagName(node) !== 'Relationship') continue;
    const id = getAttribute(node, 'Id');
    if (!id) continue;
    const numeric = parseInt(id.replace('rId', ''), 10);
    if (numeric >= nextId) {
      nextId = numeric + 1;
    }
  }

  const relNode = createNode('Relationship', {
    Id: `rId${nextId}`,
    Type: type,
    Target: target,
  });

  getChildren(root).push(relNode);
  setPartFromXml(pkg, relsPath, relsXml);
  return `rId${nextId}`;
}

async function addContentTypeOverride(
  pkg: OoxmlPackage,
  partPath: string,
  contentType: string
): Promise<void> {
  const contentTypesXml = await getPartAsXml(pkg, '[Content_Types].xml');
  if (!contentTypesXml) return;

  const typesNode = findFirstElement(contentTypesXml, 'Types');
  if (!typesNode) return;

  const children = getChildren(typesNode);
  const exists = children.some(
    (child) => getTagName(child) === 'Override' && getAttribute(child, 'PartName') === `/${partPath}`
  );
  if (!exists) {
    children.push(
      createNode('Override', {
        PartName: `/${partPath}`,
        ContentType: contentType,
      })
    );
  }

  setPartFromXml(pkg, '[Content_Types].xml', contentTypesXml);
}

function getNextPartIndex(pkg: OoxmlPackage, prefix: string, suffix: string): number {
  const parts = listParts(pkg);
  let maxIndex = 0;

  for (const part of parts) {
    if (!part.startsWith(prefix) || !part.endsWith(suffix)) continue;
    const numberText = part.slice(prefix.length, part.length - suffix.length);
    const numberValue = parseInt(numberText, 10);
    if (!Number.isNaN(numberValue) && numberValue > maxIndex) {
      maxIndex = numberValue;
    }
  }

  return maxIndex + 1;
}

function getRelsPath(partPath: string): string {
  const parts = partPath.split('/');
  const fileName = parts.pop()!;
  const dir = parts.join('/');
  return dir ? `${dir}/_rels/${fileName}.rels` : `_rels/${fileName}.rels`;
}

function normalizeColor(color: string): string {
  const trimmed = color.startsWith('#') ? color.slice(1) : color;
  return trimmed.length === 6 ? trimmed : trimmed.padStart(6, '0');
}

function findFirstElement(nodes: XmlNode[], tagName: string): XmlNode | null {
  for (const node of nodes) {
    if (getTagName(node) === tagName) {
      return node;
    }
    const found = findDescendant(node, tagName);
    if (found) return found;
  }
  return null;
}

function findDescendant(node: XmlNode, tagName: string): XmlNode | null {
  for (const child of getChildren(node)) {
    if (getTagName(child) === tagName) {
      return child;
    }
    const found = findDescendant(child, tagName);
    if (found) return found;
  }
  return null;
}

function findChild(node: XmlNode, tagName: string): XmlNode | null {
  for (const child of getChildren(node)) {
    if (getTagName(child) === tagName) {
      return child;
    }
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

