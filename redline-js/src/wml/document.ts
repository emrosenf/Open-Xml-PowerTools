/**
 * Word document handling utilities
 *
 * Provides functions for loading, extracting text, and manipulating Word documents.
 */

import {
  openPackage,
  getPartAsXml,
  getPartAsString,
  setPartFromXml,
  savePackage,
  clonePackage,
  getRelationships,
  type OoxmlPackage,
  type Relationship,
} from '../core/package';
import {
  getTagName,
  getChildren,
  getTextContent,
  findNodes,
  type XmlNode,
} from '../core/xml';

/**
 * Represents a Word document (.docx)
 */
export interface WordDocument {
  /** The underlying OOXML package */
  package: OoxmlPackage;
  /** The main document part (word/document.xml) */
  mainDocument: XmlNode[];
  /** Document relationships */
  relationships: Relationship[];
  /** Styles part if present */
  styles?: XmlNode[];
  /** Numbering part if present */
  numbering?: XmlNode[];
  /** Footnotes part if present */
  footnotes?: XmlNode[];
  /** Endnotes part if present */
  endnotes?: XmlNode[];
  /** Core properties (author, date, etc.) */
  coreProperties?: CoreProperties;
}

/**
 * Core document properties
 */
export interface CoreProperties {
  creator?: string;
  lastModifiedBy?: string;
  created?: string;
  modified?: string;
  title?: string;
  subject?: string;
}

/**
 * Load a Word document from a buffer
 */
export async function loadWordDocument(
  data: Buffer | Uint8Array | ArrayBuffer
): Promise<WordDocument> {
  const pkg = await openPackage(data);

  if (pkg.fileType !== 'word') {
    throw new Error('Not a Word document');
  }

  // Load main document
  const mainDocument = await getPartAsXml(pkg, 'word/document.xml');
  if (!mainDocument) {
    throw new Error('Invalid Word document: missing word/document.xml');
  }

  // Load relationships
  const relationships = await getRelationships(pkg, 'word/document.xml');

  // Load optional parts
  const styles = await getPartAsXml(pkg, 'word/styles.xml') ?? undefined;
  const numbering = await getPartAsXml(pkg, 'word/numbering.xml') ?? undefined;
  const footnotes = await getPartAsXml(pkg, 'word/footnotes.xml') ?? undefined;
  const endnotes = await getPartAsXml(pkg, 'word/endnotes.xml') ?? undefined;

  // Load core properties
  const coreProperties = await loadCoreProperties(pkg);

  return {
    package: pkg,
    mainDocument,
    relationships,
    styles,
    numbering,
    footnotes,
    endnotes,
    coreProperties,
  };
}

/**
 * Load core properties from docProps/core.xml
 */
async function loadCoreProperties(pkg: OoxmlPackage): Promise<CoreProperties | undefined> {
  const coreXml = await getPartAsString(pkg, 'docProps/core.xml');
  if (!coreXml) return undefined;

  const props: CoreProperties = {};

  // Simple regex extraction for core properties
  const creatorMatch = coreXml.match(/<dc:creator>([^<]*)<\/dc:creator>/);
  if (creatorMatch) props.creator = creatorMatch[1];

  const lastModifiedByMatch = coreXml.match(/<cp:lastModifiedBy>([^<]*)<\/cp:lastModifiedBy>/);
  if (lastModifiedByMatch) props.lastModifiedBy = lastModifiedByMatch[1];

  const createdMatch = coreXml.match(/<dcterms:created[^>]*>([^<]*)<\/dcterms:created>/);
  if (createdMatch) props.created = createdMatch[1];

  const modifiedMatch = coreXml.match(/<dcterms:modified[^>]*>([^<]*)<\/dcterms:modified>/);
  if (modifiedMatch) props.modified = modifiedMatch[1];

  const titleMatch = coreXml.match(/<dc:title>([^<]*)<\/dc:title>/);
  if (titleMatch) props.title = titleMatch[1];

  const subjectMatch = coreXml.match(/<dc:subject>([^<]*)<\/dc:subject>/);
  if (subjectMatch) props.subject = subjectMatch[1];

  return props;
}

/**
 * Save a Word document to a buffer
 */
export async function saveWordDocument(doc: WordDocument): Promise<Buffer> {
  // Update main document
  setPartFromXml(doc.package, 'word/document.xml', doc.mainDocument);

  return savePackage(doc.package);
}

/**
 * Clone a Word document
 */
export async function cloneWordDocument(doc: WordDocument): Promise<WordDocument> {
  const pkg = await clonePackage(doc.package);
  return loadWordDocument(await savePackage(pkg));
}

/**
 * Get the document body element
 */
export function getDocumentBody(doc: WordDocument): XmlNode | null {
  for (const node of doc.mainDocument) {
    const tagName = getTagName(node);
    if (tagName === 'w:document') {
      const children = getChildren(node);
      for (const child of children) {
        if (getTagName(child) === 'w:body') {
          return child;
        }
      }
    }
  }
  return null;
}

/**
 * Extract all text content from a Word document
 */
export function extractText(doc: WordDocument): string {
  const body = getDocumentBody(doc);
  if (!body) return '';

  const texts: string[] = [];

  function walkNode(node: XmlNode): void {
    const tagName = getTagName(node);

    // Text run content
    if (tagName === 'w:t') {
      texts.push(getTextContent(node));
      return;
    }

    // Paragraph break
    if (tagName === 'w:p') {
      const beforeLen = texts.length;
      for (const child of getChildren(node)) {
        walkNode(child);
      }
      // Add paragraph separator if we added any text
      if (texts.length > beforeLen) {
        texts.push('\n');
      }
      return;
    }

    // Line break
    if (tagName === 'w:br') {
      texts.push('\n');
      return;
    }

    // Tab
    if (tagName === 'w:tab') {
      texts.push('\t');
      return;
    }

    // Recurse into children
    for (const child of getChildren(node)) {
      walkNode(child);
    }
  }

  walkNode(body);

  return texts.join('').trim();
}

/**
 * Extract paragraphs as separate text strings
 */
export function extractParagraphs(doc: WordDocument): string[] {
  const body = getDocumentBody(doc);
  if (!body) return [];

  const paragraphs: string[] = [];

  for (const child of getChildren(body)) {
    if (getTagName(child) === 'w:p') {
      const text = extractParagraphText(child);
      paragraphs.push(text);
    }
  }

  return paragraphs;
}

/**
 * Extract text from a single paragraph.
 * Skips deleted text (w:del elements) when acceptRevisions is true.
 * Includes text from textboxes (w:txbxContent) within the paragraph.
 * Handles mc:AlternateContent by preferring mc:Fallback (VML) content.
 *
 * @param paragraph The paragraph node to extract text from
 * @param acceptRevisions If true, skips deleted text and includes inserted text
 */
export function extractParagraphText(paragraph: XmlNode, acceptRevisions = true): string {
  const texts: string[] = [];

  function walkNode(node: XmlNode): void {
    const tagName = getTagName(node);

    // Skip deleted content when accepting revisions
    if (acceptRevisions && tagName === 'w:del') {
      return;
    }

    if (tagName === 'w:t') {
      texts.push(getTextContent(node));
      return;
    }

    // Math elements (m:oMath, m:oMathPara) - treat as single atomic units like drawings
    // The C# code treats these as single atoms, not extracting individual characters
    if (tagName === 'm:oMath' || tagName === 'm:oMathPara') {
      const mathHash = getMathHash(node);
      texts.push(` MATH_${mathHash} `);
      return;
    }

    // Skip w:delText (deleted text marker) - never include
    if (tagName === 'w:delText') {
      return;
    }

    if (tagName === 'w:br') {
      texts.push('\n');
      return;
    }

    if (tagName === 'w:tab') {
      texts.push('\t');
      return;
    }

    // Include footnote/endnote references as markers for comparison
    // This allows detecting when references are added/removed
    // Add spaces around to make them separate tokens
    if (tagName === 'w:footnoteReference') {
      const attrs = node[':@'] as Record<string, string> | undefined;
      const id = attrs?.['@_w:id'] || 'unknown';
      texts.push(` FOOTNOTE_REF_${id} `);
      return;
    }

    if (tagName === 'w:endnoteReference') {
      const attrs = node[':@'] as Record<string, string> | undefined;
      const id = attrs?.['@_w:id'] || 'unknown';
      texts.push(` ENDNOTE_REF_${id} `);
      return;
    }

    // Handle mc:AlternateContent - prefer mc:Fallback (VML) for textboxes
    if (tagName === 'mc:AlternateContent') {
      const children = getChildren(node);
      // Look for mc:Fallback first (contains VML textbox)
      const fallback = children.find((c) => getTagName(c) === 'mc:Fallback');
      if (fallback) {
        walkNode(fallback);
        return;
      }
      // Otherwise use mc:Choice
      const choice = children.find((c) => getTagName(c) === 'mc:Choice');
      if (choice) {
        walkNode(choice);
        return;
      }
    }

    // Include text from textboxes with a separator to prevent change grouping
    if (tagName === 'w:txbxContent') {
      texts.push(' TXBX_START ');
      for (const child of getChildren(node)) {
        walkNode(child);
      }
      texts.push(' TXBX_END ');
      return;
    }

    // Handle drawings (DrawingML) - include content hash for comparison
    // We use a combination of dimensions, docPr, and embed reference since
    // we don't have easy access to compute SHA1 of image binary
    if (tagName === 'w:drawing') {
      const drawingInfo = getDrawingInfo(node);
      texts.push(drawingInfo);
      return;
    }

    // Handle w:pict - could be VML image or textbox
    // Check if it contains a textbox (v:textbox) - if so, process content
    // Otherwise treat as an image reference
    if (tagName === 'w:pict') {
      const hasTextbox = findNodes(node, (n) => getTagName(n) === 'v:textbox').length > 0;
      if (hasTextbox) {
        // Process textbox content (v:textbox > w:txbxContent)
        for (const child of getChildren(node)) {
          walkNode(child);
        }
        return;
      }
      // VML image - get embed reference
      const embedRef = findEmbedReference(node);
      if (embedRef) {
        texts.push(`PICT_${embedRef}`);
      } else {
        texts.push('PICT_unknown');
      }
      return;
    }

    for (const child of getChildren(node)) {
      walkNode(child);
    }
  }

  // Helper to get drawing info for comparison
  // Uses dimensions and embed reference to create unique identifier
  // Note: docPr id is NOT included as it varies between documents for same content
  // Format uses underscores so tokenizer treats it as single word
  function getDrawingInfo(node: XmlNode): string {
    const parts: string[] = [];

    // Find extent (dimensions)
    const extent = findNodeByTag(node, 'wp:extent');
    if (extent) {
      const attrs = extent[':@'] as Record<string, string> | undefined;
      if (attrs?.['@_cx']) parts.push(`cx${attrs['@_cx']}`);
      if (attrs?.['@_cy']) parts.push(`cy${attrs['@_cy']}`);
    }

    // Find embed reference (the actual image/content reference)
    const embedRef = findEmbedReference(node);
    if (embedRef) parts.push(`e${embedRef}`);

    // Join with underscore to make it a single token
    return 'DRAWING_' + (parts.join('_') || 'unknown');
  }

  // Helper to find a node by tag name recursively
  function findNodeByTag(node: XmlNode, tag: string): XmlNode | null {
    if (getTagName(node) === tag) return node;
    for (const child of getChildren(node)) {
      const found = findNodeByTag(child, tag);
      if (found) return found;
    }
    return null;
  }

  // Helper to find embed reference in drawing
  function findEmbedReference(node: XmlNode): string | null {
    const attrs = node[':@'] as Record<string, string> | undefined;

    // Check for r:embed attribute (DrawingML images)
    if (attrs?.['@_r:embed']) {
      return attrs['@_r:embed'];
    }

    // Check for o:relid attribute (VML images)
    if (attrs?.['@_o:relid']) {
      return attrs['@_o:relid'];
    }

    // Check for r:id attribute (various references)
    if (attrs?.['@_r:id']) {
      return attrs['@_r:id'];
    }

    // Recurse into children to find a:blip or v:imagedata with embed
    for (const child of getChildren(node)) {
      const childTag = getTagName(child);

      // DrawingML blip
      if (childTag === 'a:blip') {
        const childAttrs = child[':@'] as Record<string, string> | undefined;
        if (childAttrs?.['@_r:embed']) {
          return childAttrs['@_r:embed'];
        }
      }

      // VML imagedata
      if (childTag === 'v:imagedata') {
        const childAttrs = child[':@'] as Record<string, string> | undefined;
        if (childAttrs?.['@_r:id']) {
          return childAttrs['@_r:id'];
        }
      }

      // Recurse
      const found = findEmbedReference(child);
      if (found) return found;
    }

    return null;
  }

  function getMathHash(node: XmlNode): string {
    const mathTexts: string[] = [];
    function extractMathText(n: XmlNode): void {
      if (getTagName(n) === 'm:t') {
        mathTexts.push(getTextContent(n));
      }
      for (const child of getChildren(n)) {
        extractMathText(child);
      }
    }
    extractMathText(node);
    const content = mathTexts.join('');
    let hash = 0;
    for (let i = 0; i < content.length; i++) {
      hash = ((hash << 5) - hash + content.charCodeAt(i)) | 0;
    }
    return Math.abs(hash).toString(16);
  }

  walkNode(paragraph);
  return texts.join('');
}

/**
 * Find all paragraphs in a document part
 */
export function findParagraphs(nodes: XmlNode[]): XmlNode[] {
  const paragraphs: XmlNode[] = [];

  function walk(node: XmlNode) {
    if (getTagName(node) === 'w:p') {
      paragraphs.push(node);
    }
    for (const child of getChildren(node)) {
      walk(child);
    }
  }

  for (const node of nodes) {
    walk(node);
  }

  return paragraphs;
}

/**
 * Find all runs (w:r) in a node
 */
export function findRuns(node: XmlNode): XmlNode[] {
  return findNodes(node, (n) => getTagName(n) === 'w:r');
}

/**
 * Check if a node is inside a tracked change (w:ins or w:del)
 */
export function isInsideTrackedChange(_node: XmlNode, ancestors: XmlNode[]): boolean {
  for (const ancestor of ancestors) {
    const tagName = getTagName(ancestor);
    if (tagName === 'w:ins' || tagName === 'w:del') {
      return true;
    }
  }
  return false;
}
