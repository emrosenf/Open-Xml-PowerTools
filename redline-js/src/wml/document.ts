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
 * Extract text from a single paragraph
 */
export function extractParagraphText(paragraph: XmlNode): string {
  const texts: string[] = [];

  function walkNode(node: XmlNode): void {
    const tagName = getTagName(node);

    if (tagName === 'w:t') {
      texts.push(getTextContent(node));
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

    for (const child of getChildren(node)) {
      walkNode(child);
    }
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
