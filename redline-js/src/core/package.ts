/**
 * OOXML Package handling utilities
 *
 * Office documents (.docx, .xlsx, .pptx) are ZIP archives containing XML files.
 * This module provides utilities for reading and writing these packages.
 */

import JSZip from 'jszip';
import { parseXml, buildXml, addXmlDeclaration, type XmlNode } from './xml';

/**
 * Represents an OOXML package (ZIP archive)
 */
export interface OoxmlPackage {
  /** The JSZip instance */
  zip: JSZip;
  /** Parsed content types */
  contentTypes: ContentTypes;
  /** File type determined from content types */
  fileType: 'word' | 'excel' | 'powerpoint' | 'unknown';
}

/**
 * Content types from [Content_Types].xml
 */
export interface ContentTypes {
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}

/**
 * Open an OOXML package from a buffer
 */
export async function openPackage(
  data: Buffer | Uint8Array | ArrayBuffer
): Promise<OoxmlPackage> {
  const zip = await JSZip.loadAsync(data);

  // Parse content types
  const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');
  if (!contentTypesXml) {
    throw new Error('Invalid OOXML package: missing [Content_Types].xml');
  }

  const contentTypes = parseContentTypes(contentTypesXml);
  const fileType = determineFileType(contentTypes);

  return { zip, contentTypes, fileType };
}

/**
 * Parse [Content_Types].xml
 *
 * With fast-xml-parser preserveOrder: true, structure is:
 * [ { Types: [ { Default: [], ':@': { '@_Extension': '...', '@_ContentType': '...' } } ], ':@': {...} } ]
 */
function parseContentTypes(xml: string): ContentTypes {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();

  const nodes = parseXml(xml);

  // Find the Types element
  for (const node of nodes) {
    if ('Types' in node) {
      const children = node['Types'] as XmlNode[];
      for (const child of children) {
        // With preserveOrder, attributes are at same level as tag in ':@'
        const attrs = child[':@'] as Record<string, string> | undefined;
        if (!attrs) continue;

        if ('Default' in child) {
          const ext = attrs['@_Extension'];
          const type = attrs['@_ContentType'];
          if (ext && type) defaults.set(ext, type);
        } else if ('Override' in child) {
          const partName = attrs['@_PartName'];
          const type = attrs['@_ContentType'];
          if (partName && type) overrides.set(partName, type);
        }
      }
    }
  }

  return { defaults, overrides };
}

/**
 * Determine file type from content types
 */
function determineFileType(
  contentTypes: ContentTypes
): 'word' | 'excel' | 'powerpoint' | 'unknown' {
  for (const [, contentType] of contentTypes.overrides) {
    if (contentType.includes('wordprocessingml')) return 'word';
    if (contentType.includes('spreadsheetml')) return 'excel';
    if (contentType.includes('presentationml')) return 'powerpoint';
  }
  return 'unknown';
}

/**
 * Get a part from the package as a string
 */
export async function getPartAsString(
  pkg: OoxmlPackage,
  partPath: string
): Promise<string | null> {
  const file = pkg.zip.file(partPath);
  if (!file) return null;
  return file.async('string');
}

/**
 * Get a part from the package as parsed XML
 */
export async function getPartAsXml(
  pkg: OoxmlPackage,
  partPath: string
): Promise<XmlNode[] | null> {
  const content = await getPartAsString(pkg, partPath);
  if (!content) return null;
  return parseXml(content);
}

/**
 * Set a part in the package from a string
 */
export function setPartFromString(
  pkg: OoxmlPackage,
  partPath: string,
  content: string
): void {
  pkg.zip.file(partPath, content);
}

/**
 * Set a part in the package from XML nodes
 */
export function setPartFromXml(
  pkg: OoxmlPackage,
  partPath: string,
  nodes: XmlNode | XmlNode[]
): void {
  const xml = addXmlDeclaration(buildXml(nodes));
  setPartFromString(pkg, partPath, xml);
}

/**
 * Save the package to a buffer
 */
export async function savePackage(pkg: OoxmlPackage): Promise<Buffer> {
  return pkg.zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
  });
}

/**
 * Clone a package (deep copy)
 */
export async function clonePackage(pkg: OoxmlPackage): Promise<OoxmlPackage> {
  const buffer = await savePackage(pkg);
  return openPackage(buffer);
}

/**
 * List all parts in the package
 */
export function listParts(pkg: OoxmlPackage): string[] {
  const parts: string[] = [];
  pkg.zip.forEach((relativePath) => {
    parts.push(relativePath);
  });
  return parts;
}

/**
 * Check if a part exists
 */
export function hasPart(pkg: OoxmlPackage, partPath: string): boolean {
  return pkg.zip.file(partPath) !== null;
}

/**
 * Get relationships from a .rels file
 */
export async function getRelationships(
  pkg: OoxmlPackage,
  partPath: string
): Promise<Relationship[]> {
  // Convert part path to rels path
  // e.g., "word/document.xml" -> "word/_rels/document.xml.rels"
  const parts = partPath.split('/');
  const fileName = parts.pop()!;
  const dir = parts.join('/');
  const relsPath = dir ? `${dir}/_rels/${fileName}.rels` : `_rels/${fileName}.rels`;

  const relsXml = await getPartAsString(pkg, relsPath);
  if (!relsXml) return [];

  const nodes = parseXml(relsXml);
  const relationships: Relationship[] = [];

  for (const node of nodes) {
    if ('Relationships' in node) {
      const children = node['Relationships'] as XmlNode[];
      for (const child of children) {
        const attrs = child[':@'] as Record<string, string> | undefined;
        if (!attrs || !('Relationship' in child)) continue;

        relationships.push({
          id: attrs['@_Id'] || '',
          type: attrs['@_Type'] || '',
          target: attrs['@_Target'] || '',
          targetMode: attrs['@_TargetMode'] as 'Internal' | 'External' | undefined,
        });
      }
    }
  }

  return relationships;
}

/**
 * A relationship entry from a .rels file
 */
export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: 'Internal' | 'External';
}

// Common relationship types
export const REL_TYPES = {
  OFFICE_DOCUMENT:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  STYLES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  NUMBERING: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  FONT_TABLE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  FOOTNOTES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  ENDNOTES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  COMMENTS: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  SETTINGS: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  IMAGE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  HYPERLINK: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  HEADER: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  FOOTER: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  CORE_PROPERTIES:
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  EXTENDED_PROPERTIES:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
} as const;
