/**
 * XML parsing and building utilities
 *
 * Uses fast-xml-parser for high-performance XML handling.
 * Configured to preserve namespaces, attributes, and text content.
 */

import { XMLParser, XMLBuilder, type X2jOptions, type XmlBuilderOptions } from 'fast-xml-parser';

/**
 * Common parser options for OOXML documents
 */
const PARSER_OPTIONS: Partial<X2jOptions> = {
  // Preserve attribute order and namespaces
  ignoreAttributes: false,
  attributeNamePrefix: '@_',

  // Preserve text content
  textNodeName: '#text',
  cdataPropName: '#cdata',
  commentPropName: '#comment',

  // Preserve ordering
  preserveOrder: true,

  // Handle namespaces
  removeNSPrefix: false,

  // Parse numbers as strings to preserve formatting
  parseTagValue: false,
  parseAttributeValue: false,

  // Trim whitespace from text nodes
  trimValues: false,

  // Process entities
  processEntities: true,
  htmlEntities: false,

  // Allow boolean attributes
  allowBooleanAttributes: true,
};

/**
 * Common builder options for OOXML documents
 */
const BUILDER_OPTIONS: Partial<XmlBuilderOptions> = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  cdataPropName: '#cdata',
  commentPropName: '#comment',
  preserveOrder: true,
  format: false, // Don't add extra whitespace
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
};

// Shared parser and builder instances
const parser = new XMLParser(PARSER_OPTIONS);
const builder = new XMLBuilder(BUILDER_OPTIONS);

/**
 * Parse XML string to object representation
 */
export function parseXml(xml: string): XmlNode[] {
  return parser.parse(xml);
}

/**
 * Build XML string from object representation.
 * Strips any ?xml declaration nodes since addXmlDeclaration adds the declaration separately.
 */
export function buildXml(nodes: XmlNode | XmlNode[]): string {
  const nodeArray = Array.isArray(nodes) ? nodes : [nodes];
  // Filter out ?xml declaration nodes to avoid duplication when addXmlDeclaration is called
  const filteredNodes = nodeArray.filter((node) => !('?xml' in node));
  return builder.build(filteredNodes);
}

/**
 * XML node representation from fast-xml-parser with preserveOrder: true
 *
 * Structure: { tagName: [...children], ':@': { '@_attrName': 'value' } }
 * Text nodes: { '#text': 'content' }
 */
export interface XmlNode {
  [tagName: string]: XmlNode[] | XmlAttributes | string | undefined;
  ':@'?: XmlAttributes;
  '#text'?: string;
}

/**
 * XML attributes object
 */
export interface XmlAttributes {
  [attrName: string]: string;
}

/**
 * Get the tag name of an XML node (first key that isn't :@ or #text)
 */
export function getTagName(node: XmlNode): string | null {
  for (const key of Object.keys(node)) {
    if (key !== ':@' && key !== '#text') {
      return key;
    }
  }
  return null;
}

/**
 * Get child nodes of an XML node
 */
export function getChildren(node: XmlNode): XmlNode[] {
  const tagName = getTagName(node);
  if (!tagName) return [];
  const children = node[tagName];
  if (Array.isArray(children)) {
    return children as XmlNode[];
  }
  return [];
}

/**
 * Get attributes of an XML node
 */
export function getAttributes(node: XmlNode): XmlAttributes {
  return (node[':@'] as XmlAttributes) || {};
}

/**
 * Get attribute value
 */
export function getAttribute(node: XmlNode, name: string): string | undefined {
  const attrs = getAttributes(node);
  return attrs[`@_${name}`];
}

/**
 * Set attribute value
 */
export function setAttribute(node: XmlNode, name: string, value: string): void {
  if (!node[':@']) {
    node[':@'] = {};
  }
  (node[':@'] as XmlAttributes)[`@_${name}`] = value;
}

/**
 * Get text content of an XML node (recursive)
 */
export function getTextContent(node: XmlNode): string {
  // If it's a text node
  if ('#text' in node && typeof node['#text'] === 'string') {
    return node['#text'];
  }

  const tagName = getTagName(node);
  if (!tagName) return '';

  const children = node[tagName];
  if (!Array.isArray(children)) return '';

  let text = '';
  for (const child of children) {
    if (typeof child === 'object') {
      text += getTextContent(child);
    }
  }
  return text;
}

/**
 * Find all descendant nodes matching a predicate
 */
export function findNodes(
  node: XmlNode,
  predicate: (n: XmlNode) => boolean
): XmlNode[] {
  const results: XmlNode[] = [];

  function walk(n: XmlNode) {
    if (predicate(n)) {
      results.push(n);
    }
    for (const child of getChildren(n)) {
      walk(child);
    }
  }

  walk(node);
  return results;
}

/**
 * Find all descendant nodes with a specific tag name
 */
export function findByTagName(node: XmlNode, tagName: string): XmlNode[] {
  return findNodes(node, (n) => getTagName(n) === tagName);
}

/**
 * Clone an XML node (deep copy)
 */
export function cloneNode(node: XmlNode): XmlNode {
  return JSON.parse(JSON.stringify(node));
}

/**
 * Create a new XML node
 */
export function createNode(
  tagName: string,
  attrs?: XmlAttributes,
  children?: (XmlNode | string)[]
): XmlNode {
  const node: XmlNode = {
    [tagName]: [],
  };

  if (attrs) {
    node[':@'] = {};
    for (const [key, value] of Object.entries(attrs)) {
      (node[':@'] as XmlAttributes)[`@_${key}`] = value;
    }
  }

  if (children) {
    const childArray = node[tagName] as XmlNode[];
    for (const child of children) {
      if (typeof child === 'string') {
        childArray.push({ '#text': child });
      } else {
        childArray.push(child);
      }
    }
  }

  return node;
}

/**
 * Add XML declaration to output
 */
export function addXmlDeclaration(xml: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
}
