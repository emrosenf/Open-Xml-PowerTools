/**
 * Revision markup generation for Word documents
 *
 * Creates w:ins (insertion) and w:del (deletion) elements with proper
 * tracking attributes (author, date, id).
 */

import {
  cloneNode,
  getTagName,
  getChildren,
  type XmlNode,
} from '../core/xml';

/**
 * Settings for revision tracking
 */
export interface RevisionSettings {
  /** Author name for tracked changes */
  author: string;
  /** Date for tracked changes (ISO 8601 format) */
  dateTime: string;
}

/**
 * Default revision settings
 */
export const DEFAULT_REVISION_SETTINGS: RevisionSettings = {
  author: 'redline-js',
  dateTime: new Date().toISOString(),
};

// Counter for unique revision IDs
let revisionIdCounter = 1;

/**
 * Reset the revision ID counter (useful for testing)
 */
export function resetRevisionIdCounter(value = 1): void {
  revisionIdCounter = value;
}

/**
 * Get the next unique revision ID
 */
export function getNextRevisionId(): number {
  return revisionIdCounter++;
}

/**
 * Create a w:ins (insertion) element wrapping content
 *
 * @param content The content to wrap (typically w:r elements)
 * @param settings Revision tracking settings
 * @returns The w:ins element
 */
export function createInsertion(
  content: XmlNode | XmlNode[],
  settings: RevisionSettings = DEFAULT_REVISION_SETTINGS
): XmlNode {
  const children = Array.isArray(content) ? content : [content];

  const insNode: XmlNode = {
    'w:ins': children.map(cloneNode),
    ':@': {
      '@_w:author': settings.author,
      '@_w:id': String(getNextRevisionId()),
      '@_w:date': settings.dateTime,
    },
  };

  return insNode;
}

/**
 * Create a w:del (deletion) element wrapping content
 *
 * @param content The content to wrap (typically w:r elements)
 * @param settings Revision tracking settings
 * @returns The w:del element
 */
export function createDeletion(
  content: XmlNode | XmlNode[],
  settings: RevisionSettings = DEFAULT_REVISION_SETTINGS
): XmlNode {
  const children = Array.isArray(content) ? content : [content];

  // For deletions, we need to convert w:t to w:delText
  const convertedChildren = children.map((child) => convertToDeletedContent(cloneNode(child)));

  const delNode: XmlNode = {
    'w:del': convertedChildren,
    ':@': {
      '@_w:author': settings.author,
      '@_w:id': String(getNextRevisionId()),
      '@_w:date': settings.dateTime,
    },
  };

  return delNode;
}

/**
 * Convert content for deletion (w:t -> w:delText)
 *
 * In Word, deleted text uses w:delText instead of w:t.
 * This function recursively converts text elements.
 */
function convertToDeletedContent(node: XmlNode): XmlNode {
  const tagName = getTagName(node);

  // Convert w:t to w:delText
  if (tagName === 'w:t') {
    const children = node['w:t'];
    const attrs = node[':@'];

    const delTextNode: XmlNode = {
      'w:delText': children,
    };
    if (attrs) {
      delTextNode[':@'] = attrs;
    }
    return delTextNode;
  }

  // Recursively convert children
  if (tagName) {
    const children = getChildren(node);
    if (children.length > 0) {
      node[tagName] = children.map((child) => convertToDeletedContent(child));
    }
  }

  return node;
}

/**
 * Create a w:rPrChange element for run property changes
 */
export function createRunPropertyChange(
  originalProperties: XmlNode,
  settings: RevisionSettings = DEFAULT_REVISION_SETTINGS
): XmlNode {
  return {
    'w:rPrChange': [cloneNode(originalProperties)],
    ':@': {
      '@_w:author': settings.author,
      '@_w:id': String(getNextRevisionId()),
      '@_w:date': settings.dateTime,
    },
  };
}

/**
 * Create a w:pPrChange element for paragraph property changes
 */
export function createParagraphPropertyChange(
  originalProperties: XmlNode,
  settings: RevisionSettings = DEFAULT_REVISION_SETTINGS
): XmlNode {
  return {
    'w:pPrChange': [cloneNode(originalProperties)],
    ':@': {
      '@_w:author': settings.author,
      '@_w:id': String(getNextRevisionId()),
      '@_w:date': settings.dateTime,
    },
  };
}

/**
 * Create a run (w:r) element with text content
 */
export function createRun(text: string, properties?: XmlNode): XmlNode {
  const children: XmlNode[] = [];

  if (properties) {
    children.push(cloneNode(properties));
  }

  children.push({
    'w:t': [{ '#text': text }],
    ':@': text.startsWith(' ') || text.endsWith(' ') ? { '@_xml:space': 'preserve' } : undefined,
  });

  return {
    'w:r': children,
  };
}

/**
 * Create a paragraph (w:p) element
 */
export function createParagraph(runs: XmlNode[], properties?: XmlNode): XmlNode {
  const children: XmlNode[] = [];

  if (properties) {
    children.push(cloneNode(properties));
  }

  children.push(...runs.map(cloneNode));

  return {
    'w:p': children,
  };
}

/**
 * Wrap content in an insertion or deletion based on status
 */
export function wrapWithRevision(
  content: XmlNode | XmlNode[],
  status: 'inserted' | 'deleted',
  settings: RevisionSettings = DEFAULT_REVISION_SETTINGS
): XmlNode {
  if (status === 'inserted') {
    return createInsertion(content, settings);
  } else {
    return createDeletion(content, settings);
  }
}

/**
 * Check if a node is a revision element (w:ins or w:del)
 */
export function isRevisionElement(node: XmlNode): boolean {
  const tagName = getTagName(node);
  return tagName === 'w:ins' || tagName === 'w:del';
}

/**
 * Check if a node is an insertion
 */
export function isInsertion(node: XmlNode): boolean {
  return getTagName(node) === 'w:ins';
}

/**
 * Check if a node is a deletion
 */
export function isDeletion(node: XmlNode): boolean {
  return getTagName(node) === 'w:del';
}

/**
 * Check if a node is a format change (w:rPrChange or w:pPrChange)
 */
export function isFormatChange(node: XmlNode): boolean {
  const tagName = getTagName(node);
  return tagName === 'w:rPrChange' || tagName === 'w:pPrChange';
}

/**
 * Count revisions in a document tree
 */
export function countRevisions(nodes: XmlNode | XmlNode[]): {
  insertions: number;
  deletions: number;
  formatChanges: number;
  total: number;
} {
  const nodeArray = Array.isArray(nodes) ? nodes : [nodes];
  let insertions = 0;
  let deletions = 0;
  let formatChanges = 0;

  function walk(node: XmlNode) {
    if (isInsertion(node)) {
      insertions++;
    } else if (isDeletion(node)) {
      deletions++;
    } else if (isFormatChange(node)) {
      formatChanges++;
    }

    for (const child of getChildren(node)) {
      walk(child);
    }
  }

  for (const node of nodeArray) {
    walk(node);
  }

  return {
    insertions,
    deletions,
    formatChanges,
    total: insertions + deletions + formatChanges,
  };
}
