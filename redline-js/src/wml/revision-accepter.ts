/**
 * Revision Accepter - Accept tracked changes in Word documents
 *
 * TypeScript port of C# RevisionProcessor.AcceptRevisions() from Open-Xml-PowerTools.
 * Used before comparing documents to work with final content, not revision markup.
 *
 * Key transformations:
 * - w:ins → unwrap (keep content)
 * - w:del, w:moveFrom → remove
 * - w:moveTo → unwrap
 * - Property changes (w:rPrChange, w:pPrChange, etc.) → remove
 * - Range markers → remove
 * - Deleted table rows → remove
 */

import {
  getTagName,
  getChildren,
  cloneNode,
  type XmlNode,
} from '../core/xml';
import type { WordDocument } from './document';

const ELEMENTS_TO_REMOVE = new Set([
  'w:del',
  'w:delText',
  'w:delInstrText',
  'w:moveFrom',
  'w:pPrChange',
  'w:rPrChange',
  'w:tblPrChange',
  'w:tblGridChange',
  'w:tcPrChange',
  'w:trPrChange',
  'w:tblPrExChange',
  'w:sectPrChange',
  'w:numberingChange',
  'w:cellIns',
  'w:customXmlDelRangeStart',
  'w:customXmlDelRangeEnd',
  'w:customXmlInsRangeStart',
  'w:customXmlInsRangeEnd',
  'w:customXmlMoveFromRangeStart',
  'w:customXmlMoveFromRangeEnd',
  'w:customXmlMoveToRangeStart',
  'w:customXmlMoveToRangeEnd',
  'w:moveFromRangeStart',
  'w:moveFromRangeEnd',
  'w:moveToRangeStart',
  'w:moveToRangeEnd',
]);

const ELEMENTS_TO_UNWRAP = new Set([
  'w:ins',
  'w:moveTo',
]);

export function acceptRevisions(doc: WordDocument): WordDocument {
  return {
    ...doc,
    mainDocument: acceptRevisionsInNodes(doc.mainDocument),
    footnotes: doc.footnotes ? acceptRevisionsInNodes(doc.footnotes) : undefined,
    endnotes: doc.endnotes ? acceptRevisionsInNodes(doc.endnotes) : undefined,
    styles: doc.styles ? acceptRevisionsInNodes(doc.styles) : undefined,
  };
}

export function acceptRevisionsInNodes(nodes: XmlNode[]): XmlNode[] {
  const result: XmlNode[] = [];
  
  for (const node of nodes) {
    const transformed = acceptRevisionsTransform(node);
    if (transformed !== null) {
      if (Array.isArray(transformed)) {
        result.push(...transformed);
      } else {
        result.push(transformed);
      }
    }
  }
  
  return result;
}

function acceptRevisionsTransform(node: XmlNode): XmlNode | XmlNode[] | null {
  const tagName = getTagName(node);
  
  if (!tagName) {
    return cloneNode(node);
  }
  
  if (ELEMENTS_TO_REMOVE.has(tagName)) {
    return null;
  }
  
  if (ELEMENTS_TO_UNWRAP.has(tagName)) {
    const children = getChildren(node);
    if (children.length === 0) {
      return null;
    }
    const transformedChildren: XmlNode[] = [];
    for (const child of children) {
      const transformed = acceptRevisionsTransform(child);
      if (transformed !== null) {
        if (Array.isArray(transformed)) {
          transformedChildren.push(...transformed);
        } else {
          transformedChildren.push(transformed);
        }
      }
    }
    return transformedChildren.length > 0 ? transformedChildren : null;
  }
  
  if (tagName === 'w:tr' && isDeletedTableRow(node)) {
    return null;
  }
  
  if (tagName === 'm:f' && hasDeletedMathControl(node)) {
    return null;
  }
  
  const children = getChildren(node);
  if (children.length === 0) {
    return cloneNode(node);
  }
  
  const transformedChildren: XmlNode[] = [];
  for (const child of children) {
    const transformed = acceptRevisionsTransform(child);
    if (transformed !== null) {
      if (Array.isArray(transformed)) {
        transformedChildren.push(...transformed);
      } else {
        transformedChildren.push(transformed);
      }
    }
  }
  
  const newNode: XmlNode = {
    [tagName]: transformedChildren,
  };
  
  if (node[':@']) {
    const attrs = node[':@'] as Record<string, string>;
    const newAttrs: Record<string, string> = {};
    let hasNonRsidAttrs = false;
    
    for (const [key, value] of Object.entries(attrs)) {
      if (key.startsWith('@_w:rsid') || key === '@_w14:paraId' || key === '@_w14:textId') {
        continue;
      }
      newAttrs[key] = value;
      hasNonRsidAttrs = true;
    }
    
    if (hasNonRsidAttrs) {
      newNode[':@'] = newAttrs;
    }
  }
  
  return newNode;
}

function isDeletedTableRow(trNode: XmlNode): boolean {
  for (const child of getChildren(trNode)) {
    if (getTagName(child) === 'w:trPr') {
      for (const prChild of getChildren(child)) {
        if (getTagName(prChild) === 'w:del') {
          return true;
        }
      }
    }
  }
  return false;
}

function hasDeletedMathControl(mfNode: XmlNode): boolean {
  for (const child of getChildren(mfNode)) {
    if (getTagName(child) === 'm:fPr') {
      for (const fPrChild of getChildren(child)) {
        if (getTagName(fPrChild) === 'm:ctrlPr') {
          for (const ctrlChild of getChildren(fPrChild)) {
            if (getTagName(ctrlChild) === 'w:del') {
              return true;
            }
          }
        }
      }
    }
  }
  return false;
}
