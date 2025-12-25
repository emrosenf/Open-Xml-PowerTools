/**
 * Office Open XML (OOXML) Namespace definitions
 *
 * These are the XML namespaces used in Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) files.
 * Ported from PtOpenXmlUtil.cs.
 */

// Word Processing ML
export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
export const W16 = 'http://schemas.microsoft.com/office/word/2018/wordml';
export const W16_CE = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';
export const W16_CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
export const W16_SE = 'http://schemas.microsoft.com/office/word/2015/wordml/symex';
export const W16_SDT = 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash';

// Spreadsheet ML
export const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

// Presentation ML
export const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
export const P15 = 'http://schemas.microsoft.com/office/powerpoint/2012/main';

// Drawing ML
export const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
export const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
export const DGM14 = 'http://schemas.microsoft.com/office/drawing/2010/diagram';
export const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
export const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
export const WP14 = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';
export const WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
export const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
export const WPC = 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas';
export const WPI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk';

// Picture/Image
export const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
export const PIC14 = 'http://schemas.microsoft.com/office/drawing/2010/picture';

// Charts
export const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
export const C14 = 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart';
export const C15 = 'http://schemas.microsoft.com/office/drawing/2012/chart';
export const C16 = 'http://schemas.microsoft.com/office/drawing/2014/chart';
export const C16R3 = 'http://schemas.microsoft.com/office/drawing/2017/03/chart';
export const CS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';

// Relationships
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// Compatibility
export const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

// Extended Properties
export const EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
export const CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';

// Math
export const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

// VML (Vector Markup Language - legacy)
export const V = 'urn:schemas-microsoft-com:vml';
export const VE = 'http://schemas.openxmlformats.org/officeDocument/2006/vmlDrawingExtensions';
export const O = 'urn:schemas-microsoft-com:office:office';
export const OX = 'urn:schemas-microsoft-com:office:excel';
export const OW = 'urn:schemas-microsoft-com:office:word';
export const X = 'urn:schemas-microsoft-com:office:powerpoint';

// Content Types
export const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

// Dublin Core
export const DC = 'http://purl.org/dc/elements/1.1/';
export const DCTERMS = 'http://purl.org/dc/terms/';
export const DCMITYPE = 'http://purl.org/dc/dcmitype/';

// XML Schema
export const XSI = 'http://www.w3.org/2001/XMLSchema-instance';

// SVG
export const SVG = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';

// Custom XML
export const DS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';
export const CUST_DATA_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties';

// Ink
export const INK = 'http://schemas.microsoft.com/ink/2010/main';

// Office Web Extensions
export const WE = 'http://schemas.microsoft.com/office/webextensions/webextension/2010/11';
export const WE_TP = 'http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11';

/**
 * Namespace prefix map for building QNames
 */
export const NAMESPACE_PREFIXES: Record<string, string> = {
  [W]: 'w',
  [W14]: 'w14',
  [W15]: 'w15',
  [S]: 's',
  [P]: 'p',
  [A]: 'a',
  [R]: 'r',
  [MC]: 'mc',
  [M]: 'm',
  [V]: 'v',
  [O]: 'o',
  [WP]: 'wp',
  [WP14]: 'wp14',
  [WPG]: 'wpg',
  [WPS]: 'wps',
  [PIC]: 'pic',
  [C]: 'c',
  [DC]: 'dc',
  [DCTERMS]: 'dcterms',
  [CP]: 'cp',
  [EP]: 'ep',
};

/**
 * Build a qualified name with namespace
 */
export function qname(ns: string, localName: string): string {
  return `{${ns}}${localName}`;
}

/**
 * Parse a qualified name into namespace and local name
 */
export function parseQName(qname: string): { ns: string; localName: string } | null {
  const match = qname.match(/^\{([^}]+)\}(.+)$/);
  if (!match) return null;
  return { ns: match[1], localName: match[2] };
}
