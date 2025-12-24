/**
 * Debug test to understand fast-xml-parser output structure
 */

import { describe, it, expect } from 'vitest';
import { XMLParser } from 'fast-xml-parser';

describe('Debug XML Parser', () => {
  it('shows preserveOrder structure', () => {
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      preserveOrder: true,
    });

    const xml = '<root attr="value"><child>text</child></root>';
    const result = parser.parse(xml);

    console.log('Parsed result:', JSON.stringify(result, null, 2));

    // With preserveOrder, structure is different
    expect(result).toBeDefined();
  });

  it('shows content types structure', () => {
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      preserveOrder: true,
    });

    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

    const result = parser.parse(xml);
    console.log('Content Types:', JSON.stringify(result, null, 2));

    expect(result).toBeDefined();
  });
});
