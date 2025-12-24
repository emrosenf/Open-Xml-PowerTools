/**
 * @docredline/core - Document comparison library
 *
 * TypeScript port of Open-Xml-PowerTools comparers for Word, Excel, and PowerPoint.
 *
 * @packageDocumentation
 */

// Core utilities
export * from './types';
export * from './core';

// Word document handling
export * from './wml/document';
export * from './wml/revision';

// Word document comparison
export * from './wml/wml-comparer';

// Excel document comparison (to be implemented)
// export * from './sml/sml-comparer';

// PowerPoint document comparison (to be implemented)
// export * from './pml/pml-comparer';
