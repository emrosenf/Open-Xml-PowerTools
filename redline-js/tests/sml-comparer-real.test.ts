import { describe, it, expect } from 'vitest';
import { compare } from '../src/sml/sml-comparer';
import { generateExcelFile, type ExcelFileSpec, type SheetSpec, type RowSpec, type CellSpec } from './sml-test-data';

describe('SmlComparer - Real Excel Files', () => {
  const defaultSettings = {
    compareValues: true,
    compareFormulas: true,
    compareFormatting: false,
    compareComments: true,
    compareDataValidations: true,
    compareMergedCells: true,
    compareHyperlinks: true,
    enableRowAlignment: true,
    enableColumnAlignment: false,
  };

  it('should detect cell value changes', async () => {
    const older = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [{ address: 'A1', value: '100' }] },
        ],
      }],
    });

    const newer = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [{ address: 'A1', value: '150' }] },
        ],
      }],
    });

    const result = await compare(older, newer, defaultSettings);

    const valueChanges = result.changes.filter(c =>
      c.changeType === 8 // ValueChanged
    );
    expect(valueChanges.length).toBeGreaterThan(0);
    expect(valueChanges.some(c => c.cellAddress === 'A1')).toBe(true);
  });

  it('should detect cell formula changes', async () => {
    const older = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [
            { address: 'A1', value: '10' },
            { address: 'B1', formula: 'A1*2' },
          ]},
        ],
      }],
    });

    const newer = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [
            { address: 'A1', value: '10' },
            { address: 'B1', formula: 'A1*5' },
          ]},
        ],
      }],
    });

    const result = await compare(older, newer, defaultSettings);

    const formulaChanges = result.changes.filter(c =>
      c.changeType === 12 // FormulaChanged
    );
    expect(formulaChanges.length).toBeGreaterThan(0);
    expect(formulaChanges.some(c => c.cellAddress === 'A1')).toBe(true);
  });

  it('should detect cell additions', async () => {
    const older = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [{ address: 'A1', value: '100' }] },
        ],
      }],
    });

    const newer = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [
            { address: 'B1', value: '200' }],
          ],
        ],
      }],
    });

    const result = await compare(older, newer, defaultSettings);

    const added = result.changes.filter(c =>
      c.changeType === 5 // CellAdded
    );
    expect(added.length).toBeGreaterThan(0);
    expect(added.some(c => c.cellAddress === 'B1')).toBe(true);
  });

  it('should detect cell deletions', async () => {
    const older = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [
            { address: 'A1', value: '100' },
            { address: 'B1', value: '200' },
          ]},
        ],
      }],
    });

    const newer = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [
            { address: 'A1', value: '100' },
          ],
        ],
      }],
    });

    const result = await compare(older, newer, defaultSettings);

    const deleted = result.changes.filter(c =>
      c.changeType === 6 // CellDeleted
    );
    expect(deleted.length).toBeGreaterThan(0);
    expect(deleted.some(c => c.cellAddress === 'B1')).toBe(true);
  });

  it('should detect row insertions', async () => {
    const older = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [{ address: 'A1', value: '100' }] },
        ],
      }],
    });

    const newer = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [
            { address: 'A1', value: '100' },
          ],
          { index: 2, cells: [{ address: 'A2', value: '200' }] },
          ],
        ],
      }],
    });

    const result = await compare(older, newer, defaultSettings);

    const inserted = result.changes.filter(c =>
      c.changeType === 4 // RowInserted
    );
    expect(inserted.length).toBeGreaterThan(0);
    expect(inserted.some(c => c.rowIndex === 2)).toBe(true);
  });

  it('should detect row deletions', async () => {
    const older = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [{ address: 'A1', value: '100' }],
          { index: 2, cells: [{ address: 'A2', value: '200' }],
          ],
        ],
      }],
    });

    const newer = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [
            { address: 'A1', value: '100' },
          ],
        ],
      }],
    });

    const result = await compare(older, newer, defaultSettings);

    const deleted = result.changes.filter(c =>
      c.changeType === 5 // RowDeleted
    );
    expect(deleted.length).toBeGreaterThan(0);
    expect(deleted.some(c => c.rowIndex === 2)).toBe(true);
  });

  it('should detect added sheets', async () => {
    const older = await generateExcelFile({
      sheets: [{ name: 'Sheet1', rows: [] }],
    });

    const newer = await generateExcelFile({
      sheets: [
        { name: 'Sheet2', rows: [] },
      ],
    });

    const result = await compare(older, newer, defaultSettings);

    const added = result.changes.filter(c =>
      c.changeType === 0 // SheetAdded
    );
    expect(added.length).toBe(1);
    expect(added.some(c => c.sheetName === 'Sheet2')).toBe(true);
  });

  it('should detect deleted sheets', async () => {
    const older = await generateExcelFile({
      sheets: [
        { name: 'Sheet1', rows: [] },
        { name: 'Sheet2', rows: [] },
      ],
    });

    const newer = await generateExcelFile({
      sheets: [{ name: 'Sheet1', rows: [] }],
    });

    const result = await compare(older, newer, defaultSettings);

    const deleted = result.changes.filter(c =>
      c.changeType === 1 // SheetDeleted
    );
    expect(deleted.length).toBe(1);
    expect(deleted.some(c => c.sheetName === 'Sheet2')).toBe(true);
  });

  it('should detect renamed sheets', async () => {
    const older = await generateExcelFile({
      sheets: [{ name: 'Sheet1', rows: [] }],
    });

    const newer = await generateExcelFile({
      sheets: [
        { name: 'Sheet2', rows: [] },
      ],
    });

    const result = await compare(older, newer, defaultSettings);

    const renamed = result.changes.filter(c =>
      c.changeType === 2 // SheetRenamed
    );
    expect(renamed.length).toBe(1);
    expect(renamed.some(c => c.oldSheetName === 'Sheet1' && c.newSheetName === 'Sheet2')).toBe(true);
  });
});
