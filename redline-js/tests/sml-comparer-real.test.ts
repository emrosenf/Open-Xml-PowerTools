import { describe, it, expect } from 'vitest';
import { compare, buildChangeList } from '../src/sml/sml-comparer';
import { SmlChangeType } from '../src/sml/types';
import { generateExcelFile } from './sml-test-data';

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
      c.changeType === SmlChangeType.ValueChanged
    );
    expect(valueChanges.length).toBeGreaterThan(0);
    expect(valueChanges.some(c => c.cellAddress?.includes('A1'))).toBe(true);
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
      c.changeType === SmlChangeType.FormulaChanged
    );
    expect(formulaChanges.length).toBeGreaterThan(0);
    expect(formulaChanges.some(c => c.cellAddress?.includes('B1'))).toBe(true);
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
            { address: 'A1', value: '100' },
            { address: 'B1', value: '200' },
          ]},
        ],
      }],
    });

    const result = await compare(older, newer, defaultSettings);

    const added = result.changes.filter(c =>
      c.changeType === SmlChangeType.CellAdded
    );
    expect(added.length).toBeGreaterThan(0);
    expect(added.some(c => c.cellAddress?.includes('B1'))).toBe(true);
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
          { index: 1, cells: [{ address: 'A1', value: '100' }] },
        ],
      }],
    });

    const result = await compare(older, newer, defaultSettings);

    const deleted = result.changes.filter(c =>
      c.changeType === SmlChangeType.CellDeleted
    );
    expect(deleted.length).toBeGreaterThan(0);
    expect(deleted.some(c => c.cellAddress?.includes('B1'))).toBe(true);
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
          { index: 1, cells: [{ address: 'A1', value: '100' }] },
          { index: 2, cells: [{ address: 'A2', value: '200' }] },
        ],
      }],
    });

    const result = await compare(older, newer, defaultSettings);

    const changesForNewRow = result.changes.filter(c =>
      c.changeType === SmlChangeType.RowInserted ||
      c.changeType === SmlChangeType.CellAdded
    );
    expect(changesForNewRow.length).toBeGreaterThan(0);
  });

  it('should detect row deletions', async () => {
    const older = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [{ address: 'A1', value: '100' }] },
          { index: 2, cells: [{ address: 'A2', value: '200' }] },
        ],
      }],
    });

    const newer = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [{ address: 'A1', value: '100' }] },
        ],
      }],
    });

    const result = await compare(older, newer, defaultSettings);

    const changesForDeletedRow = result.changes.filter(c =>
      c.changeType === SmlChangeType.RowDeleted ||
      c.changeType === SmlChangeType.CellDeleted
    );
    expect(changesForDeletedRow.length).toBeGreaterThan(0);
  });

  it('should detect added sheets', async () => {
    const older = await generateExcelFile({
      sheets: [{ name: 'Sheet1', rows: [] }],
    });

    const newer = await generateExcelFile({
      sheets: [
        { name: 'Sheet1', rows: [] },
        { name: 'Sheet2', rows: [] },
      ],
    });

    const result = await compare(older, newer, defaultSettings);

    const added = result.changes.filter(c =>
      c.changeType === SmlChangeType.SheetAdded
    );
    expect(added.length).toBe(1);
    expect(added.some(c => c.cellAddress === 'Sheet2')).toBe(true);
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
      c.changeType === SmlChangeType.SheetDeleted
    );
    expect(deleted.length).toBe(1);
    expect(deleted.some(c => c.cellAddress === 'Sheet2')).toBe(true);
  });

  it('should detect renamed sheets', async () => {
    const older = await generateExcelFile({
      sheets: [{ name: 'Sheet1', rows: [
        { index: 1, cells: [{ address: 'A1', value: 'data' }] },
      ]}],
    });

    const newer = await generateExcelFile({
      sheets: [{ name: 'RenamedSheet', rows: [
        { index: 1, cells: [{ address: 'A1', value: 'data' }] },
      ]}],
    });

    const result = await compare(older, newer, defaultSettings);

    const renamed = result.changes.filter(c =>
      c.changeType === SmlChangeType.SheetRenamed
    );
    expect(renamed.length).toBe(1);
    expect(renamed[0].oldSheetName).toBe('Sheet1');
    expect(renamed[0].cellAddress).toBe('RenamedSheet');
  });

  it('should group adjacent change list items', async () => {
    const older = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [{ address: 'C1', value: '1' }] },
          { index: 2, cells: [{ address: 'C2', value: '2' }] },
        ],
      }],
    });

    const newer = await generateExcelFile({
      sheets: [{
        name: 'Sheet1',
        rows: [
          { index: 1, cells: [{ address: 'C1', value: '10' }] },
          { index: 2, cells: [{ address: 'C2', value: '20' }] },
        ],
      }],
    });

    const result = await compare(older, newer, defaultSettings);
    const list = buildChangeList(result);

    const valueItems = list.filter(item => item.changeType === SmlChangeType.ValueChanged);
    expect(valueItems.length).toBe(1);
    expect(valueItems[0].cellRange).toBe('C1:C2');
    expect(valueItems[0].anchor).toBe('Sheet1!C1:C2');
    expect(valueItems[0].sheetName).toBe('Sheet1');
  });
});
