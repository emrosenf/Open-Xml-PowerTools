#!/usr/bin/env bun
/**
 * Document comparison CLI
 *
 * Compares two Office documents (Word, Excel, or PowerPoint) and outputs
 * a marked-up document showing differences.
 *
 * Usage: bun bin/compare.ts <file1> <file2> [output]
 *
 * Examples:
 *   bun bin/compare.ts old.docx new.docx              # outputs comparison-result.docx
 *   bun bin/compare.ts old.xlsx new.xlsx result.xlsx  # outputs result.xlsx
 *   bun bin/compare.ts v1.pptx v2.pptx changes.pptx   # outputs changes.pptx
 */

import { readFile, writeFile } from 'node:fs/promises';
import { basename, extname, dirname, join } from 'node:path';

import { compareDocuments } from '../src/wml/wml-comparer';
import { produceMarkedWorkbook } from '../src/sml/sml-comparer';
import { produceMarkedPresentation } from '../src/pml/pml-comparer';

type FileType = 'docx' | 'xlsx' | 'pptx';

interface CompareOptions {
  author?: string;
  verbose?: boolean;
}

function usage(): never {
  console.error(`
Usage: bun bin/compare.ts [options] <file1> <file2> [output]

Compare two Office documents and produce a marked-up result showing differences.

Arguments:
  file1   Original/older document (.docx, .xlsx, or .pptx)
  file2   Revised/newer document (must be same type as file1)
  output  Output file path (optional, defaults to comparison-result.<ext>)

Options:
  --author <name>   Author name for tracked changes (Word only)
  --verbose, -v     Show detailed progress
  --help, -h        Show this help message

Examples:
  bun bin/compare.ts old.docx new.docx
  bun bin/compare.ts old.docx new.docx result.docx
  bun bin/compare.ts --author "Reviewer" v1.docx v2.docx
  bun bin/compare.ts spreadsheet-v1.xlsx spreadsheet-v2.xlsx
  bun bin/compare.ts presentation-old.pptx presentation-new.pptx
`);
  process.exit(1);
}

function detectFileType(filePath: string): FileType | null {
  const ext = extname(filePath).toLowerCase();
  switch (ext) {
    case '.docx':
      return 'docx';
    case '.xlsx':
      return 'xlsx';
    case '.pptx':
      return 'pptx';
    default:
      return null;
  }
}

function getDefaultOutputName(file1: string, fileType: FileType): string {
  const dir = dirname(file1);
  return join(dir, `comparison-result.${fileType}`);
}

async function compareWord(
  file1: Buffer,
  file2: Buffer,
  options: CompareOptions
): Promise<Buffer> {
  if (options.verbose) {
    console.error('Comparing Word documents...');
  }

  const result = await compareDocuments(file1, file2, {
    author: options.author ?? 'Document Comparer',
  });

  if (options.verbose) {
    console.error(`Found ${result.revisionCount} revisions:`);
    console.error(`  - ${result.insertions} insertions`);
    console.error(`  - ${result.deletions} deletions`);
  }

  return result.document;
}

async function compareExcel(
  file1: Buffer,
  file2: Buffer,
  options: CompareOptions
): Promise<Buffer> {
  if (options.verbose) {
    console.error('Comparing Excel spreadsheets...');
  }

  const result = await produceMarkedWorkbook(file1, file2, {});

  if (options.verbose) {
    console.error('Comparison complete. Changes highlighted in output.');
  }

  return result;
}

async function comparePowerPoint(
  file1: Buffer,
  file2: Buffer,
  options: CompareOptions
): Promise<Buffer> {
  if (options.verbose) {
    console.error('Comparing PowerPoint presentations...');
  }

  const result = await produceMarkedPresentation(file1, file2, {});

  if (options.verbose) {
    console.error('Comparison complete. Changes marked in output.');
  }

  return result;
}

async function main() {
  const args = process.argv.slice(2);

  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    usage();
  }

  const options: CompareOptions = {};
  const positionalArgs: string[] = [];

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];

    if (arg === '--author') {
      options.author = args[++i];
      if (!options.author) {
        console.error('Error: --author requires a value');
        process.exit(1);
      }
    } else if (arg === '--verbose' || arg === '-v') {
      options.verbose = true;
    } else if (arg.startsWith('-')) {
      console.error(`Error: Unknown option: ${arg}`);
      usage();
    } else {
      positionalArgs.push(arg);
    }
  }

  if (positionalArgs.length < 2) {
    console.error('Error: Two input files are required');
    usage();
  }

  const [file1Path, file2Path, outputPath] = positionalArgs;

  const type1 = detectFileType(file1Path);
  const type2 = detectFileType(file2Path);

  if (!type1) {
    console.error(`Error: Unsupported file type: ${file1Path}`);
    console.error('Supported types: .docx, .xlsx, .pptx');
    process.exit(1);
  }

  if (!type2) {
    console.error(`Error: Unsupported file type: ${file2Path}`);
    console.error('Supported types: .docx, .xlsx, .pptx');
    process.exit(1);
  }

  if (type1 !== type2) {
    console.error(`Error: File types must match`);
    console.error(`  File 1: ${type1} (${basename(file1Path)})`);
    console.error(`  File 2: ${type2} (${basename(file2Path)})`);
    process.exit(1);
  }

  const fileType = type1;
  const finalOutputPath = outputPath ?? getDefaultOutputName(file1Path, fileType);

  if (options.verbose) {
    console.error(`Reading ${basename(file1Path)}...`);
  }
  let file1: Buffer;
  try {
    file1 = await readFile(file1Path);
  } catch (err) {
    console.error(`Error: Cannot read file: ${file1Path}`);
    console.error((err as Error).message);
    process.exit(1);
  }

  if (options.verbose) {
    console.error(`Reading ${basename(file2Path)}...`);
  }
  let file2: Buffer;
  try {
    file2 = await readFile(file2Path);
  } catch (err) {
    console.error(`Error: Cannot read file: ${file2Path}`);
    console.error((err as Error).message);
    process.exit(1);
  }

  let result: Buffer;
  try {
    switch (fileType) {
      case 'docx':
        result = await compareWord(file1, file2, options);
        break;
      case 'xlsx':
        result = await compareExcel(file1, file2, options);
        break;
      case 'pptx':
        result = await comparePowerPoint(file1, file2, options);
        break;
    }
  } catch (err) {
    console.error('Error during comparison:');
    console.error((err as Error).message);
    if (options.verbose) {
      console.error((err as Error).stack);
    }
    process.exit(1);
  }

  if (options.verbose) {
    console.error(`Writing ${basename(finalOutputPath)}...`);
  }
  try {
    await writeFile(finalOutputPath, result);
  } catch (err) {
    console.error(`Error: Cannot write file: ${finalOutputPath}`);
    console.error((err as Error).message);
    process.exit(1);
  }

  console.log(finalOutputPath);
}

main().catch((err) => {
  console.error('Unexpected error:', err);
  process.exit(1);
});
