import { describe, it, expect } from 'vitest';
import { detectDuplicates } from './duplicateDetector.js';
import type { ParsedModule, SourceFile, ParsedLine } from './types.js';

/** Helper to create a ParsedModule from an array of executable line strings. */
function makeModule(
  path: string,
  moduleName: string,
  lines: string[],
): ParsedModule {
  const source: SourceFile = {
    path,
    type: 'bas',
    moduleName,
    content: lines.join('\n'),
  };

  const parsedLines: ParsedLine[] = lines.map((text, i) => ({
    lineNumber: i + 1,
    text,
    isComment: false,
    isPreprocessor: false,
    isExecutable: true,
    originalLines: [i + 1],
  }));

  return { source, lines: parsedLines, attributeLines: [] };
}

/** Generate N distinct executable lines. */
function generateLines(prefix: string, count: number): string[] {
  return Array.from({ length: count }, (_, i) => `${prefix} = ${prefix} + ${i + 1}`);
}

describe('detectDuplicates', () => {
  it('should detect exact duplicates of minLines consecutive lines', () => {
    const sharedBlock = generateLines('x', 10);
    const modA = makeModule('fileA.bas', 'ModA', [
      'Dim x As Long',
      ...sharedBlock,
      'x = 0',
    ]);
    const modB = makeModule('fileB.bas', 'ModB', [
      'Dim y As Long',
      ...sharedBlock,
      'y = 0',
    ]);

    const results = detectDuplicates([modA, modB], 10);
    expect(results.length).toBe(1);
    expect(results[0].type).toBe('exact');
    expect(results[0].lineCount).toBe(10);
    expect(results[0].fileA).toBe('fileA.bas');
    expect(results[0].fileB).toBe('fileB.bas');
  });

  it('should not report blocks shorter than minLines', () => {
    const sharedBlock = generateLines('x', 9);
    const modA = makeModule('fileA.bas', 'ModA', [
      'Dim x As Long',
      ...sharedBlock,
    ]);
    const modB = makeModule('fileB.bas', 'ModB', [
      'Dim y As Long',
      ...sharedBlock,
    ]);

    const results = detectDuplicates([modA, modB], 10);
    expect(results.length).toBe(0);
  });

  it('should detect exact duplicates within the same file (non-overlapping)', () => {
    const block = generateLines('val', 10);
    const modA = makeModule('fileA.bas', 'ModA', [
      ...block,
      'separator = 1',
      'separator = 2',
      'separator = 3',
      ...block,
    ]);

    const results = detectDuplicates([modA], 10);
    expect(results.length).toBe(1);
    expect(results[0].type).toBe('exact');
    expect(results[0].fileA).toBe('fileA.bas');
    expect(results[0].fileB).toBe('fileA.bas');
  });

  it('should not report overlapping blocks in the same file', () => {
    // A single block of 11 lines — two windows of 10 overlap heavily
    const block = generateLines('z', 11);
    const modA = makeModule('fileA.bas', 'ModA', block);

    const results = detectDuplicates([modA], 10);
    expect(results.length).toBe(0);
  });

  it('should detect near-duplicates differing only in identifier names', () => {
    const blockA = Array.from({ length: 10 }, (_, i) =>
      `alpha = alpha + beta * ${i + 1}`,
    );
    const blockB = Array.from({ length: 10 }, (_, i) =>
      `gamma = gamma + delta * ${i + 1}`,
    );

    const modA = makeModule('fileA.bas', 'ModA', blockA);
    const modB = makeModule('fileB.bas', 'ModB', blockB);

    const results = detectDuplicates([modA, modB], 10);
    expect(results.length).toBe(1);
    expect(results[0].type).toBe('near-duplicate');
  });

  it('should normalize whitespace before comparison', () => {
    const blockA = Array.from({ length: 10 }, (_, i) =>
      `x  =  x  +  ${i + 1}`,
    );
    const blockB = Array.from({ length: 10 }, (_, i) =>
      `x = x + ${i + 1}`,
    );

    const modA = makeModule('fileA.bas', 'ModA', blockA);
    const modB = makeModule('fileB.bas', 'ModB', blockB);

    const results = detectDuplicates([modA, modB], 10);
    expect(results.length).toBe(1);
    expect(results[0].type).toBe('exact');
  });

  it('should skip comment and non-executable lines', () => {
    const source: SourceFile = {
      path: 'fileA.bas',
      type: 'bas',
      moduleName: 'ModA',
      content: '',
    };

    const lines: ParsedLine[] = [
      // 5 comment lines then 5 executable — not enough executable for minLines=10
      ...Array.from({ length: 5 }, (_, i) => ({
        lineNumber: i + 1,
        text: `' comment ${i}`,
        isComment: true,
        isPreprocessor: false,
        isExecutable: false,
        originalLines: [i + 1],
      })),
      ...Array.from({ length: 5 }, (_, i) => ({
        lineNumber: i + 6,
        text: `x = x + ${i}`,
        isComment: false,
        isPreprocessor: false,
        isExecutable: true,
        originalLines: [i + 6],
      })),
    ];

    const modA: ParsedModule = { source, lines, attributeLines: [] };
    const modB = makeModule('fileB.bas', 'ModB',
      Array.from({ length: 5 }, (_, i) => `x = x + ${i}`),
    );

    const results = detectDuplicates([modA, modB], 10);
    expect(results.length).toBe(0);
  });

  it('should not report the same pair twice', () => {
    const block = generateLines('v', 10);
    const modA = makeModule('fileA.bas', 'ModA', block);
    const modB = makeModule('fileB.bas', 'ModB', block);
    const modC = makeModule('fileC.bas', 'ModC', block);

    const results = detectDuplicates([modA, modB, modC], 10);
    // Should find 3 pairs: A-B, A-C, B-C
    expect(results.length).toBe(3);
    const keys = results.map(r =>
      [r.fileA, r.fileB].sort().join('|'),
    );
    const uniqueKeys = new Set(keys);
    expect(uniqueKeys.size).toBe(3);
  });

  it('should report correct line numbers', () => {
    const padding = ['a = 1', 'b = 2', 'c = 3'];
    const block = generateLines('x', 10);
    const modA = makeModule('fileA.bas', 'ModA', [...padding, ...block]);
    const modB = makeModule('fileB.bas', 'ModB', block);

    const results = detectDuplicates([modA, modB], 10);
    expect(results.length).toBe(1);
    // In modA, block starts at line 4 (after 3 padding lines)
    expect(results[0].startLineA).toBe(4);
    expect(results[0].endLineA).toBe(13);
    // In modB, block starts at line 1
    expect(results[0].startLineB).toBe(1);
    expect(results[0].endLineB).toBe(10);
  });

  it('should use default minLines of 10', () => {
    const block = generateLines('x', 10);
    const modA = makeModule('fileA.bas', 'ModA', block);
    const modB = makeModule('fileB.bas', 'ModB', block);

    // Call without minLines argument
    const results = detectDuplicates([modA, modB]);
    expect(results.length).toBe(1);
  });

  it('should not flag structurally different blocks as near-duplicates', () => {
    const blockA = Array.from({ length: 10 }, (_, i) =>
      `x = x + ${i + 1}`,
    );
    const blockB = Array.from({ length: 10 }, (_, i) =>
      `If x > ${i + 1} Then y = y - 1`,
    );

    const modA = makeModule('fileA.bas', 'ModA', blockA);
    const modB = makeModule('fileB.bas', 'ModB', blockB);

    const results = detectDuplicates([modA, modB], 10);
    expect(results.length).toBe(0);
  });

  it('should handle empty modules gracefully', () => {
    const modA = makeModule('fileA.bas', 'ModA', []);
    const modB = makeModule('fileB.bas', 'ModB', ['x = 1']);

    const results = detectDuplicates([modA, modB], 10);
    expect(results.length).toBe(0);
  });
});
