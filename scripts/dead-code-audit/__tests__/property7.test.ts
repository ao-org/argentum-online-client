/**
 * Property 7: Duplicate code detection
 *
 * For any two code blocks of 10 or more consecutive executable lines that are
 * textually identical after whitespace normalization, the pair must be reported
 * as a duplicate. Blocks shorter than 10 lines must not be reported.
 *
 * Validates: Requirements 5.1, 5.2
 */
import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { detectDuplicates } from '../src/duplicateDetector.js';
import type { ParsedModule, ParsedLine, SourceFile } from '../src/types.js';

// --- Helpers ---

function makeModule(path: string, moduleName: string, lines: string[]): ParsedModule {
  const source: SourceFile = { path, type: 'bas', moduleName, content: lines.join('\n') };
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

// --- Arbitraries ---

/**
 * Generate a single VB6-like executable line.
 * Uses realistic VB6 patterns to avoid empty/whitespace-only lines.
 */
const vb6LineArb: fc.Arbitrary<string> = fc.oneof(
  fc.integer({ min: 0, max: 999 }).map(n => `Dim var${n} As Long`),
  fc.integer({ min: 0, max: 999 }).map(n => `var${n} = var${n} + 1`),
  fc.integer({ min: 0, max: 999 }).map(n => `Call DoSomething${n}(var${n})`),
  fc.integer({ min: 0, max: 999 }).map(n => `If var${n} > 0 Then`),
  fc.constant('End If'),
  fc.integer({ min: 0, max: 999 }).map(n => `Debug.Print var${n}`),
);

/**
 * Generate a block of N unique VB6-like lines.
 */
function codeBlockArb(size: number): fc.Arbitrary<string[]> {
  return fc.array(vb6LineArb, { minLength: size, maxLength: size });
}


/**
 * Generate a block of N lines that are distinct from a given block
 * (used as filler/padding around duplicates).
 */
function fillerBlockArb(size: number): fc.Arbitrary<string[]> {
  return fc.array(
    fc.integer({ min: 1000, max: 9999 }).map(n => `fillerVar${n} = ${n}`),
    { minLength: size, maxLength: size },
  );
}

// --- Property Tests ---

describe('Feature: dead-code-audit, Property 7: Duplicate code detection', () => {
  it('blocks of ?10 identical lines across two modules are detected as exact duplicates', () => {
    /**
     * Validates: Requirements 5.1, 5.2
     *
     * Strategy:
     * 1. Generate a code block of size N (10–20 lines)
     * 2. Insert the identical block into two separate modules (with filler padding)
     * 3. Call detectDuplicates with minLines=10
     * 4. Verify at least one exact duplicate pair is found covering the inserted block
     */
    fc.assert(
      fc.property(
        fc.integer({ min: 10, max: 20 }).chain((blockSize) =>
          fc.tuple(
            codeBlockArb(blockSize),
            fillerBlockArb(5),
            fillerBlockArb(5),
          ).map(([dupBlock, fillerA, fillerB]) => ({
            blockSize,
            dupBlock,
            fillerA,
            fillerB,
          })),
        ),
        ({ blockSize, dupBlock, fillerA, fillerB }) => {
          // Module A: filler + duplicate block
          const linesA = [...fillerA, ...dupBlock];
          const modA = makeModule('CODIGO/ModA.bas', 'ModA', linesA);

          // Module B: filler + duplicate block
          const linesB = [...fillerB, ...dupBlock];
          const modB = makeModule('CODIGO/ModB.bas', 'ModB', linesB);

          const results = detectDuplicates([modA, modB], 10);

          // There must be at least one exact duplicate pair found
          const exactPairs = results.filter(p => p.type === 'exact');
          expect(exactPairs.length).toBeGreaterThanOrEqual(1);

          // At least one pair must cover the inserted duplicate block
          const found = exactPairs.some(p =>
            p.lineCount >= 10 &&
            ((p.fileA === 'CODIGO/ModA.bas' && p.fileB === 'CODIGO/ModB.bas') ||
             (p.fileA === 'CODIGO/ModB.bas' && p.fileB === 'CODIGO/ModA.bas'))
          );
          expect(found).toBe(true);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('blocks shorter than 10 lines are NOT reported as duplicates', () => {
    /**
     * Validates: Requirements 5.1, 5.2
     *
     * Strategy:
     * 1. Generate a code block of size N (1–9 lines)
     * 2. Insert the identical block into two separate modules with enough filler
     *    so each module has ?10 total lines (to ensure modules are processed)
     * 3. Call detectDuplicates with minLines=10
     * 4. Verify no duplicate pairs are reported
     */
    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 9 }).chain((blockSize) =>
          fc.tuple(
            codeBlockArb(blockSize),
            // Ensure each module has enough total lines to be processed
            fillerBlockArb(Math.max(10, 15 - blockSize)),
            fillerBlockArb(Math.max(10, 15 - blockSize)),
          ).map(([dupBlock, fillerA, fillerB]) => ({
            blockSize,
            dupBlock,
            fillerA,
            fillerB,
          })),
        ),
        ({ dupBlock, fillerA, fillerB }) => {
          // Module A: filler + short duplicate block
          const linesA = [...fillerA, ...dupBlock];
          const modA = makeModule('CODIGO/ModA.bas', 'ModA', linesA);

          // Module B: different filler + same short duplicate block
          const linesB = [...fillerB, ...dupBlock];
          const modB = makeModule('CODIGO/ModB.bas', 'ModB', linesB);

          const results = detectDuplicates([modA, modB], 10);

          // No pairs should be reported since the shared block is < 10 lines
          // and the filler blocks are unique to each module
          expect(results.length).toBe(0);
        },
      ),
      { numRuns: 100 },
    );
  });
});
