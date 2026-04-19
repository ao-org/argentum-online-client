/**
 * Property 8: Near-duplicate structural detection
 *
 * For any two code blocks that differ only in identifier names but share
 * identical structure (same operators, keywords, control flow, and literal
 * values), the pair must be reported as "near-duplicate".
 *
 * Validates: Requirements 5.4
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
 * Generate a pair of identifier names (original + renamed variant).
 * Both are valid VB6 identifiers that are NOT VB6 keywords.
 */
const identPairArb: fc.Arbitrary<{ original: string; renamed: string }> = fc
  .integer({ min: 0, max: 499 })
  .map(n => ({
    original: `origVar${n}`,
    renamed: `renamedVar${n}`,
  }));

/**
 * Template-based code block generator.
 * Produces a block of ?10 lines using VB6 patterns with placeholder identifiers.
 * Returns a function that, given an identifier mapping, produces the final lines.
 */
function templateBlockArb(size: number): fc.Arbitrary<{
  makeBlock: (ids: { v1: string; v2: string; fn: string }) => string[];
}> {
  // Generate a fixed template of VB6-like lines using 3 identifiers
  return fc.integer({ min: 100, max: 999 }).map(seed => ({
    makeBlock: (ids: { v1: string; v2: string; fn: string }) => {
      const lines: string[] = [];
      lines.push(`Dim ${ids.v1} As Long`);
      lines.push(`Dim ${ids.v2} As Long`);
      lines.push(`${ids.v1} = ${seed}`);
      lines.push(`${ids.v2} = ${ids.v1} + 1`);
      lines.push(`Call ${ids.fn}(${ids.v1}, ${ids.v2})`);
      lines.push(`If ${ids.v1} > 0 Then`);
      lines.push(`  ${ids.v2} = ${ids.v2} * 2`);
      lines.push(`  Debug.Print ${ids.v1}`);
      lines.push(`End If`);
      lines.push(`${ids.v1} = ${ids.v2} + ${seed}`);
      // Add extra lines if size > 10
      for (let i = 10; i < size; i++) {
        lines.push(`${ids.v2} = ${ids.v1} + ${i}`);
      }
      return lines;
    },
  }));
}


/**
 * Generate filler lines that are unique and won't match any template block.
 */
function fillerBlockArb(size: number): fc.Arbitrary<string[]> {
  return fc.array(
    fc.integer({ min: 5000, max: 9999 }).map(n => `uniqueFiller${n} = ${n}`),
    { minLength: size, maxLength: size },
  );
}

// --- Property Tests ---

describe('Feature: dead-code-audit, Property 8: Near-duplicate structural detection', () => {
  it('code blocks differing only in identifier names are reported as near-duplicate', () => {
    /**
     * Validates: Requirements 5.4
     *
     * Strategy:
     * 1. Generate a template code block of ?10 lines
     * 2. Instantiate it twice with different identifier names but same structure
     * 3. Place each version in a separate module with unique filler padding
     * 4. Call detectDuplicates
     * 5. Verify at least one near-duplicate pair is found
     */
    fc.assert(
      fc.property(
        fc.integer({ min: 10, max: 15 }).chain((blockSize) =>
          fc.tuple(
            templateBlockArb(blockSize),
            fillerBlockArb(5),
            fillerBlockArb(5),
          ).map(([template, fillerA, fillerB]) => ({
            blockSize,
            template,
            fillerA,
            fillerB,
          })),
        ),
        ({ template, fillerA, fillerB }) => {
          // Version A: original identifiers
          const blockA = template.makeBlock({
            v1: 'alphaCount',
            v2: 'alphaTotal',
            fn: 'ProcessAlpha',
          });

          // Version B: different identifiers, same structure
          const blockB = template.makeBlock({
            v1: 'betaCount',
            v2: 'betaTotal',
            fn: 'ProcessBeta',
          });

          // Sanity: blocks should NOT be textually identical
          const isDifferent = blockA.some((line, i) => line !== blockB[i]);
          expect(isDifferent).toBe(true);

          const linesA = [...fillerA, ...blockA];
          const linesB = [...fillerB, ...blockB];

          const modA = makeModule('CODIGO/ModA.bas', 'ModA', linesA);
          const modB = makeModule('CODIGO/ModB.bas', 'ModB', linesB);

          const results = detectDuplicates([modA, modB], 10);

          // At least one pair should be found (near-duplicate or exact)
          const nearDups = results.filter(p => p.type === 'near-duplicate');
          expect(nearDups.length).toBeGreaterThanOrEqual(1);

          // The near-duplicate pair should span across the two modules
          const found = nearDups.some(p =>
            ((p.fileA === 'CODIGO/ModA.bas' && p.fileB === 'CODIGO/ModB.bas') ||
             (p.fileA === 'CODIGO/ModB.bas' && p.fileB === 'CODIGO/ModA.bas'))
          );
          expect(found).toBe(true);
        },
      ),
      { numRuns: 100 },
    );
  });
});
