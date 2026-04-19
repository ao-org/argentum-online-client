/**
 * Property 9: Report summary count accuracy
 *
 * For any set of findings, the summary section's per-category counts must
 * exactly equal the number of findings in each category. Client and server
 * counts must be tallied independently.
 *
 * Validates: Requirements 6.2
 */
import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { computeSummary } from '../src/reportGenerator.js';
import type {
  Finding,
  FindingCategory,
  Confidence,
  DuplicatePair,
} from '../src/types.js';

// --- Arbitraries ---

const confidenceArb: fc.Arbitrary<Confidence> = fc.constantFrom('confirmed', 'review-needed');

const filePathArb = fc.stringMatching(/^CODIGO\/[A-Z][a-zA-Z0-9]{1,8}\.(bas|cls|frm)$/)
  .filter(s => s.length > 8);

const symbolNameArb = fc.stringMatching(/^[a-zA-Z][a-zA-Z0-9_]{1,12}$/)
  .filter(s => s.length >= 2);

const lineNumberArb = fc.integer({ min: 1, max: 5000 });

const reasonArb = fc.stringMatching(/^[A-Za-z ]{5,30}$/).filter(s => s.length >= 5);

const allCategories: FindingCategory[] = [
  'unused-procedure',
  'unused-variable',
  'write-only-variable',
  'unused-const',
  'unused-enum',
  'unused-type',
  'unused-declare',
  'unreachable-code',
  'dead-branch',
  'commented-out-block',
];

/**
 * Generate a random Finding with a given category.
 */
function findingOfCategoryArb(category: FindingCategory): fc.Arbitrary<Finding> {
  return fc.record({
    confidence: confidenceArb,
    filePath: filePathArb,
    startLine: lineNumberArb,
    span: fc.integer({ min: 0, max: 50 }),
    symbolName: symbolNameArb,
    reason: reasonArb,
  }).map(({ confidence, filePath, startLine, span, symbolName, reason }) => ({
    id: `finding-${category}-${startLine}`,
    category,
    confidence,
    filePath,
    startLine,
    endLine: startLine + span,
    symbolName,
    reason,
    removable: true,
  }));
}

/**
 * Generate a random Finding of any category.
 */
const anyFindingArb: fc.Arbitrary<Finding> = fc.constantFrom(...allCategories).chain(
  (cat) => findingOfCategoryArb(cat),
);

/**
 * Generate a random DuplicatePair.
 */
const duplicatePairArb: fc.Arbitrary<DuplicatePair> = fc.record({
  fileA: filePathArb,
  startLineA: lineNumberArb,
  lineCount: fc.integer({ min: 10, max: 50 }),
  fileB: filePathArb,
  startLineB: lineNumberArb,
  type: fc.constantFrom('exact' as const, 'near-duplicate' as const),
}).map(({ fileA, startLineA, lineCount, fileB, startLineB, type }) => ({
  fileA,
  startLineA,
  endLineA: startLineA + lineCount - 1,
  fileB,
  startLineB,
  endLineB: startLineB + lineCount - 1,
  lineCount,
  type,
}));

// --- Helpers ---

/**
 * Count findings per summary category manually.
 */
function countByCategory(findings: Finding[]) {
  let unusedProcedures = 0;
  let unusedVariables = 0;
  let unusedConstsEnumsTypes = 0;
  let unreachableCode = 0;
  let commentedOutBlocks = 0;

  for (const f of findings) {
    switch (f.category) {
      case 'unused-procedure':
        unusedProcedures++;
        break;
      case 'unused-variable':
      case 'write-only-variable':
        unusedVariables++;
        break;
      case 'unused-const':
      case 'unused-enum':
      case 'unused-type':
      case 'unused-declare':
        unusedConstsEnumsTypes++;
        break;
      case 'unreachable-code':
      case 'dead-branch':
        unreachableCode++;
        break;
      case 'commented-out-block':
        commentedOutBlocks++;
        break;
    }
  }

  return { unusedProcedures, unusedVariables, unusedConstsEnumsTypes, unreachableCode, commentedOutBlocks };
}

// --- Property Tests ---

describe('Feature: dead-code-audit, Property 9: Report summary count accuracy', () => {
  it('summary counts exactly match the number of findings per category', () => {
    /**
     * Validates: Requirements 6.2
     *
     * Strategy:
     * 1. Generate a random set of findings with varying categories and counts
     * 2. Generate a random set of duplicate pairs
     * 3. Call computeSummary to build the summary
     * 4. Manually count findings per category
     * 5. Verify each summary field matches the manual count exactly
     */
    fc.assert(
      fc.property(
        fc.tuple(
          fc.array(anyFindingArb, { minLength: 0, maxLength: 20 }),
          fc.array(duplicatePairArb, { minLength: 0, maxLength: 5 }),
        ),
        ([findings, duplicates]) => {
          const summary = computeSummary(findings, duplicates);
          const expected = countByCategory(findings);

          expect(summary.unusedProcedures).toBe(expected.unusedProcedures);
          expect(summary.unusedVariables).toBe(expected.unusedVariables);
          expect(summary.unusedConstsEnumsTypes).toBe(expected.unusedConstsEnumsTypes);
          expect(summary.unreachableCode).toBe(expected.unreachableCode);
          expect(summary.commentedOutBlocks).toBe(expected.commentedOutBlocks);
          expect(summary.duplicateBlocks).toBe(duplicates.length);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('empty findings produce all-zero summary counts', () => {
    /**
     * Validates: Requirements 6.2
     *
     * Edge case: when there are no findings and no duplicates,
     * all summary counts must be zero.
     */
    fc.assert(
      fc.property(
        fc.constant(null),
        () => {
          const summary = computeSummary([], []);

          expect(summary.unusedProcedures).toBe(0);
          expect(summary.unusedVariables).toBe(0);
          expect(summary.unusedConstsEnumsTypes).toBe(0);
          expect(summary.unreachableCode).toBe(0);
          expect(summary.commentedOutBlocks).toBe(0);
          expect(summary.duplicateBlocks).toBe(0);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('summary counts are additive — combining two finding sets equals sum of individual summaries', () => {
    /**
     * Validates: Requirements 6.2
     *
     * Generate two independent finding sets and duplicate sets.
     * Verify that computeSummary(A ? B) equals computeSummary(A) + computeSummary(B)
     * for each category.
     */
    fc.assert(
      fc.property(
        fc.tuple(
          fc.array(anyFindingArb, { minLength: 0, maxLength: 10 }),
          fc.array(duplicatePairArb, { minLength: 0, maxLength: 3 }),
          fc.array(anyFindingArb, { minLength: 0, maxLength: 10 }),
          fc.array(duplicatePairArb, { minLength: 0, maxLength: 3 }),
        ),
        ([findingsA, dupsA, findingsB, dupsB]) => {
          const summaryA = computeSummary(findingsA, dupsA);
          const summaryB = computeSummary(findingsB, dupsB);
          const summaryCombined = computeSummary(
            [...findingsA, ...findingsB],
            [...dupsA, ...dupsB],
          );

          expect(summaryCombined.unusedProcedures).toBe(
            summaryA.unusedProcedures + summaryB.unusedProcedures,
          );
          expect(summaryCombined.unusedVariables).toBe(
            summaryA.unusedVariables + summaryB.unusedVariables,
          );
          expect(summaryCombined.unusedConstsEnumsTypes).toBe(
            summaryA.unusedConstsEnumsTypes + summaryB.unusedConstsEnumsTypes,
          );
          expect(summaryCombined.unreachableCode).toBe(
            summaryA.unreachableCode + summaryB.unreachableCode,
          );
          expect(summaryCombined.commentedOutBlocks).toBe(
            summaryA.commentedOutBlocks + summaryB.commentedOutBlocks,
          );
          expect(summaryCombined.duplicateBlocks).toBe(
            summaryA.duplicateBlocks + summaryB.duplicateBlocks,
          );
        },
      ),
      { numRuns: 100 },
    );
  });
});
