/**
 * Property 2: Finding report field completeness
 *
 * For any finding produced by the detector, the formatted report entry must
 * contain all required fields for that finding's category: file name, line number,
 * symbol name (when applicable), and category-specific fields (visibility for
 * procedures, scope and data type for variables, symbol kind for constants/enums/types,
 * start and end line for unreachable code, and confidence level).
 *
 * Validates: Requirements 1.3, 2.3, 3.4, 4.3, 5.3, 6.3
 */
import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { generateReport } from '../src/reportGenerator.js';
import type {
  Finding,
  FindingCategory,
  Confidence,
  DuplicatePair,
  AuditReport,
} from '../src/types.js';

// --- Arbitraries ---

const confidenceArb: fc.Arbitrary<Confidence> = fc.constantFrom('confirmed', 'review-needed');

const filePathArb = fc.stringMatching(/^CODIGO\/[A-Z][a-zA-Z0-9]{1,8}\.(bas|cls|frm)$/)
  .filter(s => s.length > 8);

const symbolNameArb = fc.stringMatching(/^[a-zA-Z][a-zA-Z0-9_]{1,12}$/)
  .filter(s => s.length >= 2);

const lineNumberArb = fc.integer({ min: 1, max: 5000 });

const reasonArb = fc.stringMatching(/^[A-Za-z ]{5,30}$/).filter(s => s.length >= 5);

/** Categories that represent symbol-based findings (have symbolName). */
const symbolCategories: FindingCategory[] = [
  'unused-procedure',
  'unused-variable',
  'write-only-variable',
  'unused-const',
  'unused-enum',
  'unused-type',
  'unused-declare',
];

/** Categories that represent code-block findings (unreachable/commented-out/dead-branch). */
const blockCategories: FindingCategory[] = [
  'unreachable-code',
  'dead-branch',
  'commented-out-block',
];

const allCategories: FindingCategory[] = [...symbolCategories, ...blockCategories];

/**
 * Generate a Finding for a symbol-based category.
 */
function symbolFindingArb(category: FindingCategory): fc.Arbitrary<Finding> {
  return fc.record({
    confidence: confidenceArb,
    filePath: filePathArb,
    startLine: lineNumberArb,
    symbolName: symbolNameArb,
    reason: reasonArb,
  }).map(({ confidence, filePath, startLine, symbolName, reason }) => ({
    id: `finding-${category}-${startLine}`,
    category,
    confidence,
    filePath,
    startLine,
    endLine: startLine,
    symbolName,
    reason,
    removable: true,
  }));
}

/**
 * Generate a Finding for a block-based category (unreachable-code, dead-branch, commented-out-block).
 */
function blockFindingArb(category: FindingCategory): fc.Arbitrary<Finding> {
  return fc.record({
    confidence: confidenceArb,
    filePath: filePathArb,
    startLine: lineNumberArb,
    span: fc.integer({ min: 1, max: 50 }),
    reason: reasonArb,
  }).map(({ confidence, filePath, startLine, span, reason }) => ({
    id: `finding-${category}-${startLine}`,
    category,
    confidence,
    filePath,
    startLine,
    endLine: startLine + span,
    reason,
    removable: true,
  }));
}

/**
 * Generate a DuplicatePair entry.
 */
const duplicatePairArb: fc.Arbitrary<DuplicatePair> = fc.record({
  fileA: filePathArb,
  startLineA: lineNumberArb,
  lineCountA: fc.integer({ min: 10, max: 50 }),
  fileB: filePathArb,
  startLineB: lineNumberArb,
  lineCountB: fc.integer({ min: 10, max: 50 }),
  type: fc.constantFrom('exact' as const, 'near-duplicate' as const),
}).map(({ fileA, startLineA, lineCountA, fileB, startLineB, lineCountB, type }) => ({
  fileA,
  startLineA,
  endLineA: startLineA + lineCountA - 1,
  fileB,
  startLineB,
  endLineB: startLineB + lineCountB - 1,
  lineCount: lineCountA,
  type,
}));

/**
 * Generate a Finding for any category.
 */
const anyFindingArb: fc.Arbitrary<Finding> = fc.constantFrom(...allCategories).chain(
  (cat) => symbolCategories.includes(cat) ? symbolFindingArb(cat) : blockFindingArb(cat),
);

/**
 * Build an AuditReport from findings and duplicates.
 */
function buildReport(findings: Finding[], duplicates: DuplicatePair[]): AuditReport {
  let unusedProcedures = 0;
  let unusedVariables = 0;
  let unusedConstsEnumsTypes = 0;
  let unreachableCode = 0;
  let commentedOutBlocks = 0;

  for (const f of findings) {
    switch (f.category) {
      case 'unused-procedure': unusedProcedures++; break;
      case 'unused-variable':
      case 'write-only-variable': unusedVariables++; break;
      case 'unused-const':
      case 'unused-enum':
      case 'unused-type':
      case 'unused-declare': unusedConstsEnumsTypes++; break;
      case 'unreachable-code':
      case 'dead-branch': unreachableCode++; break;
      case 'commented-out-block': commentedOutBlocks++; break;
    }
  }

  return {
    codebase: 'client',
    timestamp: '2025-01-15T12:00:00Z',
    summary: {
      unusedProcedures,
      unusedVariables,
      unusedConstsEnumsTypes,
      unreachableCode,
      commentedOutBlocks,
      duplicateBlocks: duplicates.length,
    },
    findings,
    duplicates,
  };
}

// --- Property Tests ---

describe('Feature: dead-code-audit, Property 2: Finding report field completeness', () => {
  it('symbol-based findings contain file name, line number, symbol name, and confidence', () => {
    /**
     * Validates: Requirements 1.3, 2.3, 3.4, 6.3
     *
     * For each symbol-based category (unused-procedure, unused-variable,
     * write-only-variable, unused-const, unused-enum, unused-type, unused-declare),
     * the report entry must contain the file path, line number, symbol name,
     * and confidence level.
     */
    fc.assert(
      fc.property(
        fc.constantFrom(...symbolCategories).chain((cat) =>
          symbolFindingArb(cat).map((finding) => ({ cat, finding })),
        ),
        ({ cat, finding }) => {
          const report = buildReport([finding], []);
          const markdown = generateReport(report);

          // The report must contain the file path
          expect(markdown).toContain(finding.filePath);
          // The report must contain the line number
          expect(markdown).toContain(`${finding.startLine}`);
          // The report must contain the symbol name
          expect(markdown).toContain(finding.symbolName!);
          // The report must contain the confidence level
          expect(markdown).toContain(finding.confidence);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('unreachable code and dead branch findings contain file name, start line, end line, and confidence', () => {
    /**
     * Validates: Requirements 4.3, 6.3
     *
     * For unreachable-code and dead-branch categories, the report entry must
     * contain the file path, start line, end line (as a range), and confidence.
     */
    fc.assert(
      fc.property(
        fc.constantFrom('unreachable-code' as FindingCategory, 'dead-branch' as FindingCategory).chain(
          (cat) => blockFindingArb(cat),
        ),
        (finding) => {
          const report = buildReport([finding], []);
          const markdown = generateReport(report);

          // The report must contain the file path
          expect(markdown).toContain(finding.filePath);
          // The report must contain the line range (startLine-endLine)
          expect(markdown).toContain(`${finding.startLine}-${finding.endLine}`);
          // The report must contain the confidence level
          expect(markdown).toContain(finding.confidence);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('duplicate findings contain both file locations, line ranges, line count, and type', () => {
    /**
     * Validates: Requirements 5.3, 6.3
     *
     * For duplicate code entries, the report must contain both file paths,
     * both line ranges, the line count, and the duplicate type (exact/near-duplicate).
     */
    fc.assert(
      fc.property(
        duplicatePairArb,
        (dup) => {
          const report = buildReport([], [dup]);
          const markdown = generateReport(report);

          // Both file locations
          expect(markdown).toContain(dup.fileA);
          expect(markdown).toContain(dup.fileB);
          // Line ranges for both
          expect(markdown).toContain(`${dup.startLineA}-${dup.endLineA}`);
          expect(markdown).toContain(`${dup.startLineB}-${dup.endLineB}`);
          // Line count
          expect(markdown).toContain(`${dup.lineCount}`);
          // Type
          expect(markdown).toContain(dup.type);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('mixed findings report contains all required fields for every category', () => {
    /**
     * Validates: Requirements 1.3, 2.3, 3.4, 4.3, 5.3, 6.3
     *
     * Generate a set of findings spanning all categories plus duplicates.
     * Verify every finding's required fields appear in the generated report.
     */
    fc.assert(
      fc.property(
        fc.tuple(
          fc.array(anyFindingArb, { minLength: 1, maxLength: 10 }),
          fc.array(duplicatePairArb, { minLength: 0, maxLength: 3 }),
        ),
        ([findings, duplicates]) => {
          const report = buildReport(findings, duplicates);
          const markdown = generateReport(report);

          for (const f of findings) {
            // Every finding must have file path and confidence in the report
            expect(markdown).toContain(f.filePath);
            expect(markdown).toContain(f.confidence);

            if (f.symbolName) {
              // Symbol-based findings must include the symbol name
              expect(markdown).toContain(f.symbolName);
            }

            if (f.startLine !== f.endLine) {
              // Block findings must include the line range
              expect(markdown).toContain(`${f.startLine}-${f.endLine}`);
            } else {
              // Single-line findings must include the line number
              expect(markdown).toContain(`${f.startLine}`);
            }
          }

          for (const d of duplicates) {
            expect(markdown).toContain(d.fileA);
            expect(markdown).toContain(d.fileB);
            expect(markdown).toContain(`${d.lineCount}`);
            expect(markdown).toContain(d.type);
          }
        },
      ),
      { numRuns: 100 },
    );
  });
});
