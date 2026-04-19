/**
 * Property 10: Batch grouping by module
 *
 * For any set of confirmed findings, each removal batch must contain findings
 * from exactly one module file. No batch may mix findings from different files.
 *
 * Validates: Requirements 7.1
 */
import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { createBatches } from '../src/removalEngine.js';
import type { Finding, FindingCategory } from '../src/types.js';

// --- Arbitraries ---

const filePathArb = fc.constantFrom(
  'CODIGO/ModuleA.bas',
  'CODIGO/ModuleB.bas',
  'CODIGO/ModuleC.cls',
  'CODIGO/FormD.frm',
  'CODIGO/ModuleE.bas',
);

const categoryArb: fc.Arbitrary<FindingCategory> = fc.constantFrom(
  'unused-procedure',
  'unused-variable',
  'unused-const',
  'unused-enum',
  'unused-type',
  'unused-declare',
  'write-only-variable',
);

/**
 * Generate a confirmed, removable finding for a given file path.
 */
function findingArb(filePath: fc.Arbitrary<string>): fc.Arbitrary<Finding> {
  return fc.record({
    filePath,
    category: categoryArb,
    startLine: fc.integer({ min: 1, max: 500 }),
    lineSpan: fc.integer({ min: 1, max: 10 }),
    symbolName: fc.stringMatching(/^[a-zA-Z][a-zA-Z0-9]{0,7}$/).filter(s => s.length >= 1),
  }).map(({ filePath, category, startLine, lineSpan, symbolName }, idx) => ({
    id: `finding-${startLine}-${symbolName}`,
    category,
    confidence: 'confirmed' as const,
    filePath,
    startLine,
    endLine: startLine + lineSpan - 1,
    symbolName,
    reason: `Unused ${category}`,
    removable: true,
  }));
}

// --- Property Test ---

describe('Feature: dead-code-audit, Property 10: Batch grouping by module', () => {
  it('each batch contains findings from exactly one file path', () => {
    /**
     * Validates: Requirements 7.1
     *
     * Strategy:
     * 1. Generate 2–20 findings spread across 2–5 different file paths
     * 2. Call createBatches
     * 3. Verify each batch has findings from exactly one file path
     * 4. Verify no batch mixes findings from different files
     */
    fc.assert(
      fc.property(
        fc.array(findingArb(filePathArb), { minLength: 2, maxLength: 20 }),
        (findings) => {
          const batches = createBatches(findings);

          for (const batch of batches) {
            // Every finding in the batch must share the same filePath
            const paths = new Set(batch.findings.map(f => f.filePath));
            expect(paths.size).toBe(1);

            // The batch's filePath must match the findings' filePath
            expect(batch.filePath).toBe(batch.findings[0].filePath);
          }
        },
      ),
      { numRuns: 100 },
    );
  });

  it('all confirmed removable findings are included in exactly one batch', () => {
    /**
     * Validates: Requirements 7.1
     *
     * Strategy:
     * 1. Generate findings from multiple files
     * 2. Call createBatches
     * 3. Verify every confirmed+removable finding appears in exactly one batch
     */
    fc.assert(
      fc.property(
        fc.array(findingArb(filePathArb), { minLength: 2, maxLength: 20 }),
        (findings) => {
          const batches = createBatches(findings);

          // Collect all finding IDs from batches
          const batchedIds = new Set<string>();
          for (const batch of batches) {
            for (const f of batch.findings) {
              // No finding should appear in multiple batches
              expect(batchedIds.has(f.id)).toBe(false);
              batchedIds.add(f.id);
            }
          }

          // Every confirmed+removable input finding should be batched
          const eligible = findings.filter(
            f => f.confidence === 'confirmed' && f.removable === true,
          );
          expect(batchedIds.size).toBe(eligible.length);
        },
      ),
      { numRuns: 100 },
    );
  });
});
