/**
 * Property 12: Removal summary tallying
 *
 * For any set of removal results, the removal summary must report per-module
 * line counts that sum to the total lines removed, and the total must equal
 * the sum of all individual batch line counts for successful removals.
 *
 * Validates: Requirements 7.6
 */
import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { generateRemovalSummary } from '../src/removalEngine.js';
import type { RemovalBatch, RemovalResult, Finding } from '../src/types.js';

// --- Arbitraries ---

const moduleNameArb = fc.stringMatching(/^[A-Z][a-zA-Z0-9]{0,7}$/)
  .filter(s => s.length >= 1);

/**
 * Generate a RemovalBatch with a given module name and random linesRemoved.
 */
function removalBatchArb(moduleName: string): fc.Arbitrary<RemovalBatch> {
  return fc.integer({ min: 1, max: 200 }).map((linesRemoved) => ({
    filePath: `CODIGO/${moduleName}.bas`,
    moduleName,
    findings: [] as Finding[], // Findings content not needed for summary tallying
    linesRemoved,
  }));
}

/**
 * Generate a RemovalResult with varying success/reverted states.
 */
function removalResultArb(moduleName: string): fc.Arbitrary<RemovalResult> {
  return removalBatchArb(moduleName).chain((batch) =>
    fc.constantFrom(
      // Successful removal
      { success: true, syntaxValid: true, testsPass: null as boolean | null, reverted: false },
      // Failed and reverted
      { success: false, syntaxValid: false, testsPass: null as boolean | null, reverted: true },
      // Failed without revert
      { success: false, syntaxValid: false, testsPass: null as boolean | null, reverted: false },
    ).map((status) => ({
      batch,
      ...status,
    })),
  );
}

/**
 * Parse the total lines removed from the summary markdown output.
 */
function parseTotalFromSummary(summary: string): number {
  const match = summary.match(/\*\*Total lines removed:\s*(\d+)\*\*/);
  return match ? parseInt(match[1], 10) : -1;
}

/**
 * Parse per-module line counts from the summary markdown table.
 * Returns an array of { moduleName, linesRemoved } entries.
 */
function parseModuleCountsFromSummary(summary: string): { moduleName: string; linesRemoved: number }[] {
  const results: { moduleName: string; linesRemoved: number }[] = [];
  const lines = summary.split('\n');

  for (const line of lines) {
    // Match table rows: | ModuleName | 42 | Removed |
    const match = line.match(/^\|\s*(\S+)\s*\|\s*(\d+)\s*\|/);
    if (match && match[1] !== 'Module' && match[1] !== '--------') {
      results.push({
        moduleName: match[1],
        linesRemoved: parseInt(match[2], 10),
      });
    }
  }

  return results;
}

// --- Property Tests ---

describe('Feature: dead-code-audit, Property 12: Removal summary tallying', () => {
  it('per-module counts sum to total and total equals sum of successful batch line counts', () => {
    /**
     * Validates: Requirements 7.6
     *
     * Strategy:
     * 1. Generate 1–8 RemovalResult entries with unique module names
     * 2. Call generateRemovalSummary
     * 3. Parse the output to extract per-module counts and total
     * 4. Verify: sum of per-module counts === reported total
     * 5. Verify: total === sum of linesRemoved for successful batches only
     */
    const modulePool = ['ModA', 'ModB', 'ModC', 'ModD', 'ModE', 'ModF', 'ModG', 'ModH'];

    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 8 }).chain((count) => {
          const names = modulePool.slice(0, count);
          return fc.tuple(...names.map(n => removalResultArb(n)));
        }),
        (results) => {
          const summary = generateRemovalSummary(results);

          // Parse total from summary
          const reportedTotal = parseTotalFromSummary(summary);
          expect(reportedTotal).toBeGreaterThanOrEqual(0);

          // Parse per-module counts
          const moduleCounts = parseModuleCountsFromSummary(summary);

          // Sum of per-module counts must equal reported total
          const sumOfModuleCounts = moduleCounts.reduce(
            (sum, m) => sum + m.linesRemoved, 0,
          );
          expect(sumOfModuleCounts).toBe(reportedTotal);

          // Total must equal sum of linesRemoved for successful batches
          const expectedTotal = results.reduce(
            (sum, r) => sum + (r.success ? r.batch.linesRemoved : 0), 0,
          );
          expect(reportedTotal).toBe(expectedTotal);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('every result module appears in the summary output', () => {
    /**
     * Validates: Requirements 7.6
     *
     * Strategy:
     * 1. Generate RemovalResult entries with unique module names
     * 2. Call generateRemovalSummary
     * 3. Verify every module name from the input appears in the summary
     */
    const modulePool = ['ResA', 'ResB', 'ResC', 'ResD', 'ResE'];

    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 5 }).chain((count) => {
          const names = modulePool.slice(0, count);
          return fc.tuple(...names.map(n => removalResultArb(n)));
        }),
        (results) => {
          const summary = generateRemovalSummary(results);

          for (const r of results) {
            expect(summary).toContain(r.batch.moduleName);
          }
        },
      ),
      { numRuns: 100 },
    );
  });
});
