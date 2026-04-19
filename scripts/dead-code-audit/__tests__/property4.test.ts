/**
 * Property 4: Write-only variable classification
 *
 * For any variable that has one or more write references but zero read references
 * in executable code, the variable must be classified as "write-only"
 * (writeCount > 0 AND readCount === 0).
 * Variables with both reads and writes must NOT be classified as write-only.
 * Variables with zero writes and zero reads must be classified as "unused"
 * (not "write-only").
 *
 * Validates: Requirements 2.4
 */
import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { analyzeUsage } from '../src/crossRefAnalyzer.js';
import type {
  SymbolDefinition,
  SymbolReference,
  SymbolTable,
  ReferenceMap,
  Visibility,
  VariableScope,
} from '../src/types.js';

// --- Arbitraries ---

const visibilityArb: fc.Arbitrary<Visibility> = fc.constantFrom(
  'Public', 'Private', 'Friend',
);

const variableScopeArb: fc.Arbitrary<VariableScope> = fc.constantFrom('module', 'local');

const moduleNameArb = fc.stringMatching(/^[A-Z][a-zA-Z0-9]{0,7}$/)
  .filter(s => s.length >= 1);

/**
 * Generate a Variable definition with a given unique name.
 * kind is always 'Variable' to focus on write-only classification.
 */
function variableDefArb(name: string): fc.Arbitrary<SymbolDefinition> {
  return fc.record({
    visibility: visibilityArb,
    scope: variableScopeArb,
    lineNumber: fc.integer({ min: 1, max: 5000 }),
    moduleName: moduleNameArb,
  }).map(({ visibility, scope, lineNumber, moduleName }) => ({
    id: `${moduleName}::${name}::${lineNumber}`,
    name,
    kind: 'Variable' as const,
    visibility,
    moduleName,
    filePath: `CODIGO/${moduleName}.bas`,
    lineNumber,
    scope,
    isEventHandler: false,
    dataType: 'Long',
  }));
}

/**
 * Generate write-only references (context: 'write', isInComment: false).
 */
function writeRefsArb(symbolName: string): fc.Arbitrary<SymbolReference[]> {
  return fc.array(
    fc.record({
      referencingModule: moduleNameArb,
      lineNumber: fc.integer({ min: 1, max: 5000 }),
      isDynamic: fc.constant(false),
    }).map(({ referencingModule, lineNumber, isDynamic }) => ({
      symbolName: symbolName.toLowerCase(),
      referencingModule,
      filePath: `CODIGO/${referencingModule}.bas`,
      lineNumber,
      isInComment: false,
      context: 'write' as const,
      isDynamic,
    })),
    { minLength: 1, maxLength: 5 },
  );
}

/**
 * Generate read-only references (context: 'read', isInComment: false).
 */
function readRefsArb(symbolName: string): fc.Arbitrary<SymbolReference[]> {
  return fc.array(
    fc.record({
      referencingModule: moduleNameArb,
      lineNumber: fc.integer({ min: 1, max: 5000 }),
      isDynamic: fc.constant(false),
    }).map(({ referencingModule, lineNumber, isDynamic }) => ({
      symbolName: symbolName.toLowerCase(),
      referencingModule,
      filePath: `CODIGO/${referencingModule}.bas`,
      lineNumber,
      isInComment: false,
      context: 'read' as const,
      isDynamic,
    })),
    { minLength: 1, maxLength: 5 },
  );
}

// --- Helpers ---

function buildSymbolTable(definitions: SymbolDefinition[]): SymbolTable {
  const symbols = new Map<string, SymbolDefinition[]>();
  const byModule = new Map<string, SymbolDefinition[]>();

  for (const def of definitions) {
    const key = def.name.toLowerCase();
    if (!symbols.has(key)) symbols.set(key, []);
    symbols.get(key)!.push(def);

    if (!byModule.has(def.moduleName)) byModule.set(def.moduleName, []);
    byModule.get(def.moduleName)!.push(def);
  }

  return { symbols, byModule };
}

function buildReferenceMap(refs: SymbolReference[]): ReferenceMap {
  const references = new Map<string, SymbolReference[]>();

  for (const ref of refs) {
    const key = ref.symbolName.toLowerCase();
    if (!references.has(key)) references.set(key, []);
    references.get(key)!.push(ref);
  }

  return { references };
}


// --- Property Tests ---

describe('Feature: dead-code-audit, Property 4: Write-only variable classification', () => {
  it('variables with writes but no reads are classified as write-only', () => {
    /**
     * Validates: Requirements 2.4
     *
     * Strategy:
     * 1. Generate a Variable definition with a unique name
     * 2. Generate 1+ write references and 0 read references
     * 3. Build SymbolTable and ReferenceMap
     * 4. Call analyzeUsage
     * 5. Verify writeCount > 0 AND readCount === 0 (write-only)
     */
    const namePool = ['varAlpha', 'varBravo', 'varCharlie', 'varDelta', 'varEcho'];

    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 5 }).chain((count) => {
          const names = namePool.slice(0, count);
          return fc.tuple(
            ...names.map((n) =>
              variableDefArb(n).chain((def) =>
                writeRefsArb(n).map((writeRefs) => ({ def, writeRefs })),
              ),
            ),
          );
        }),
        (entries) => {
          const allDefs: SymbolDefinition[] = [];
          const allRefs: SymbolReference[] = [];

          for (const { def, writeRefs } of entries) {
            allDefs.push(def);
            allRefs.push(...writeRefs);
          }

          const symbolTable = buildSymbolTable(allDefs);
          const referenceMap = buildReferenceMap(allRefs);
          const usages = analyzeUsage(symbolTable, referenceMap);

          expect(usages.length).toBe(allDefs.length);

          for (const usage of usages) {
            // Write-only: writeCount > 0 AND readCount === 0
            expect(usage.writeCount).toBeGreaterThan(0);
            expect(usage.readCount).toBe(0);

            // totalReferences should equal writeCount (only writes, no reads)
            expect(usage.totalReferences).toBe(usage.writeCount);
          }
        },
      ),
      { numRuns: 100 },
    );
  });

  it('variables with both reads and writes are NOT classified as write-only', () => {
    /**
     * Validates: Requirements 2.4
     *
     * Strategy:
     * 1. Generate a Variable definition with a unique name
     * 2. Generate 1+ write references AND 1+ read references
     * 3. Build SymbolTable and ReferenceMap
     * 4. Call analyzeUsage
     * 5. Verify writeCount > 0 AND readCount > 0 (NOT write-only)
     */
    const namePool = ['mixAlpha', 'mixBravo', 'mixCharlie', 'mixDelta', 'mixEcho'];

    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 5 }).chain((count) => {
          const names = namePool.slice(0, count);
          return fc.tuple(
            ...names.map((n) =>
              variableDefArb(n).chain((def) =>
                fc.tuple(writeRefsArb(n), readRefsArb(n)).map(
                  ([writeRefs, readRefs]) => ({ def, writeRefs, readRefs }),
                ),
              ),
            ),
          );
        }),
        (entries) => {
          const allDefs: SymbolDefinition[] = [];
          const allRefs: SymbolReference[] = [];

          for (const { def, writeRefs, readRefs } of entries) {
            allDefs.push(def);
            allRefs.push(...writeRefs, ...readRefs);
          }

          const symbolTable = buildSymbolTable(allDefs);
          const referenceMap = buildReferenceMap(allRefs);
          const usages = analyzeUsage(symbolTable, referenceMap);

          expect(usages.length).toBe(allDefs.length);

          for (const usage of usages) {
            // Both reads and writes present ? NOT write-only
            expect(usage.writeCount).toBeGreaterThan(0);
            expect(usage.readCount).toBeGreaterThan(0);
          }
        },
      ),
      { numRuns: 100 },
    );
  });

  it('variables with zero writes and zero reads are classified as unused, not write-only', () => {
    /**
     * Validates: Requirements 2.4
     *
     * Strategy:
     * 1. Generate a Variable definition with a unique name
     * 2. Generate NO references at all (empty reference map)
     * 3. Build SymbolTable and ReferenceMap
     * 4. Call analyzeUsage
     * 5. Verify writeCount === 0 AND readCount === 0 (unused, not write-only)
     */
    const namePool = ['unusedA', 'unusedB', 'unusedC', 'unusedD', 'unusedE'];

    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 5 }).chain((count) => {
          const names = namePool.slice(0, count);
          return fc.tuple(...names.map((n) => variableDefArb(n)));
        }),
        (defs) => {
          const allDefs = [...defs];

          const symbolTable = buildSymbolTable(allDefs);
          // Empty reference map — no references at all
          const referenceMap = buildReferenceMap([]);
          const usages = analyzeUsage(symbolTable, referenceMap);

          expect(usages.length).toBe(allDefs.length);

          for (const usage of usages) {
            // Unused: writeCount === 0 AND readCount === 0
            expect(usage.writeCount).toBe(0);
            expect(usage.readCount).toBe(0);
            expect(usage.totalReferences).toBe(0);

            // This is "unused", NOT "write-only"
            // write-only requires writeCount > 0
            const isWriteOnly = usage.writeCount > 0 && usage.readCount === 0;
            expect(isWriteOnly).toBe(false);
          }
        },
      ),
      { numRuns: 100 },
    );
  });
});
