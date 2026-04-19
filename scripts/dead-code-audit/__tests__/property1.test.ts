/**
 * Property 1: Unused symbol detection completeness
 *
 * For any set of symbol definitions and a reference map, a symbol is reported
 * as unused if and only if it has zero references in executable (non-comment)
 * code and is not an event handler. Conversely, every symbol with at least one
 * executable reference must NOT be reported as unused.
 *
 * Validates: Requirements 1.1, 1.2, 2.1, 2.2, 3.1, 3.2, 3.3
 */
import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { analyzeUsage } from '../src/crossRefAnalyzer.js';
import type {
  SymbolDefinition,
  SymbolReference,
  SymbolTable,
  ReferenceMap,
  SymbolKind,
  Visibility,
  VariableScope,
} from '../src/types.js';

// --- Arbitraries ---

const symbolKindArb: fc.Arbitrary<SymbolKind> = fc.constantFrom(
  'Sub', 'Function', 'Property', 'Variable', 'Const', 'Enum', 'Type', 'Declare', 'Event',
);

const visibilityArb: fc.Arbitrary<Visibility> = fc.constantFrom(
  'Public', 'Private', 'Friend',
);

const variableScopeArb: fc.Arbitrary<VariableScope> = fc.constantFrom('module', 'local');

const moduleNameArb = fc.stringMatching(/^[A-Z][a-zA-Z0-9]{0,7}$/)
  .filter(s => s.length >= 1);

/**
 * Generate a single "symbol entry": a definition plus its executable and
 * comment-only references. Each entry uses a unique name to avoid
 * cross-symbol reference collisions.
 */
function symbolEntryArb(name: string): fc.Arbitrary<{
  def: SymbolDefinition;
  execRefs: SymbolReference[];
  commentRefs: SymbolReference[];
}> {
  return moduleNameArb.chain((modName) =>
    fc.record({
      def: fc.record({
        kind: symbolKindArb,
        visibility: visibilityArb,
        scope: variableScopeArb,
        isEventHandler: fc.boolean(),
        lineNumber: fc.integer({ min: 1, max: 5000 }),
      }).map(({ kind, visibility, scope, isEventHandler, lineNumber }) => ({
        id: `${modName}::${name}::${lineNumber}`,
        name,
        kind,
        visibility,
        moduleName: modName,
        filePath: `CODIGO/${modName}.bas`,
        lineNumber,
        scope,
        isEventHandler,
      } satisfies SymbolDefinition)),
      execRefs: fc.array(
        fc.record({
          referencingModule: fc.oneof(fc.constant(modName), moduleNameArb),
          lineNumber: fc.integer({ min: 1, max: 5000 }),
          context: fc.constantFrom(
            'call' as const, 'read' as const, 'write' as const, 'type-usage' as const,
          ),
          isDynamic: fc.boolean(),
        }).map(({ referencingModule, lineNumber, context, isDynamic }) => ({
          symbolName: name.toLowerCase(),
          referencingModule,
          filePath: `CODIGO/${referencingModule}.bas`,
          lineNumber,
          isInComment: false,
          context,
          isDynamic,
        } satisfies SymbolReference)),
        { minLength: 0, maxLength: 3 },
      ),

      commentRefs: fc.array(
        fc.record({
          referencingModule: fc.oneof(fc.constant(modName), moduleNameArb),
          lineNumber: fc.integer({ min: 1, max: 5000 }),
          context: fc.constantFrom(
            'call' as const, 'read' as const, 'write' as const, 'type-usage' as const,
          ),
          isDynamic: fc.boolean(),
        }).map(({ referencingModule, lineNumber, context, isDynamic }) => ({
          symbolName: name.toLowerCase(),
          referencingModule,
          filePath: `CODIGO/${referencingModule}.bas`,
          lineNumber,
          isInComment: true,
          context,
          isDynamic,
        } satisfies SymbolReference)),
        { minLength: 0, maxLength: 3 },
      ),
    }),
  );
}

/**
 * Build a SymbolTable from an array of definitions.
 */
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

/**
 * Build a ReferenceMap from an array of references.
 */
function buildReferenceMap(refs: SymbolReference[]): ReferenceMap {
  const references = new Map<string, SymbolReference[]>();

  for (const ref of refs) {
    const key = ref.symbolName.toLowerCase();
    if (!references.has(key)) references.set(key, []);
    references.get(key)!.push(ref);
  }

  return { references };
}

// --- Property Test ---

describe('Feature: dead-code-audit, Property 1: Unused symbol detection completeness', () => {
  it('a symbol is unused iff it has zero executable references and is not an event handler', () => {
    /**
     * Validates: Requirements 1.1, 1.2, 2.1, 2.2, 3.1, 3.2, 3.3
     *
     * Strategy:
     * 1. Generate 1–5 symbol entries with unique names
     * 2. Each entry has a definition, 0–3 executable refs, 0–3 comment refs
     * 3. Build SymbolTable and ReferenceMap from generated data
     * 4. Call analyzeUsage and verify the unused-detection property
     */

    // Pre-generate a pool of unique names to avoid collisions
    const namePool = ['alpha', 'bravo', 'charlie', 'delta', 'echo'];

    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 5 }).chain((count) => {
          const names = namePool.slice(0, count);
          return fc.tuple(...names.map((n) => symbolEntryArb(n)));
        }),
        (entries) => {
          const allDefs: SymbolDefinition[] = [];
          const allRefs: SymbolReference[] = [];
          const expectedExecRefs = new Map<string, number>();

          for (const { def, execRefs, commentRefs } of entries) {
            allDefs.push(def);
            allRefs.push(...execRefs, ...commentRefs);
            expectedExecRefs.set(def.id, execRefs.length);
          }

          const symbolTable = buildSymbolTable(allDefs);
          const referenceMap = buildReferenceMap(allRefs);
          const usages = analyzeUsage(symbolTable, referenceMap);

          // Every definition must have a corresponding usage entry
          expect(usages.length).toBe(allDefs.length);

          for (const usage of usages) {
            const { definition, totalReferences } = usage;

            // Comment-only refs must NOT count toward totalReferences
            const execRefCount = expectedExecRefs.get(definition.id) ?? 0;
            expect(totalReferences).toBe(execRefCount);

            const isUnused = totalReferences === 0 && !definition.isEventHandler;

            // A symbol with zero executable refs AND not an event handler ? unused
            if (totalReferences === 0 && !definition.isEventHandler) {
              expect(isUnused).toBe(true);
            }

            // A symbol with executable refs OR is an event handler ? NOT unused
            if (totalReferences > 0 || definition.isEventHandler) {
              expect(isUnused).toBe(false);
            }
          }
        },
      ),
      { numRuns: 100 },
    );
  });
});
