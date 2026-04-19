/**
 * Property 3: Comment-aware reference exclusion
 *
 * For any symbol that is referenced exclusively within commented-out lines
 * (lines starting with `'`), the symbol must be classified as unused.
 * References in comments must not count toward a symbol's active reference count.
 *
 * Validates: Requirements 1.4
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

/** Non-event-handler symbol kinds to isolate the comment property. */
const symbolKindArb: fc.Arbitrary<SymbolKind> = fc.constantFrom(
  'Sub', 'Function', 'Property', 'Variable', 'Const', 'Enum', 'Type', 'Declare',
);

const visibilityArb: fc.Arbitrary<Visibility> = fc.constantFrom(
  'Public', 'Private', 'Friend',
);

const variableScopeArb: fc.Arbitrary<VariableScope> = fc.constantFrom('module', 'local');

const moduleNameArb = fc.stringMatching(/^[A-Z][a-zA-Z0-9]{0,7}$/)
  .filter(s => s.length >= 1);

const refContextArb = fc.constantFrom(
  'call' as const, 'read' as const, 'write' as const, 'type-usage' as const,
);

/**
 * Generate a symbol definition (non-event-handler) with a given unique name.
 */
function symbolDefArb(name: string): fc.Arbitrary<SymbolDefinition> {
  return fc.record({
    kind: symbolKindArb,
    visibility: visibilityArb,
    scope: variableScopeArb,
    lineNumber: fc.integer({ min: 1, max: 5000 }),
    moduleName: moduleNameArb,
  }).map(({ kind, visibility, scope, lineNumber, moduleName }) => ({
    id: `${moduleName}::${name}::${lineNumber}`,
    name,
    kind,
    visibility,
    moduleName,
    filePath: `CODIGO/${moduleName}.bas`,
    lineNumber,
    scope,
    isEventHandler: false,
  }));
}

/**
 * Generate comment-only references (isInComment: true) for a given symbol name.
 */
function commentRefsArb(symbolName: string): fc.Arbitrary<SymbolReference[]> {
  return fc.array(
    fc.record({
      referencingModule: moduleNameArb,
      lineNumber: fc.integer({ min: 1, max: 5000 }),
      context: refContextArb,
      isDynamic: fc.boolean(),
    }).map(({ referencingModule, lineNumber, context, isDynamic }) => ({
      symbolName: symbolName.toLowerCase(),
      referencingModule,
      filePath: `CODIGO/${referencingModule}.bas`,
      lineNumber,
      isInComment: true,
      context,
      isDynamic,
    })),
    { minLength: 1, maxLength: 5 },
  );
}


/**
 * Generate a symbol entry: a definition plus ONLY comment references.
 * No executable references are generated — all refs are isInComment: true.
 */
function commentOnlySymbolEntryArb(name: string): fc.Arbitrary<{
  def: SymbolDefinition;
  commentRefs: SymbolReference[];
}> {
  return symbolDefArb(name).chain((def) =>
    commentRefsArb(name).map((commentRefs) => ({ def, commentRefs })),
  );
}

// --- Helpers ---

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

describe('Feature: dead-code-audit, Property 3: Comment-aware reference exclusion', () => {
  it('symbols referenced only in comments have totalReferences === 0 and commentOnlyRefs matches comment ref count', () => {
    /**
     * Validates: Requirements 1.4
     *
     * Strategy:
     * 1. Generate 1–5 non-event-handler symbol definitions with unique names
     * 2. Generate ONLY comment references (isInComment: true) for each symbol
     * 3. Build SymbolTable and ReferenceMap
     * 4. Call analyzeUsage
     * 5. Verify totalReferences === 0 for all symbols (all refs are comment-only)
     * 6. Verify commentOnlyRefs matches the number of comment references generated
     */

    const namePool = ['alpha', 'bravo', 'charlie', 'delta', 'echo'];

    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 5 }).chain((count) => {
          const names = namePool.slice(0, count);
          return fc.tuple(...names.map((n) => commentOnlySymbolEntryArb(n)));
        }),
        (entries) => {
          const allDefs: SymbolDefinition[] = [];
          const allRefs: SymbolReference[] = [];
          const expectedCommentRefs = new Map<string, number>();

          for (const { def, commentRefs } of entries) {
            allDefs.push(def);
            allRefs.push(...commentRefs);

            const key = def.name.toLowerCase();
            const prev = expectedCommentRefs.get(key) ?? 0;
            expectedCommentRefs.set(key, prev + commentRefs.length);
          }

          const symbolTable = buildSymbolTable(allDefs);
          const referenceMap = buildReferenceMap(allRefs);
          const usages = analyzeUsage(symbolTable, referenceMap);

          // Every definition must have a corresponding usage entry
          expect(usages.length).toBe(allDefs.length);

          for (const usage of usages) {
            const { definition, totalReferences, commentOnlyRefs } = usage;

            // Comment-only references must NOT count toward totalReferences
            expect(totalReferences).toBe(0);

            // commentOnlyRefs must match the number of comment references generated
            // for this symbol name (all definitions with the same name share refs)
            const expectedCount = expectedCommentRefs.get(definition.name.toLowerCase()) ?? 0;
            expect(commentOnlyRefs).toBe(expectedCount);

            // Since totalReferences === 0 and isEventHandler === false,
            // the symbol must be classified as unused
            const isUnused = totalReferences === 0 && !definition.isEventHandler;
            expect(isUnused).toBe(true);
          }
        },
      ),
      { numRuns: 100 },
    );
  });
});
