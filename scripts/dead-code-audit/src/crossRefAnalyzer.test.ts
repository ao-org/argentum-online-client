import { describe, it, expect } from 'vitest';
import { analyzeUsage } from './crossRefAnalyzer.js';
import type {
  SymbolTable,
  ReferenceMap,
  SymbolDefinition,
  SymbolReference,
} from './types.js';

// Helper to create a minimal SymbolDefinition
function makeDef(overrides: Partial<SymbolDefinition> & { name: string; moduleName: string }): SymbolDefinition {
  return {
    id: `${overrides.moduleName}::${overrides.name}::${overrides.lineNumber ?? 1}`,
    kind: 'Sub',
    visibility: 'Public',
    filePath: `${overrides.moduleName}.bas`,
    lineNumber: 1,
    scope: 'module',
    isEventHandler: false,
    ...overrides,
  };
}

// Helper to create a minimal SymbolReference
function makeRef(overrides: Partial<SymbolReference> & { symbolName: string; referencingModule: string }): SymbolReference {
  return {
    filePath: `${overrides.referencingModule}.bas`,
    lineNumber: 10,
    isInComment: false,
    context: 'call',
    isDynamic: false,
    ...overrides,
  };
}

// Helper to build a SymbolTable from definitions
function buildTable(defs: SymbolDefinition[]): SymbolTable {
  const symbols = new Map<string, SymbolDefinition[]>();
  const byModule = new Map<string, SymbolDefinition[]>();
  for (const d of defs) {
    const key = d.name.toLowerCase();
    if (!symbols.has(key)) symbols.set(key, []);
    symbols.get(key)!.push(d);
    if (!byModule.has(d.moduleName)) byModule.set(d.moduleName, []);
    byModule.get(d.moduleName)!.push(d);
  }
  return { symbols, byModule };
}

// Helper to build a ReferenceMap from references
function buildRefMap(refs: SymbolReference[]): ReferenceMap {
  const references = new Map<string, SymbolReference[]>();
  for (const r of refs) {
    const key = r.symbolName.toLowerCase();
    if (!references.has(key)) references.set(key, []);
    references.get(key)!.push(r);
  }
  return { references };
}

describe('analyzeUsage', () => {
  it('returns empty array for empty symbol table', () => {
    const table = buildTable([]);
    const refMap = buildRefMap([]);
    const result = analyzeUsage(table, refMap);
    expect(result).toEqual([]);
  });

  it('reports zero references for a symbol with no references', () => {
    const def = makeDef({ name: 'DoStuff', moduleName: 'ModA' });
    const table = buildTable([def]);
    const refMap = buildRefMap([]);
    const result = analyzeUsage(table, refMap);

    expect(result).toHaveLength(1);
    expect(result[0].totalReferences).toBe(0);
    expect(result[0].intraModuleRefs).toBe(0);
    expect(result[0].crossModuleRefs).toBe(0);
    expect(result[0].commentOnlyRefs).toBe(0);
    expect(result[0].writeCount).toBe(0);
    expect(result[0].readCount).toBe(0);
    expect(result[0].isDynamicRef).toBe(false);
    expect(result[0].definition).toBe(def);
  });

  it('counts intra-module references correctly', () => {
    const def = makeDef({ name: 'Init', moduleName: 'ModA' });
    const table = buildTable([def]);
    const refMap = buildRefMap([
      makeRef({ symbolName: 'init', referencingModule: 'ModA', context: 'call' }),
      makeRef({ symbolName: 'init', referencingModule: 'ModA', context: 'call', lineNumber: 20 }),
    ]);
    const result = analyzeUsage(table, refMap);

    expect(result[0].totalReferences).toBe(2);
    expect(result[0].intraModuleRefs).toBe(2);
    expect(result[0].crossModuleRefs).toBe(0);
  });

  it('counts cross-module references correctly', () => {
    const def = makeDef({ name: 'Helper', moduleName: 'ModA' });
    const table = buildTable([def]);
    const refMap = buildRefMap([
      makeRef({ symbolName: 'helper', referencingModule: 'ModB', context: 'call' }),
      makeRef({ symbolName: 'helper', referencingModule: 'ModC', context: 'call' }),
    ]);
    const result = analyzeUsage(table, refMap);

    expect(result[0].totalReferences).toBe(2);
    expect(result[0].intraModuleRefs).toBe(0);
    expect(result[0].crossModuleRefs).toBe(2);
  });

  it('counts mixed intra and cross module references', () => {
    const def = makeDef({ name: 'Calc', moduleName: 'ModA' });
    const table = buildTable([def]);
    const refMap = buildRefMap([
      makeRef({ symbolName: 'calc', referencingModule: 'ModA', context: 'call' }),
      makeRef({ symbolName: 'calc', referencingModule: 'ModB', context: 'read' }),
    ]);
    const result = analyzeUsage(table, refMap);

    expect(result[0].totalReferences).toBe(2);
    expect(result[0].intraModuleRefs).toBe(1);
    expect(result[0].crossModuleRefs).toBe(1);
  });

  it('excludes comment-only references from totalReferences', () => {
    const def = makeDef({ name: 'OldFunc', moduleName: 'ModA' });
    const table = buildTable([def]);
    const refMap = buildRefMap([
      makeRef({ symbolName: 'oldfunc', referencingModule: 'ModA', isInComment: true }),
      makeRef({ symbolName: 'oldfunc', referencingModule: 'ModB', isInComment: true }),
    ]);
    const result = analyzeUsage(table, refMap);

    expect(result[0].totalReferences).toBe(0);
    expect(result[0].commentOnlyRefs).toBe(2);
  });

  it('counts comment refs separately from executable refs', () => {
    const def = makeDef({ name: 'Mixed', moduleName: 'ModA' });
    const table = buildTable([def]);
    const refMap = buildRefMap([
      makeRef({ symbolName: 'mixed', referencingModule: 'ModA', isInComment: true }),
      makeRef({ symbolName: 'mixed', referencingModule: 'ModB', context: 'call' }),
    ]);
    const result = analyzeUsage(table, refMap);

    expect(result[0].totalReferences).toBe(1);
    expect(result[0].crossModuleRefs).toBe(1);
    expect(result[0].commentOnlyRefs).toBe(1);
  });

  it('counts write and read references', () => {
    const def = makeDef({ name: 'Counter', moduleName: 'ModA', kind: 'Variable' });
    const table = buildTable([def]);
    const refMap = buildRefMap([
      makeRef({ symbolName: 'counter', referencingModule: 'ModA', context: 'write' }),
      makeRef({ symbolName: 'counter', referencingModule: 'ModA', context: 'write', lineNumber: 15 }),
      makeRef({ symbolName: 'counter', referencingModule: 'ModA', context: 'read', lineNumber: 20 }),
    ]);
    const result = analyzeUsage(table, refMap);

    expect(result[0].writeCount).toBe(2);
    expect(result[0].readCount).toBe(1);
    expect(result[0].totalReferences).toBe(3);
  });

  it('detects dynamic references', () => {
    const def = makeDef({ name: 'DynProc', moduleName: 'ModA' });
    const table = buildTable([def]);
    const refMap = buildRefMap([
      makeRef({ symbolName: 'dynproc', referencingModule: 'ModA', isDynamic: true }),
    ]);
    const result = analyzeUsage(table, refMap);

    expect(result[0].isDynamicRef).toBe(true);
    expect(result[0].totalReferences).toBe(1);
  });

  it('sets isDynamicRef false when no dynamic references exist', () => {
    const def = makeDef({ name: 'StaticProc', moduleName: 'ModA' });
    const table = buildTable([def]);
    const refMap = buildRefMap([
      makeRef({ symbolName: 'staticproc', referencingModule: 'ModA', isDynamic: false }),
    ]);
    const result = analyzeUsage(table, refMap);

    expect(result[0].isDynamicRef).toBe(false);
  });

  it('handles ambiguous references — same name in multiple modules', () => {
    const defA = makeDef({ name: 'Init', moduleName: 'ModA' });
    const defB = makeDef({ name: 'Init', moduleName: 'ModB' });
    const table = buildTable([defA, defB]);
    const refMap = buildRefMap([
      makeRef({ symbolName: 'init', referencingModule: 'ModC', context: 'call' }),
    ]);
    const result = analyzeUsage(table, refMap);

    // Both definitions should get the reference counted
    expect(result).toHaveLength(2);
    const usageA = result.find(u => u.definition.moduleName === 'ModA')!;
    const usageB = result.find(u => u.definition.moduleName === 'ModB')!;
    expect(usageA.totalReferences).toBe(1);
    expect(usageA.crossModuleRefs).toBe(1);
    expect(usageB.totalReferences).toBe(1);
    expect(usageB.crossModuleRefs).toBe(1);
  });

  it('handles case-insensitive matching via lowercase normalization', () => {
    const def = makeDef({ name: 'MyFunc', moduleName: 'ModA' });
    const table = buildTable([def]);
    // Reference uses different casing — but the reference map key is already lowercase
    const refMap = buildRefMap([
      makeRef({ symbolName: 'myfunc', referencingModule: 'ModB', context: 'call' }),
    ]);
    const result = analyzeUsage(table, refMap);

    expect(result[0].totalReferences).toBe(1);
  });

  it('does not count call/type-usage contexts as write or read', () => {
    const def = makeDef({ name: 'MyType', moduleName: 'ModA', kind: 'Type' });
    const table = buildTable([def]);
    const refMap = buildRefMap([
      makeRef({ symbolName: 'mytype', referencingModule: 'ModA', context: 'call' }),
      makeRef({ symbolName: 'mytype', referencingModule: 'ModB', context: 'type-usage' as any }),
    ]);
    const result = analyzeUsage(table, refMap);

    expect(result[0].totalReferences).toBe(2);
    expect(result[0].writeCount).toBe(0);
    expect(result[0].readCount).toBe(0);
  });

  it('handles multiple definitions across modules with mixed ref types', () => {
    const defA = makeDef({ name: 'Count', moduleName: 'ModA', kind: 'Variable' });
    const defB = makeDef({ name: 'Count', moduleName: 'ModB', kind: 'Variable' });
    const table = buildTable([defA, defB]);
    const refMap = buildRefMap([
      makeRef({ symbolName: 'count', referencingModule: 'ModA', context: 'write' }),
      makeRef({ symbolName: 'count', referencingModule: 'ModA', context: 'read', lineNumber: 15 }),
      makeRef({ symbolName: 'count', referencingModule: 'ModC', context: 'read', lineNumber: 5 }),
      makeRef({ symbolName: 'count', referencingModule: 'ModB', isInComment: true }),
    ]);
    const result = analyzeUsage(table, refMap);

    expect(result).toHaveLength(2);

    const usageA = result.find(u => u.definition.moduleName === 'ModA')!;
    expect(usageA.totalReferences).toBe(3);
    expect(usageA.intraModuleRefs).toBe(2);
    expect(usageA.crossModuleRefs).toBe(1);
    expect(usageA.commentOnlyRefs).toBe(1);
    expect(usageA.writeCount).toBe(1);
    expect(usageA.readCount).toBe(2);

    const usageB = result.find(u => u.definition.moduleName === 'ModB')!;
    expect(usageB.totalReferences).toBe(3);
    expect(usageB.intraModuleRefs).toBe(0);
    expect(usageB.crossModuleRefs).toBe(3);
    expect(usageB.commentOnlyRefs).toBe(1);
  });
});
