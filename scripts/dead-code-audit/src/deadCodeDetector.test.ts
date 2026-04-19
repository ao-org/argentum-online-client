import { describe, it, expect } from 'vitest';
import {
  detectUnusedSymbols,
  detectUnreachableCode,
  detectCommentedOutBlocks,
} from './deadCodeDetector.js';
import type {
  SymbolUsage,
  SymbolDefinition,
  ParsedModule,
  ParsedLine,
  SourceFile,
} from './types.js';

// ?? Helpers ?????????????????????????????????????????????????????????????

function makeDef(overrides: Partial<SymbolDefinition> = {}): SymbolDefinition {
  return {
    id: 'TestModule::TestSym::1',
    name: 'TestSym',
    kind: 'Sub',
    visibility: 'Public',
    moduleName: 'TestModule',
    filePath: 'test.bas',
    lineNumber: 1,
    scope: 'module',
    isEventHandler: false,
    ...overrides,
  };
}

function makeUsage(
  defOverrides: Partial<SymbolDefinition> = {},
  usageOverrides: Partial<Omit<SymbolUsage, 'definition'>> = {},
): SymbolUsage {
  return {
    definition: makeDef(defOverrides),
    totalReferences: 0,
    intraModuleRefs: 0,
    crossModuleRefs: 0,
    commentOnlyRefs: 0,
    writeCount: 0,
    readCount: 0,
    isDynamicRef: false,
    ...usageOverrides,
  };
}

function makeLine(lineNumber: number, text: string, opts: Partial<ParsedLine> = {}): ParsedLine {
  const trimmed = text.trim();
  return {
    lineNumber,
    text,
    isComment: trimmed.startsWith("'"),
    isPreprocessor: false,
    isExecutable: !trimmed.startsWith("'") && trimmed.length > 0,
    originalLines: [lineNumber],
    ...opts,
  };
}

function makeModule(lines: ParsedLine[], filePath = 'test.bas'): ParsedModule {
  const source: SourceFile = {
    path: filePath,
    type: 'bas',
    moduleName: 'TestModule',
    content: '',
  };
  return { source, lines, attributeLines: [] };
}

// ?? detectUnusedSymbols ?????????????????????????????????????????????????

describe('detectUnusedSymbols', () => {
  it('reports unused Sub as unused-procedure', () => {
    const usages = [makeUsage({ kind: 'Sub', name: 'DoStuff' })];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('unused-procedure');
    expect(findings[0].symbolName).toBe('DoStuff');
    expect(findings[0].confidence).toBe('confirmed');
    expect(findings[0].removable).toBe(true);
  });

  it('reports unused Function as unused-procedure', () => {
    const usages = [makeUsage({ kind: 'Function', name: 'CalcValue' })];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('unused-procedure');
  });

  it('reports unused Variable as unused-variable', () => {
    const usages = [makeUsage({ kind: 'Variable', name: 'myVar' })];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('unused-variable');
  });

  it('reports unused Const as unused-const', () => {
    const usages = [makeUsage({ kind: 'Const', name: 'MAX_VAL' })];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('unused-const');
  });

  it('reports unused Enum as unused-enum', () => {
    const usages = [makeUsage({ kind: 'Enum', name: 'Colors' })];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('unused-enum');
  });

  it('reports unused Type as unused-type', () => {
    const usages = [makeUsage({ kind: 'Type', name: 'Position' })];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('unused-type');
  });

  it('reports unused Declare as unused-declare', () => {
    const usages = [makeUsage({ kind: 'Declare', name: 'SendMessage' })];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('unused-declare');
  });

  it('skips event handlers', () => {
    const usages = [makeUsage({ isEventHandler: true, name: 'Form_Load' })];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(0);
  });

  it('classifies write-only variables', () => {
    const usages = [
      makeUsage(
        { kind: 'Variable', name: 'counter' },
        { writeCount: 3, readCount: 0, totalReferences: 3 },
      ),
    ];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('write-only-variable');
    expect(findings[0].reason).toContain('assigned but never read');
  });

  it('does not flag variables with both reads and writes', () => {
    const usages = [
      makeUsage(
        { kind: 'Variable', name: 'x' },
        { writeCount: 2, readCount: 1, totalReferences: 3 },
      ),
    ];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(0);
  });

  it('sets review-needed for dynamic references', () => {
    const usages = [makeUsage({ name: 'DynProc' }, { isDynamicRef: true })];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(1);
    expect(findings[0].confidence).toBe('review-needed');
    expect(findings[0].removable).toBe(false);
  });

  it('does not report symbols with references', () => {
    const usages = [makeUsage({}, { totalReferences: 5, readCount: 3, writeCount: 2 })];
    const findings = detectUnusedSymbols(usages);
    expect(findings).toHaveLength(0);
  });

  it('generates correct finding id', () => {
    const usages = [makeUsage({ filePath: 'Module1.bas', lineNumber: 42, kind: 'Sub' })];
    const findings = detectUnusedSymbols(usages);
    expect(findings[0].id).toBe('finding-unused-procedure-Module1.bas-42');
  });

  it('uses endLineNumber when available', () => {
    const usages = [makeUsage({ lineNumber: 10, endLineNumber: 25, kind: 'Sub' })];
    const findings = detectUnusedSymbols(usages);
    expect(findings[0].startLine).toBe(10);
    expect(findings[0].endLine).toBe(25);
  });
});

// ?? detectUnreachableCode ???????????????????????????????????????????????

describe('detectUnreachableCode', () => {
  it('detects code after Exit Sub', () => {
    const lines = [
      makeLine(1, 'Private Sub Foo()'),
      makeLine(2, '  x = 1'),
      makeLine(3, '  Exit Sub'),
      makeLine(4, '  y = 2'),
      makeLine(5, 'End Sub'),
    ];
    const findings = detectUnreachableCode([makeModule(lines)]);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('unreachable-code');
    expect(findings[0].startLine).toBe(4);
    expect(findings[0].endLine).toBe(4);
  });

  it('detects code after Exit Function', () => {
    const lines = [
      makeLine(1, 'Public Function Bar() As Long'),
      makeLine(2, '  Bar = 42'),
      makeLine(3, '  Exit Function'),
      makeLine(4, '  Bar = 0'),
      makeLine(5, '  x = 1'),
      makeLine(6, 'End Function'),
    ];
    const findings = detectUnreachableCode([makeModule(lines)]);
    expect(findings).toHaveLength(1);
    expect(findings[0].startLine).toBe(4);
    expect(findings[0].endLine).toBe(5);
  });

  it('detects code after unconditional GoTo', () => {
    const lines = [
      makeLine(1, 'Private Sub Test()'),
      makeLine(2, '  GoTo Done'),
      makeLine(3, '  x = 1'),
      makeLine(4, 'Done:'),
      makeLine(5, 'End Sub'),
    ];
    const findings = detectUnreachableCode([makeModule(lines)]);
    expect(findings).toHaveLength(1);
    expect(findings[0].startLine).toBe(3);
    expect(findings[0].endLine).toBe(3);
  });

  it('resets unreachable state at labels', () => {
    const lines = [
      makeLine(1, 'Private Sub Test()'),
      makeLine(2, '  Exit Sub'),
      makeLine(3, '  x = 1'),
      makeLine(4, 'Handler:'),
      makeLine(5, '  y = 2'),
      makeLine(6, 'End Sub'),
    ];
    const findings = detectUnreachableCode([makeModule(lines)]);
    expect(findings).toHaveLength(1);
    expect(findings[0].startLine).toBe(3);
    expect(findings[0].endLine).toBe(3);
  });

  it('detects If False Then dead branches', () => {
    const lines = [
      makeLine(1, 'Private Sub Test()'),
      makeLine(2, 'If False Then'),
      makeLine(3, '  x = 1'),
      makeLine(4, 'End If'),
      makeLine(5, 'End Sub'),
    ];
    const findings = detectUnreachableCode([makeModule(lines)]);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('dead-branch');
    expect(findings[0].startLine).toBe(2);
    expect(findings[0].endLine).toBe(4);
  });

  it('detects If 0 Then dead branches', () => {
    const lines = [
      makeLine(1, 'Private Sub Test()'),
      makeLine(2, 'If 0 Then'),
      makeLine(3, '  x = 1'),
      makeLine(4, 'End If'),
      makeLine(5, 'End Sub'),
    ];
    const findings = detectUnreachableCode([makeModule(lines)]);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('dead-branch');
  });

  it('does not flag code inside procedures without exits', () => {
    const lines = [
      makeLine(1, 'Private Sub Normal()'),
      makeLine(2, '  x = 1'),
      makeLine(3, '  y = 2'),
      makeLine(4, 'End Sub'),
    ];
    const findings = detectUnreachableCode([makeModule(lines)]);
    expect(findings).toHaveLength(0);
  });

  it('handles multiple procedures in one module', () => {
    const lines = [
      makeLine(1, 'Private Sub A()'),
      makeLine(2, '  Exit Sub'),
      makeLine(3, '  dead = 1'),
      makeLine(4, 'End Sub'),
      makeLine(5, 'Private Sub B()'),
      makeLine(6, '  alive = 1'),
      makeLine(7, 'End Sub'),
    ];
    const findings = detectUnreachableCode([makeModule(lines)]);
    expect(findings).toHaveLength(1);
    expect(findings[0].startLine).toBe(3);
  });
});

// ?? detectCommentedOutBlocks ????????????????????????????????????????????

describe('detectCommentedOutBlocks', () => {
  it('detects 6+ consecutive comment lines with VB6 keywords', () => {
    const lines = [
      makeLine(1, "' Sub OldProc()"),
      makeLine(2, "'   Dim x As Long"),
      makeLine(3, "'   x = 1"),
      makeLine(4, "'   If x > 0 Then"),
      makeLine(5, "'     Call DoStuff"),
      makeLine(6, "'   End If"),
      makeLine(7, "' End Sub"),
    ];
    const findings = detectCommentedOutBlocks([makeModule(lines)]);
    expect(findings).toHaveLength(1);
    expect(findings[0].category).toBe('commented-out-block');
    expect(findings[0].startLine).toBe(1);
    expect(findings[0].endLine).toBe(7);
    expect(findings[0].reason).toContain('Commented-out code candidate');
  });

  it('does not flag blocks shorter than 6 lines', () => {
    const lines = [
      makeLine(1, "' Sub OldProc()"),
      makeLine(2, "'   Dim x As Long"),
      makeLine(3, "'   x = 1"),
      makeLine(4, "'   Call DoStuff"),
      makeLine(5, "' End Sub"),
    ];
    const findings = detectCommentedOutBlocks([makeModule(lines)]);
    expect(findings).toHaveLength(0);
  });

  it('does not flag comment blocks without VB6 keywords', () => {
    const lines = [
      makeLine(1, "' This is a documentation comment"),
      makeLine(2, "' that describes the module"),
      makeLine(3, "' and its purpose in the system"),
      makeLine(4, "' Author: John"),
      makeLine(5, "' Date: 2024-01-01"),
      makeLine(6, "' Version: 1.0"),
      makeLine(7, "' License: MIT"),
    ];
    const findings = detectCommentedOutBlocks([makeModule(lines)]);
    expect(findings).toHaveLength(0);
  });

  it('handles multiple comment blocks in one module', () => {
    const lines = [
      makeLine(1, "' Sub A()"),
      makeLine(2, "'   Dim x As Long"),
      makeLine(3, "'   x = 1"),
      makeLine(4, "'   If x > 0 Then"),
      makeLine(5, "'     Call DoStuff"),
      makeLine(6, "'   End If"),
      makeLine(7, 'x = 1'),
      makeLine(8, "' Sub B()"),
      makeLine(9, "'   Dim y As Long"),
      makeLine(10, "'   y = 2"),
      makeLine(11, "'   For i = 1 To 10"),
      makeLine(12, "'     Call Other"),
      makeLine(13, "'   Next i"),
    ];
    const findings = detectCommentedOutBlocks([makeModule(lines)]);
    expect(findings).toHaveLength(2);
  });

  it('breaks blocks on non-comment lines', () => {
    const lines = [
      makeLine(1, "' Sub A()"),
      makeLine(2, "'   Dim x As Long"),
      makeLine(3, "'   x = 1"),
      makeLine(4, 'y = 2'),
      makeLine(5, "' Sub B()"),
      makeLine(6, "'   Dim z As Long"),
      makeLine(7, "'   z = 3"),
    ];
    // Each block is only 3 lines, below threshold
    const findings = detectCommentedOutBlocks([makeModule(lines)]);
    expect(findings).toHaveLength(0);
  });
});
