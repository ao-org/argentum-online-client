import { describe, it, expect } from 'vitest';
import { tokenizeLine, classifyContext, detectCallByName, scanReferences } from './referenceScanner.js';
import { parseModule } from './parser.js';
import { extractSymbols, buildSymbolTable } from './symbolExtractor.js';
import type { SourceFile, ParsedModule, SymbolTable, SymbolDefinition } from './types.js';

// ?? Helper to build a module + symbol table from VB6 source ??????????????????

function buildTestContext(sources: Array<{ name: string; content: string; type?: 'bas' | 'cls' | 'frm' }>) {
  const modules: ParsedModule[] = [];
  const allDefs: SymbolDefinition[] = [];

  for (const src of sources) {
    const sourceFile: SourceFile = {
      path: `${src.name}.${src.type ?? 'bas'}`,
      type: src.type ?? 'bas',
      moduleName: src.name,
      content: src.content,
    };
    const mod = parseModule(sourceFile);
    modules.push(mod);
    allDefs.push(...extractSymbols(mod));
  }

  const symbolTable = buildSymbolTable(allDefs);
  return { modules, symbolTable };
}

// ?? tokenizeLine ?????????????????????????????????????????????????????????????

describe('tokenizeLine', () => {
  it('splits a simple assignment into tokens', () => {
    const tokens = tokenizeLine('x = y + z');
    expect(tokens).toContain('x');
    expect(tokens).toContain('y');
    expect(tokens).toContain('z');
  });

  it('handles function calls with parentheses', () => {
    const tokens = tokenizeLine('result = Foo(a, b)');
    expect(tokens).toContain('result');
    expect(tokens).toContain('Foo');
    expect(tokens).toContain('a');
    expect(tokens).toContain('b');
  });

  it('handles module-qualified references', () => {
    const tokens = tokenizeLine('x = MyModule.MyFunc(a)');
    expect(tokens).toContain('MyModule.MyFunc');
    expect(tokens).toContain('MyModule');
    expect(tokens).toContain('MyFunc');
  });

  it('handles VB6 operators and delimiters', () => {
    const tokens = tokenizeLine('If x > 5 And y < 10 Then');
    expect(tokens).toContain('If');
    expect(tokens).toContain('x');
    expect(tokens).toContain('5');
    expect(tokens).toContain('And');
    expect(tokens).toContain('y');
    expect(tokens).toContain('10');
    expect(tokens).toContain('Then');
  });
});

// ?? classifyContext ??????????????????????????????????????????????????????????

describe('classifyContext', () => {
  it('classifies type-usage after As keyword', () => {
    const ctx = classifyContext('MyType', 'Dim x As MyType', undefined);
    expect(ctx).toBe('type-usage');
  });

  it('classifies type-usage after As New', () => {
    const ctx = classifyContext('Collection', 'Dim x As New Collection', undefined);
    expect(ctx).toBe('type-usage');
  });

  it('classifies write for simple assignment', () => {
    const ctx = classifyContext('x', 'x = 5', undefined);
    expect(ctx).toBe('write');
  });

  it('classifies write for Set assignment', () => {
    const ctx = classifyContext('obj', 'Set obj = Nothing', undefined);
    expect(ctx).toBe('write');
  });

  it('classifies call for explicit Call statement', () => {
    const subDef: SymbolDefinition[] = [{
      id: 'M::DoStuff::1', name: 'DoStuff', kind: 'Sub', visibility: 'Public',
      moduleName: 'M', filePath: 'M.bas', lineNumber: 1, scope: 'module', isEventHandler: false,
    }];
    const ctx = classifyContext('DoStuff', 'Call DoStuff', subDef);
    expect(ctx).toBe('call');
  });

  it('classifies call for Sub at start of line', () => {
    const subDef: SymbolDefinition[] = [{
      id: 'M::Init::1', name: 'Init', kind: 'Sub', visibility: 'Public',
      moduleName: 'M', filePath: 'M.bas', lineNumber: 1, scope: 'module', isEventHandler: false,
    }];
    const ctx = classifyContext('Init', 'Init', subDef);
    expect(ctx).toBe('call');
  });

  it('classifies call for Function with parens', () => {
    const funcDef: SymbolDefinition[] = [{
      id: 'M::GetVal::1', name: 'GetVal', kind: 'Function', visibility: 'Public',
      moduleName: 'M', filePath: 'M.bas', lineNumber: 1, scope: 'module', isEventHandler: false,
    }];
    const ctx = classifyContext('GetVal', 'x = GetVal(1)', funcDef);
    expect(ctx).toBe('call');
  });

  it('classifies read for variable on right side of assignment', () => {
    const ctx = classifyContext('y', 'x = y + 1', undefined);
    expect(ctx).toBe('read');
  });
});

// ?? detectCallByName ?????????????????????????????????????????????????????????

describe('detectCallByName', () => {
  it('detects CallByName pattern', () => {
    const content = 'CallByName(obj, "MyProc", VbMethod)';
    const names = detectCallByName(content);
    expect(names).toEqual(['MyProc']);
  });

  it('detects multiple CallByName patterns', () => {
    const content = [
      'CallByName(obj1, "ProcA", VbMethod)',
      'CallByName(obj2, "ProcB", VbGet)',
    ].join('\n');
    const names = detectCallByName(content);
    expect(names).toEqual(['ProcA', 'ProcB']);
  });

  it('returns empty for no CallByName', () => {
    const content = 'x = 5\nCall DoStuff';
    const names = detectCallByName(content);
    expect(names).toEqual([]);
  });
});


// ?? scanReferences integration tests ?????????????????????????????????????????

describe('scanReferences', () => {
  it('detects a simple call reference', () => {
    const { modules, symbolTable } = buildTestContext([{
      name: 'TestMod',
      content: [
        'Public Sub Init()',
        'End Sub',
        '',
        'Public Sub Main()',
        '  Call Init',
        'End Sub',
      ].join('\n'),
    }]);

    const refMap = scanReferences(modules, symbolTable);
    const initRefs = refMap.references.get('init') ?? [];
    expect(initRefs.length).toBe(1);
    expect(initRefs[0].context).toBe('call');
    expect(initRefs[0].lineNumber).toBe(5);
    expect(initRefs[0].isInComment).toBe(false);
    expect(initRefs[0].isDynamic).toBe(false);
  });

  it('does not count a symbol declaration as a reference', () => {
    const { modules, symbolTable } = buildTestContext([{
      name: 'TestMod',
      content: [
        'Public Sub Unused()',
        'End Sub',
      ].join('\n'),
    }]);

    const refMap = scanReferences(modules, symbolTable);
    const refs = refMap.references.get('unused') ?? [];
    expect(refs.length).toBe(0);
  });

  it('tracks references in comments with isInComment=true', () => {
    const { modules, symbolTable } = buildTestContext([{
      name: 'TestMod',
      content: [
        'Public Sub DoWork()',
        'End Sub',
        '',
        "' DoWork is no longer needed",
      ].join('\n'),
    }]);

    const refMap = scanReferences(modules, symbolTable);
    const refs = refMap.references.get('dowork') ?? [];
    expect(refs.length).toBe(1);
    expect(refs[0].isInComment).toBe(true);
  });

  it('classifies variable write and read correctly', () => {
    const { modules, symbolTable } = buildTestContext([{
      name: 'TestMod',
      content: [
        'Dim myVar As Long',
        'Public Sub Test()',
        '  myVar = 10',
        '  Dim x As Long',
        '  x = myVar + 1',
        'End Sub',
      ].join('\n'),
    }]);

    const refMap = scanReferences(modules, symbolTable);
    const refs = refMap.references.get('myvar') ?? [];
    // Should have 2 refs: one write (line 3) and one read (line 5)
    expect(refs.length).toBe(2);
    const writeRef = refs.find(r => r.context === 'write');
    const readRef = refs.find(r => r.context === 'read');
    expect(writeRef).toBeDefined();
    expect(readRef).toBeDefined();
    expect(writeRef!.lineNumber).toBe(3);
    expect(readRef!.lineNumber).toBe(5);
  });

  it('classifies type-usage for As keyword', () => {
    const { modules, symbolTable } = buildTestContext([{
      name: 'TestMod',
      content: [
        'Public Type MyRecord',
        '  x As Long',
        'End Type',
        '',
        'Dim rec As MyRecord',
      ].join('\n'),
    }]);

    const refMap = scanReferences(modules, symbolTable);
    const refs = refMap.references.get('myrecord') ?? [];
    expect(refs.length).toBe(1);
    expect(refs[0].context).toBe('type-usage');
  });

  it('detects CallByName dynamic references', () => {
    const { modules, symbolTable } = buildTestContext([{
      name: 'TestMod',
      content: [
        'Public Sub TargetProc()',
        'End Sub',
        '',
        'Public Sub Caller()',
        '  CallByName(Me, "TargetProc", VbMethod)',
        'End Sub',
      ].join('\n'),
    }]);

    const refMap = scanReferences(modules, symbolTable);
    const refs = refMap.references.get('targetproc') ?? [];
    // Should have at least one dynamic reference
    const dynamicRef = refs.find(r => r.isDynamic);
    expect(dynamicRef).toBeDefined();
    expect(dynamicRef!.context).toBe('call');
  });

  it('handles module-qualified references', () => {
    const { modules, symbolTable } = buildTestContext([
      {
        name: 'UtilMod',
        content: [
          'Public Function Helper() As Long',
          '  Helper = 42',
          'End Function',
        ].join('\n'),
      },
      {
        name: 'MainMod',
        content: [
          'Public Sub Main()',
          '  Dim x As Long',
          '  x = UtilMod.Helper()',
          'End Sub',
        ].join('\n'),
      },
    ]);

    const refMap = scanReferences(modules, symbolTable);
    const refs = refMap.references.get('helper') ?? [];
    // Should have a reference from MainMod (qualified) + self-ref from line 2 (Helper = 42)
    const mainRef = refs.find(r => r.referencingModule === 'MainMod');
    expect(mainRef).toBeDefined();
  });

  it('handles cross-module references', () => {
    const { modules, symbolTable } = buildTestContext([
      {
        name: 'ModA',
        content: [
          'Public Sub SharedProc()',
          'End Sub',
        ].join('\n'),
      },
      {
        name: 'ModB',
        content: [
          'Public Sub Caller()',
          '  SharedProc',
          'End Sub',
        ].join('\n'),
      },
    ]);

    const refMap = scanReferences(modules, symbolTable);
    const refs = refMap.references.get('sharedproc') ?? [];
    expect(refs.length).toBe(1);
    expect(refs[0].referencingModule).toBe('ModB');
    expect(refs[0].context).toBe('call');
  });

  it('does not count variable declaration as a reference', () => {
    const { modules, symbolTable } = buildTestContext([{
      name: 'TestMod',
      content: [
        'Dim unusedVar As Long',
      ].join('\n'),
    }]);

    const refMap = scanReferences(modules, symbolTable);
    const refs = refMap.references.get('unusedvar') ?? [];
    expect(refs.length).toBe(0);
  });

  it('handles multiple references to the same symbol on different lines', () => {
    const { modules, symbolTable } = buildTestContext([{
      name: 'TestMod',
      content: [
        'Public Const MAX_VAL As Long = 100',
        '',
        'Public Sub Test()',
        '  If x > MAX_VAL Then',
        '    y = MAX_VAL',
        '  End If',
        'End Sub',
      ].join('\n'),
    }]);

    const refMap = scanReferences(modules, symbolTable);
    const refs = refMap.references.get('max_val') ?? [];
    expect(refs.length).toBe(2);
  });
});
