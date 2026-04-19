import { describe, it, expect } from 'vitest';
import { extractSymbols, buildSymbolTable } from './symbolExtractor.js';
import { parseModule } from './parser.js';
import type { SourceFile, SymbolDefinition } from './types.js';

function makeSource(content: string, type: 'bas' | 'cls' | 'frm' = 'bas', moduleName = 'TestMod'): SourceFile {
  return { path: `test.${type}`, type, moduleName, content };
}

function extract(content: string, type: 'bas' | 'cls' | 'frm' = 'bas', moduleName = 'TestMod'): SymbolDefinition[] {
  const src = makeSource(content, type, moduleName);
  const parsed = parseModule(src);
  return extractSymbols(parsed);
}

function findByName(defs: SymbolDefinition[], name: string): SymbolDefinition | undefined {
  return defs.find(d => d.name.toLowerCase() === name.toLowerCase());
}

// ?? Sub / Function extraction ????????????????????????????????????????????????

describe('extractSymbols — Sub/Function', () => {
  it('extracts a Public Sub', () => {
    const defs = extract('Public Sub Init()\nEnd Sub');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('Init');
    expect(defs[0].kind).toBe('Sub');
    expect(defs[0].visibility).toBe('Public');
    expect(defs[0].endLineNumber).toBe(2);
  });

  it('extracts a Private Function with return type', () => {
    const defs = extract('Private Function GetValue() As Long\nEnd Function');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('GetValue');
    expect(defs[0].kind).toBe('Function');
    expect(defs[0].visibility).toBe('Private');
    expect(defs[0].dataType).toBe('Long');
    expect(defs[0].endLineNumber).toBe(2);
  });

  it('defaults to Public visibility for Sub/Function in .bas files', () => {
    const defs = extract('Sub DoStuff()\nEnd Sub', 'bas');
    expect(defs[0].visibility).toBe('Public');
  });

  it('defaults to Private visibility for Sub/Function in .cls files', () => {
    const defs = extract('Sub DoStuff()\nEnd Sub', 'cls');
    expect(defs[0].visibility).toBe('Private');
  });

  it('extracts Friend Sub', () => {
    const defs = extract('Friend Sub Helper()\nEnd Sub');
    expect(defs).toHaveLength(1);
    expect(defs[0].visibility).toBe('Friend');
  });

  it('handles Static modifier', () => {
    const defs = extract('Private Static Function Calc() As Long\nEnd Function');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('Calc');
    expect(defs[0].kind).toBe('Function');
  });

  it('generates correct id', () => {
    const defs = extract('Public Sub Init()\nEnd Sub');
    expect(defs[0].id).toBe('TestMod::Init::1');
  });
});

// ?? Property Get/Let/Set ?????????????????????????????????????????????????????

describe('extractSymbols — Property', () => {
  it('extracts Property Get', () => {
    const defs = extract('Public Property Get Name() As String\nEnd Property');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('Name');
    expect(defs[0].kind).toBe('Property');
    expect(defs[0].endLineNumber).toBe(2);
  });

  it('extracts Property Let', () => {
    const defs = extract('Property Let Name(val As String)\nEnd Property');
    expect(defs).toHaveLength(1);
    expect(defs[0].kind).toBe('Property');
  });

  it('extracts Property Set', () => {
    const defs = extract('Friend Property Set Obj(val As Object)\nEnd Property');
    expect(defs).toHaveLength(1);
    expect(defs[0].visibility).toBe('Friend');
  });
});

// ?? Module-level variables ???????????????????????????????????????????????????

describe('extractSymbols — Variables', () => {
  it('extracts Public variable', () => {
    const defs = extract('Public bSkins As Boolean');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('bSkins');
    expect(defs[0].kind).toBe('Variable');
    expect(defs[0].visibility).toBe('Public');
    expect(defs[0].dataType).toBe('Boolean');
    expect(defs[0].scope).toBe('module');
  });

  it('extracts Private variable', () => {
    const defs = extract('Private mCount As Long');
    expect(defs[0].visibility).toBe('Private');
  });

  it('extracts Dim at module level as Private', () => {
    const defs = extract('Dim gData As String');
    expect(defs[0].visibility).toBe('Private');
    expect(defs[0].scope).toBe('module');
  });

  it('extracts multiple variables on one line', () => {
    const defs = extract('Dim x As Long, y As String');
    expect(defs).toHaveLength(2);
    expect(defs[0].name).toBe('x');
    expect(defs[0].dataType).toBe('Long');
    expect(defs[1].name).toBe('y');
    expect(defs[1].dataType).toBe('String');
  });

  it('extracts array variable with bounds', () => {
    const defs = extract('Public NpcWorlds(1 To 2000) As Byte');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('NpcWorlds');
    expect(defs[0].dataType).toBe('Byte');
  });

  it('extracts Global variable as Public', () => {
    const defs = extract('Global SomeVar As Integer');
    expect(defs).toHaveLength(1);
    expect(defs[0].visibility).toBe('Public');
  });

  it('handles WithEvents keyword', () => {
    const defs = extract('Public WithEvents tmrTimer As Timer');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('tmrTimer');
    expect(defs[0].dataType).toBe('Timer');
  });

  it('extracts variable without explicit type', () => {
    const defs = extract('Dim x');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('x');
    expect(defs[0].dataType).toBeUndefined();
  });
});

// ?? Local variables ??????????????????????????????????????????????????????????

describe('extractSymbols — Local variables', () => {
  it('extracts Dim inside a procedure as local', () => {
    const defs = extract('Public Sub Test()\nDim i As Long\nEnd Sub');
    const localVar = findByName(defs, 'i');
    expect(localVar).toBeDefined();
    expect(localVar!.scope).toBe('local');
    expect(localVar!.parentProcedure).toBe('Test');
    expect(localVar!.visibility).toBe('Private');
  });

  it('extracts multiple local variables', () => {
    const defs = extract('Private Sub Calc()\nDim a As Long, b As String\nEnd Sub');
    const locals = defs.filter(d => d.scope === 'local');
    expect(locals).toHaveLength(2);
    expect(locals[0].parentProcedure).toBe('Calc');
    expect(locals[1].parentProcedure).toBe('Calc');
  });
});

// ?? Const ????????????????????????????????????????????????????????????????????

describe('extractSymbols — Const', () => {
  it('extracts Public Const with type', () => {
    const defs = extract('Public Const NO_WEAPON As Byte = 2');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('NO_WEAPON');
    expect(defs[0].kind).toBe('Const');
    expect(defs[0].visibility).toBe('Public');
    expect(defs[0].dataType).toBe('Byte');
  });

  it('extracts Private Const without type', () => {
    const defs = extract('Private Const MAX_ITEMS = 100');
    expect(defs[0].visibility).toBe('Private');
    expect(defs[0].dataType).toBeUndefined();
  });

  it('extracts Const without visibility as Public', () => {
    const defs = extract('Const MY_VAL As Long = 5');
    expect(defs[0].visibility).toBe('Public');
  });

  it('extracts local Const inside procedure', () => {
    const defs = extract('Public Sub Test()\nConst LOCAL_VAL = 10\nEnd Sub');
    const c = findByName(defs, 'LOCAL_VAL');
    expect(c).toBeDefined();
    expect(c!.scope).toBe('local');
    expect(c!.parentProcedure).toBe('Test');
  });
});

// ?? Enum ?????????????????????????????????????????????????????????????????????

describe('extractSymbols — Enum', () => {
  it('extracts Public Enum with endLineNumber', () => {
    const defs = extract('Public Enum tMacro\n  dobleclick = 1\n  Coordenadas = 2\nEnd Enum');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('tMacro');
    expect(defs[0].kind).toBe('Enum');
    expect(defs[0].visibility).toBe('Public');
    expect(defs[0].endLineNumber).toBe(4);
  });

  it('extracts Private Enum', () => {
    const defs = extract('Private Enum eDirection\n  North\n  South\nEnd Enum');
    expect(defs[0].visibility).toBe('Private');
  });

  it('does not extract enum members as separate symbols', () => {
    const defs = extract('Public Enum tMacro\n  dobleclick = 1\n  Coordenadas = 2\nEnd Enum');
    expect(defs).toHaveLength(1);
  });
});

// ?? Type (UDT) ???????????????????????????????????????????????????????????????

describe('extractSymbols — Type', () => {
  it('extracts Private Type with endLineNumber', () => {
    const defs = extract('Private Type Position\n  X As Long\n  Y As Long\nEnd Type');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('Position');
    expect(defs[0].kind).toBe('Type');
    expect(defs[0].visibility).toBe('Private');
    expect(defs[0].endLineNumber).toBe(4);
  });

  it('does not extract type members as separate symbols', () => {
    const defs = extract('Public Type tUser\n  Name As String\n  Level As Long\nEnd Type');
    expect(defs).toHaveLength(1);
  });
});

// ?? API Declare ??????????????????????????????????????????????????????????????

describe('extractSymbols — Declare', () => {
  it('extracts Private Declare Function', () => {
    const defs = extract('Private Declare Function SendMessage Lib "user32" (ByVal hWnd As Long) As Long');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('SendMessage');
    expect(defs[0].kind).toBe('Declare');
    expect(defs[0].visibility).toBe('Private');
  });

  it('extracts Public Declare Sub', () => {
    const defs = extract('Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)');
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('Sleep');
    expect(defs[0].kind).toBe('Declare');
  });

  it('extracts Declare without visibility', () => {
    const defs = extract('Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long) As Long');
    expect(defs).toHaveLength(1);
    expect(defs[0].visibility).toBe('Public');
  });
});

// ?? Event handlers ???????????????????????????????????????????????????????????

describe('extractSymbols — Event handlers', () => {
  it('detects Form_Load as event handler', () => {
    const defs = extract('Private Sub Form_Load()\nEnd Sub');
    expect(defs[0].isEventHandler).toBe(true);
    expect(defs[0].kind).toBe('Sub');
  });

  it('detects cmdMas_Click as event handler', () => {
    const defs = extract('Private Sub cmdMas_Click()\nEnd Sub');
    expect(defs[0].isEventHandler).toBe(true);
  });

  it('detects Timer1_Timer as event handler', () => {
    const defs = extract('Private Sub Timer1_Timer()\nEnd Sub');
    expect(defs[0].isEventHandler).toBe(true);
  });

  it('detects txtNombre_Change as event handler', () => {
    const defs = extract('Private Sub txtNombre_Change()\nEnd Sub');
    expect(defs[0].isEventHandler).toBe(true);
  });

  it('detects OpcionImg_MouseMove as event handler', () => {
    const defs = extract('Private Sub OpcionImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)\nEnd Sub');
    expect(defs[0].isEventHandler).toBe(true);
  });

  it('does not mark regular Sub with underscore as event handler', () => {
    const defs = extract('Private Sub Do_Something()\nEnd Sub');
    expect(defs[0].isEventHandler).toBe(false);
  });

  it('does not mark Sub without underscore as event handler', () => {
    const defs = extract('Public Sub Init()\nEnd Sub');
    expect(defs[0].isEventHandler).toBe(false);
  });
});


// ?? Scope tracking ???????????????????????????????????????????????????????????

describe('extractSymbols — Scope tracking', () => {
  it('tracks module-level vs local scope correctly', () => {
    const code = [
      'Public gCount As Long',
      'Public Sub Process()',
      'Dim localVar As String',
      'End Sub',
      'Private mData As Integer',
    ].join('\n');
    const defs = extract(code);
    const gCount = findByName(defs, 'gCount');
    const localVar = findByName(defs, 'localVar');
    const mData = findByName(defs, 'mData');

    expect(gCount!.scope).toBe('module');
    expect(localVar!.scope).toBe('local');
    expect(localVar!.parentProcedure).toBe('Process');
    expect(mData!.scope).toBe('module');
  });

  it('resets scope after End Sub', () => {
    const code = [
      'Public Sub First()',
      'Dim a As Long',
      'End Sub',
      'Dim moduleVar As Long',
      'Public Sub Second()',
      'Dim b As Long',
      'End Sub',
    ].join('\n');
    const defs = extract(code);
    const a = findByName(defs, 'a');
    const moduleVar = findByName(defs, 'moduleVar');
    const b = findByName(defs, 'b');

    expect(a!.scope).toBe('local');
    expect(a!.parentProcedure).toBe('First');
    expect(moduleVar!.scope).toBe('module');
    expect(moduleVar!.parentProcedure).toBeUndefined();
    expect(b!.scope).toBe('local');
    expect(b!.parentProcedure).toBe('Second');
  });
});

// ?? Multi-line block endLineNumber ???????????????????????????????????????????

describe('extractSymbols — endLineNumber tracking', () => {
  it('sets endLineNumber for Sub', () => {
    const defs = extract('Public Sub Test()\nDim x As Long\nx = 1\nEnd Sub');
    expect(defs[0].endLineNumber).toBe(4);
  });

  it('sets endLineNumber for Function', () => {
    const defs = extract('Private Function Calc() As Long\nCalc = 42\nEnd Function');
    expect(defs[0].endLineNumber).toBe(3);
  });

  it('sets endLineNumber for Enum', () => {
    const defs = extract('Public Enum Colors\n  Red\n  Green\n  Blue\nEnd Enum');
    expect(defs[0].endLineNumber).toBe(5);
  });

  it('sets endLineNumber for Type', () => {
    const defs = extract('Private Type Point\n  X As Long\n  Y As Long\nEnd Type');
    expect(defs[0].endLineNumber).toBe(4);
  });
});

// ?? buildSymbolTable ?????????????????????????????????????????????????????????

describe('buildSymbolTable', () => {
  it('indexes symbols by lowercase name', () => {
    const defs = extract('Public Sub Init()\nEnd Sub\nPublic gCount As Long');
    const table = buildSymbolTable(defs);

    expect(table.symbols.get('init')).toHaveLength(1);
    expect(table.symbols.get('gcount')).toHaveLength(1);
  });

  it('groups same-name symbols from different modules', () => {
    const defs1 = extract('Public Sub Init()\nEnd Sub', 'bas', 'ModA');
    const defs2 = extract('Public Sub Init()\nEnd Sub', 'bas', 'ModB');
    const table = buildSymbolTable([...defs1, ...defs2]);

    expect(table.symbols.get('init')).toHaveLength(2);
  });

  it('indexes symbols by module name', () => {
    const defs = extract('Public Sub Init()\nEnd Sub\nPublic gCount As Long', 'bas', 'MyModule');
    const table = buildSymbolTable(defs);

    expect(table.byModule.get('MyModule')).toHaveLength(2);
  });

  it('handles empty definitions array', () => {
    const table = buildSymbolTable([]);
    expect(table.symbols.size).toBe(0);
    expect(table.byModule.size).toBe(0);
  });
});

// ?? Complex real-world patterns ??????????????????????????????????????????????

describe('extractSymbols — Real-world patterns', () => {
  it('handles a typical .bas module header', () => {
    const code = [
      "' Module comment",
      'Option Explicit',
      'Public langPrefix As String',
      'Public SeguroGame As Boolean',
      '',
      'Public Enum tMacro',
      '  dobleclick = 1',
      '  Coordenadas = 2',
      'End Enum',
      '',
      'Public Const NO_WEAPON As Byte = 2',
      '',
      'Public Sub Init()',
      '  Dim i As Long',
      'End Sub',
    ].join('\n');
    const defs = extract(code);

    const names = defs.map(d => d.name);
    expect(names).toContain('langPrefix');
    expect(names).toContain('SeguroGame');
    expect(names).toContain('tMacro');
    expect(names).toContain('NO_WEAPON');
    expect(names).toContain('Init');
    expect(names).toContain('i');

    expect(findByName(defs, 'i')!.scope).toBe('local');
    expect(findByName(defs, 'langPrefix')!.scope).toBe('module');
  });

  it('handles .frm event handlers after preamble', () => {
    const code = [
      'VERSION 5.00',
      'Begin VB.Form frmTest',
      '   Caption = "Test"',
      'End',
      'Attribute VB_Name = "frmTest"',
      'Option Explicit',
      'Private Sub Form_Load()',
      '  Dim x As Long',
      'End Sub',
      'Private Sub cmdOk_Click()',
      'End Sub',
    ].join('\n');
    const defs = extract(code, 'frm', 'frmTest');

    const formLoad = findByName(defs, 'Form_Load');
    expect(formLoad).toBeDefined();
    expect(formLoad!.isEventHandler).toBe(true);

    const cmdOk = findByName(defs, 'cmdOk_Click');
    expect(cmdOk).toBeDefined();
    expect(cmdOk!.isEventHandler).toBe(true);
  });

  it('skips comment lines and blank lines', () => {
    const code = [
      "' This is a comment",
      '',
      'Public Sub Test()',
      "' Another comment",
      'End Sub',
    ].join('\n');
    const defs = extract(code);
    expect(defs).toHaveLength(1);
    expect(defs[0].name).toBe('Test');
  });

  it('does not confuse Public Const with Public variable', () => {
    const code = 'Public Const MAX_VAL As Long = 100';
    const defs = extract(code);
    expect(defs).toHaveLength(1);
    expect(defs[0].kind).toBe('Const');
  });

  it('does not confuse Declare with regular Sub', () => {
    const code = 'Private Declare Function GetTickCount Lib "kernel32" () As Long';
    const defs = extract(code);
    expect(defs).toHaveLength(1);
    expect(defs[0].kind).toBe('Declare');
  });
});
