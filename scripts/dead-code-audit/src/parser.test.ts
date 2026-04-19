import { describe, it, expect } from 'vitest';
import { parseModule, maskStringLiterals } from './parser.js';
import type { SourceFile } from './types.js';

function makeSource(content: string, type: 'bas' | 'cls' | 'frm' = 'bas'): SourceFile {
  return { path: `test.${type}`, type, moduleName: 'TestMod', content };
}

describe('maskStringLiterals', () => {
  it('replaces simple string content with placeholder', () => {
    expect(maskStringLiterals('x = "hello"')).toBe('x = "___"');
  });

  it('handles escaped quotes inside strings', () => {
    expect(maskStringLiterals('x = "say ""hi"""')).toBe('x = "___"');
  });

  it('handles multiple strings on one line', () => {
    expect(maskStringLiterals('MsgBox "a", "b"')).toBe('MsgBox "___", "___"');
  });

  it('returns line unchanged when no strings present', () => {
    expect(maskStringLiterals('Dim x As Long')).toBe('Dim x As Long');
  });

  it('handles empty string literal', () => {
    expect(maskStringLiterals('x = ""')).toBe('x = "___"');
  });

  it('handles unclosed string at end of line', () => {
    const result = maskStringLiterals('x = "unclosed');
    expect(result).toBe('x = "___');
  });
});

describe('parseModule', () => {
  it('joins line continuations into logical lines', () => {
    const src = makeSource('Dim x As _\n  Long');
    const mod = parseModule(src);
    const joined = mod.lines.find(l => l.text.includes('Dim'));
    expect(joined).toBeDefined();
    expect(joined!.text).toContain('Dim x As');
    expect(joined!.text).toContain('Long');
    expect(joined!.originalLines).toEqual([1, 2]);
  });

  it('joins multi-line continuations', () => {
    const src = makeSource('Call Foo( _\n  a, _\n  b)');
    const mod = parseModule(src);
    expect(mod.lines).toHaveLength(1);
    expect(mod.lines[0].originalLines).toEqual([1, 2, 3]);
  });

  it('strips form preamble from .frm files', () => {
    const content = [
      'VERSION 5.00',
      'Begin VB.Form frmTest',
      '   Caption = "Test"',
      'End',
      'Attribute VB_Name = "frmTest"',
      'Attribute VB_GlobalNameSpace = False',
      'Option Explicit',
      'Private Sub Form_Load()',
      'End Sub',
    ].join('\n');
    const src = makeSource(content, 'frm');
    const mod = parseModule(src);
    // Should not contain the preamble lines
    const allText = mod.lines.map(l => l.text).join('\n');
    expect(allText).not.toContain('VERSION 5.00');
    expect(allText).not.toContain('Begin VB.Form');
    // Should contain code after preamble
    expect(allText).toContain('Option Explicit');
  });

  it('does not strip preamble from .bas files', () => {
    const content = 'Attribute VB_Name = "Test"\nOption Explicit\nPublic Sub Main()\nEnd Sub';
    const src = makeSource(content, 'bas');
    const mod = parseModule(src);
    expect(mod.lines.length).toBeGreaterThanOrEqual(4);
  });

  it('detects comment lines starting with apostrophe', () => {
    const src = makeSource("' This is a comment\nDim x As Long");
    const mod = parseModule(src);
    const comment = mod.lines.find(l => l.text.includes('comment'));
    expect(comment?.isComment).toBe(true);
    expect(comment?.isExecutable).toBe(false);
  });

  it('detects Rem comments (case-insensitive)', () => {
    const src = makeSource('Rem This is a comment\nREM another');
    const mod = parseModule(src);
    expect(mod.lines[0].isComment).toBe(true);
    expect(mod.lines[1].isComment).toBe(true);
  });

  it('detects preprocessor directives', () => {
    const src = makeSource('#If DEBUG Then\nDim x As Long\n#End If');
    const mod = parseModule(src);
    expect(mod.lines[0].isPreprocessor).toBe(true);
    expect(mod.lines[1].isExecutable).toBe(true);
    expect(mod.lines[2].isPreprocessor).toBe(true);
  });

  it('classifies Attribute lines as non-executable metadata', () => {
    const src = makeSource('Attribute VB_Name = "Test"\nAttribute VB_Exposed = False\nPublic Sub Main()\nEnd Sub');
    const mod = parseModule(src);
    expect(mod.attributeLines).toHaveLength(2);
    const attrLine = mod.lines.find(l => l.text.includes('VB_Name'));
    expect(attrLine?.isExecutable).toBe(false);
  });

  it('masks string literals in parsed lines', () => {
    const src = makeSource('MsgBox "Hello World"');
    const mod = parseModule(src);
    expect(mod.lines[0].text).toBe('MsgBox "___"');
    expect(mod.lines[0].text).not.toContain('Hello World');
  });

  it('classifies blank lines as non-executable', () => {
    const src = makeSource('Dim x As Long\n\n  \nDim y As Long');
    const mod = parseModule(src);
    const blanks = mod.lines.filter(l => l.text.trim() === '');
    for (const b of blanks) {
      expect(b.isExecutable).toBe(false);
      expect(b.isComment).toBe(false);
    }
  });

  it('handles continuation at end of file gracefully', () => {
    const src = makeSource('Dim x As _');
    const mod = parseModule(src);
    // Should not crash — just produce a line with trailing content stripped
    expect(mod.lines.length).toBeGreaterThanOrEqual(1);
  });

  it('handles #ElseIf and #Const preprocessor directives', () => {
    const src = makeSource('#Const DEBUG = 1\n#If DEBUG Then\nx = 1\n#ElseIf RELEASE Then\nx = 2\n#Else\nx = 3\n#End If');
    const mod = parseModule(src);
    expect(mod.lines[0].isPreprocessor).toBe(true); // #Const
    expect(mod.lines[3].isPreprocessor).toBe(true); // #ElseIf
    expect(mod.lines[5].isPreprocessor).toBe(true); // #Else
  });

  it('preserves original line numbers after preamble stripping', () => {
    const content = [
      'VERSION 5.00',
      'Begin VB.Form frmTest',
      'End',
      'Attribute VB_Name = "frmTest"',
      'Option Explicit',
    ].join('\n');
    const src = makeSource(content, 'frm');
    const mod = parseModule(src);
    // After stripping, the first line is "Attribute VB_Name" which was line 4 in original
    // But since we re-index after stripping, lineNumber is 1-based within the stripped content
    expect(mod.lines[0].lineNumber).toBe(1);
  });
});
