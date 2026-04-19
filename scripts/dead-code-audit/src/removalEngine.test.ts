import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import {
  createBatches,
  validateVB6Syntax,
  applyBatch,
  generateRemovalSummary,
} from './removalEngine.js';
import type { Finding, RemovalBatch, RemovalResult } from './types.js';

/* ------------------------------------------------------------------ */
/*  Helpers                                                           */
/* ------------------------------------------------------------------ */

function makeFinding(overrides: Partial<Finding> = {}): Finding {
  return {
    id: 'f1',
    category: 'unused-procedure',
    confidence: 'confirmed',
    filePath: 'Module1.bas',
    startLine: 5,
    endLine: 10,
    reason: 'Never called',
    removable: true,
    ...overrides,
  };
}

/* ------------------------------------------------------------------ */
/*  12.1 – createBatches                                              */
/* ------------------------------------------------------------------ */

describe('createBatches', () => {
  it('groups confirmed removable findings by file path', () => {
    const findings: Finding[] = [
      makeFinding({ id: 'f1', filePath: 'A.bas', startLine: 1, endLine: 3 }),
      makeFinding({ id: 'f2', filePath: 'B.bas', startLine: 10, endLine: 12 }),
      makeFinding({ id: 'f3', filePath: 'A.bas', startLine: 5, endLine: 7 }),
    ];
    const batches = createBatches(findings);
    expect(batches).toHaveLength(2);

    const batchA = batches.find((b) => b.filePath === 'A.bas')!;
    expect(batchA.findings).toHaveLength(2);
    expect(batchA.moduleName).toBe('A');

    const batchB = batches.find((b) => b.filePath === 'B.bas')!;
    expect(batchB.findings).toHaveLength(1);
    expect(batchB.moduleName).toBe('B');
  });

  it('excludes findings that are not confirmed', () => {
    const findings: Finding[] = [
      makeFinding({ confidence: 'review-needed' }),
    ];
    expect(createBatches(findings)).toHaveLength(0);
  });

  it('excludes findings that are not removable', () => {
    const findings: Finding[] = [
      makeFinding({ removable: false }),
    ];
    expect(createBatches(findings)).toHaveLength(0);
  });

  it('calculates linesRemoved as the union of line ranges', () => {
    // Overlapping ranges: [1,5] and [3,7] ? union is {1,2,3,4,5,6,7} = 7 lines
    const findings: Finding[] = [
      makeFinding({ id: 'f1', filePath: 'X.bas', startLine: 1, endLine: 5 }),
      makeFinding({ id: 'f2', filePath: 'X.bas', startLine: 3, endLine: 7 }),
    ];
    const batches = createBatches(findings);
    expect(batches).toHaveLength(1);
    expect(batches[0].linesRemoved).toBe(7);
  });

  it('returns empty array when no findings are eligible', () => {
    expect(createBatches([])).toHaveLength(0);
  });

  it('extracts moduleName from filePath without extension', () => {
    const findings: Finding[] = [
      makeFinding({ filePath: 'path/to/MyModule.cls' }),
    ];
    const batches = createBatches(findings);
    expect(batches[0].moduleName).toBe('MyModule');
  });
});

/* ------------------------------------------------------------------ */
/*  12.2 – validateVB6Syntax                                          */
/* ------------------------------------------------------------------ */

describe('validateVB6Syntax', () => {
  it('returns valid for balanced Sub/End Sub', () => {
    const content = [
      'Public Sub Foo()',
      '  Dim x As Long',
      'End Sub',
    ].join('\n');
    const result = validateVB6Syntax(content);
    expect(result.valid).toBe(true);
    expect(result.errors).toHaveLength(0);
  });

  it('returns valid for balanced Function/End Function', () => {
    const content = [
      'Private Function Bar() As Long',
      '  Bar = 42',
      'End Function',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('returns valid for balanced Property/End Property', () => {
    const content = [
      'Public Property Get Name() As String',
      '  Name = mName',
      'End Property',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('returns valid for balanced If/End If', () => {
    const content = [
      'If x > 0 Then',
      '  y = 1',
      'End If',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('returns valid for balanced For/Next', () => {
    const content = [
      'For i = 1 To 10',
      '  x = x + 1',
      'Next',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('returns valid for balanced Do/Loop', () => {
    const content = [
      'Do While x > 0',
      '  x = x - 1',
      'Loop',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('returns valid for balanced While/Wend', () => {
    const content = [
      'While x > 0',
      '  x = x - 1',
      'Wend',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('returns valid for balanced Select Case/End Select', () => {
    const content = [
      'Select Case x',
      '  Case 1',
      '    y = 1',
      'End Select',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('returns valid for balanced With/End With', () => {
    const content = [
      'With obj',
      '  .Name = "test"',
      'End With',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('returns valid for balanced Type/End Type', () => {
    const content = [
      'Private Type Position',
      '  X As Long',
      '  Y As Long',
      'End Type',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('returns valid for balanced Enum/End Enum', () => {
    const content = [
      'Public Enum Colors',
      '  Red = 1',
      '  Blue = 2',
      'End Enum',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('detects unclosed Sub', () => {
    const content = [
      'Public Sub Foo()',
      '  Dim x As Long',
    ].join('\n');
    const result = validateVB6Syntax(content);
    expect(result.valid).toBe(false);
    expect(result.errors.length).toBeGreaterThan(0);
    expect(result.errors[0]).toContain('Sub');
  });

  it('detects End Sub without matching Sub', () => {
    const content = 'End Sub';
    const result = validateVB6Syntax(content);
    expect(result.valid).toBe(false);
    expect(result.errors.length).toBeGreaterThan(0);
  });

  it('detects mismatched blocks', () => {
    const content = [
      'Public Sub Foo()',
      'End Function',
    ].join('\n');
    const result = validateVB6Syntax(content);
    expect(result.valid).toBe(false);
  });

  it('skips comment lines', () => {
    const content = [
      "' This is a comment",
      'Public Sub Foo()',
      "' Another comment",
      'End Sub',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('handles nested blocks', () => {
    const content = [
      'Public Sub Foo()',
      '  If x > 0 Then',
      '    For i = 1 To 10',
      '      Do While y > 0',
      '        y = y - 1',
      '      Loop',
      '    Next',
      '  End If',
      'End Sub',
    ].join('\n');
    expect(validateVB6Syntax(content).valid).toBe(true);
  });

  it('returns valid for empty content', () => {
    expect(validateVB6Syntax('').valid).toBe(true);
  });
});


/* ------------------------------------------------------------------ */
/*  12.3 – applyBatch                                                 */
/* ------------------------------------------------------------------ */

describe('applyBatch', () => {
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'removal-test-'));
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  it('removes identified lines and creates a .bak backup', () => {
    const filePath = path.join(tmpDir, 'Module1.bas');
    const original = [
      'Public Sub Foo()',       // line 1
      '  Dim x As Long',       // line 2
      '  x = 42',              // line 3
      'End Sub',               // line 4
      'Public Sub Bar()',      // line 5
      '  Dim y As Long',       // line 6
      'End Sub',               // line 7
    ].join('\n');
    fs.writeFileSync(filePath, original);

    const batch: RemovalBatch = {
      filePath,
      moduleName: 'Module1',
      findings: [
        makeFinding({ filePath, startLine: 5, endLine: 7 }),
      ],
      linesRemoved: 3,
    };

    const result = applyBatch(batch);
    expect(result.success).toBe(true);
    expect(result.syntaxValid).toBe(true);
    expect(result.reverted).toBe(false);
    expect(result.testsPass).toBeNull();

    // Backup should exist
    expect(fs.existsSync(filePath + '.bak')).toBe(true);
    expect(fs.readFileSync(filePath + '.bak', 'utf-8')).toBe(original);

    // Modified file should not contain removed lines
    const modified = fs.readFileSync(filePath, 'utf-8');
    expect(modified).not.toContain('Public Sub Bar()');
    expect(modified).toContain('Public Sub Foo()');
  });

  it('reverts when syntax is invalid after removal', () => {
    const filePath = path.join(tmpDir, 'Module2.bas');
    // Removing lines 2-3 would leave Sub without End Sub
    const original = [
      'Public Sub Foo()',       // line 1
      '  x = 42',              // line 2
      'End Sub',               // line 3
    ].join('\n');
    fs.writeFileSync(filePath, original);

    const batch: RemovalBatch = {
      filePath,
      moduleName: 'Module2',
      findings: [
        makeFinding({ filePath, startLine: 2, endLine: 3 }),
      ],
      linesRemoved: 2,
    };

    const result = applyBatch(batch);
    expect(result.success).toBe(false);
    expect(result.syntaxValid).toBe(false);
    expect(result.reverted).toBe(true);

    // File should be reverted to original
    const content = fs.readFileSync(filePath, 'utf-8');
    expect(content).toBe(original);
  });

  it('reclassifies findings as review-needed on syntax failure', () => {
    const filePath = path.join(tmpDir, 'Module3.bas');
    const original = [
      'Public Sub Foo()',
      '  x = 42',
      'End Sub',
    ].join('\n');
    fs.writeFileSync(filePath, original);

    const finding = makeFinding({ filePath, startLine: 2, endLine: 3 });
    expect(finding.confidence).toBe('confirmed');

    const batch: RemovalBatch = {
      filePath,
      moduleName: 'Module3',
      findings: [finding],
      linesRemoved: 2,
    };

    applyBatch(batch);
    expect(finding.confidence).toBe('review-needed');
  });

  it('returns failure when file does not exist', () => {
    const batch: RemovalBatch = {
      filePath: path.join(tmpDir, 'nonexistent.bas'),
      moduleName: 'nonexistent',
      findings: [makeFinding()],
      linesRemoved: 1,
    };
    const result = applyBatch(batch);
    expect(result.success).toBe(false);
    expect(result.reverted).toBe(false);
  });
});

/* ------------------------------------------------------------------ */
/*  12.4 – generateRemovalSummary                                     */
/* ------------------------------------------------------------------ */

describe('generateRemovalSummary', () => {
  it('produces markdown with per-module counts and total', () => {
    const results: RemovalResult[] = [
      {
        batch: { filePath: 'A.bas', moduleName: 'A', findings: [], linesRemoved: 10 },
        success: true,
        syntaxValid: true,
        testsPass: null,
        reverted: false,
      },
      {
        batch: { filePath: 'B.bas', moduleName: 'B', findings: [], linesRemoved: 5 },
        success: true,
        syntaxValid: true,
        testsPass: null,
        reverted: false,
      },
    ];

    const summary = generateRemovalSummary(results);
    expect(summary).toContain('# Removal Summary');
    expect(summary).toContain('| A | 10 | Removed |');
    expect(summary).toContain('| B | 5 | Removed |');
    expect(summary).toContain('**Total lines removed: 15**');
  });

  it('shows 0 lines for reverted batches', () => {
    const results: RemovalResult[] = [
      {
        batch: { filePath: 'C.bas', moduleName: 'C', findings: [], linesRemoved: 8 },
        success: false,
        syntaxValid: false,
        testsPass: null,
        reverted: true,
      },
    ];

    const summary = generateRemovalSummary(results);
    expect(summary).toContain('| C | 0 | Reverted |');
    expect(summary).toContain('**Total lines removed: 0**');
  });

  it('shows Failed status for non-reverted failures', () => {
    const results: RemovalResult[] = [
      {
        batch: { filePath: 'D.bas', moduleName: 'D', findings: [], linesRemoved: 3 },
        success: false,
        syntaxValid: false,
        testsPass: null,
        reverted: false,
      },
    ];

    const summary = generateRemovalSummary(results);
    expect(summary).toContain('| D | 0 | Failed |');
  });

  it('returns summary with zero total for empty results', () => {
    const summary = generateRemovalSummary([]);
    expect(summary).toContain('**Total lines removed: 0**');
  });
});
