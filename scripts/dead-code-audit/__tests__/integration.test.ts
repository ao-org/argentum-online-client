/**
 * Integration tests for the Dead Code Audit pipeline.
 *
 * Creates synthetic VB6 project fixtures with known dead code,
 * runs the full pipeline, and verifies all findings.
 *
 * Validates: Requirements 1.1, 1.2, 2.1, 3.1, 4.1, 5.1, 7.1, 7.2
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';

import { discoverFiles } from '../src/fileDiscovery.js';
import { parseModule } from '../src/parser.js';
import { extractSymbols, buildSymbolTable } from '../src/symbolExtractor.js';
import { scanReferences } from '../src/referenceScanner.js';
import { analyzeUsage } from '../src/crossRefAnalyzer.js';
import {
  detectUnusedSymbols,
  detectUnreachableCode,
  detectCommentedOutBlocks,
} from '../src/deadCodeDetector.js';
import { detectDuplicates } from '../src/duplicateDetector.js';
import {
  createBatches,
  applyBatch,
  validateVB6Syntax,
} from '../src/removalEngine.js';
import type { Finding, ParsedModule, SymbolDefinition } from '../src/types.js';

// ?? Fixture content ?????????????????????????????????????????????????

const MODULE1_BAS = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub UsedSub()
    Dim usedLocal As Long
    usedLocal = 42
    MsgBox CStr(usedLocal)
End Sub

Public Sub UnusedSub()
    Dim x As Long
    x = 1
End Sub

Private mUnusedVar As String

Private mUsedVar As Long

Public Sub CallerSub()
    mUsedVar = 10
    UsedSub
    Dim y As Long
    y = mUsedVar + 1
    MsgBox CStr(y)
    Exit Sub
    Dim unreachableVar As Long
    unreachableVar = 99
    MsgBox "never reached"
End Sub

' This is a commented-out code block
' Sub OldRoutine()
'     Dim oldVal As Long
'     oldVal = 1
'     If oldVal > 0 Then
'         Call DoSomething
'     End If
' End Sub

Public Sub DuplicateBlock()
    Dim a As Long
    Dim b As Long
    a = 1
    b = 2
    If a > b Then
        MsgBox "a is greater"
    Else
        MsgBox "b is greater"
    End If
    MsgBox "done"
End Sub
`;

const MODULE2_BAS = `Attribute VB_Name = "Module2"
Option Explicit

Public Const UNUSED_CONST As Long = 999

Public Type UnusedType
    x As Long
    y As Long
End Type

Public Sub DuplicateBlock2()
    Dim a As Long
    Dim b As Long
    a = 1
    b = 2
    If a > b Then
        MsgBox "a is greater"
    Else
        MsgBox "b is greater"
    End If
    MsgBox "done"
End Sub
`;

const CLASS1_CLS = `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "MyClass"
Option Explicit

Private mName As String

Public Property Get Name() As String
    Name = mName
End Property

Private Function UnusedFunction() As Long
    Dim tmp As Long
    tmp = 42
End Function

Public Sub UseProperty()
    mName = "test"
    Dim n As String
    n = Name
    MsgBox n
End Sub
`;

// ?? Helpers ?????????????????????????????????????????????????????????

const tempDirs: string[] = [];

function createTempProject(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vb6-integration-'));
  tempDirs.push(dir);
  return dir;
}

function writeFixture(rootDir: string, filename: string, content: string): string {
  const filePath = path.join(rootDir, filename);
  fs.writeFileSync(filePath, content, 'utf-8');
  return filePath;
}

/** Run the full analysis pipeline on a directory and return all findings + duplicates. */
function runPipeline(rootDir: string) {
  const sourceFiles = discoverFiles(rootDir);
  const modules: ParsedModule[] = sourceFiles.map(sf => parseModule(sf));

  const allDefs: SymbolDefinition[] = [];
  for (const mod of modules) {
    allDefs.push(...extractSymbols(mod));
  }
  const symbolTable = buildSymbolTable(allDefs);
  const referenceMap = scanReferences(modules, symbolTable);
  const usages = analyzeUsage(symbolTable, referenceMap);

  const unusedFindings = detectUnusedSymbols(usages);
  const unreachableFindings = detectUnreachableCode(modules);
  const commentedOutFindings = detectCommentedOutBlocks(modules);
  const duplicates = detectDuplicates(modules);

  const allFindings: Finding[] = [
    ...unusedFindings,
    ...unreachableFindings,
    ...commentedOutFindings,
  ];

  return { allFindings, duplicates, modules, sourceFiles };
}

afterEach(() => {
  for (const dir of tempDirs) {
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch { /* ignore cleanup errors */ }
  }
  tempDirs.length = 0;
});

// ?? Tests ???????????????????????????????????????????????????????????

describe('Integration: full pipeline on synthetic VB6 project', () => {
  it('detects all known dead code and does not flag used code', () => {
    const root = createTempProject();
    writeFixture(root, 'Module1.bas', MODULE1_BAS);
    writeFixture(root, 'Module2.bas', MODULE2_BAS);
    writeFixture(root, 'MyClass.cls', CLASS1_CLS);

    const { allFindings, duplicates } = runPipeline(root);

    const findingNames = allFindings.map(f => f.symbolName?.toLowerCase());
    const findingCategories = allFindings.map(f => f.category);

    // ?? Unused Sub is found ??
    const unusedSubFinding = allFindings.find(
      f => f.category === 'unused-procedure' && f.symbolName?.toLowerCase() === 'unusedsub',
    );
    expect(unusedSubFinding).toBeDefined();

    // ?? Unused variable is found ??
    const unusedVarFinding = allFindings.find(
      f => f.category === 'unused-variable' && f.symbolName?.toLowerCase() === 'munusedvar',
    );
    expect(unusedVarFinding).toBeDefined();

    // ?? Used Sub is NOT flagged ??
    const usedSubFlagged = allFindings.find(
      f => f.category === 'unused-procedure' && f.symbolName?.toLowerCase() === 'usedsub',
    );
    expect(usedSubFlagged).toBeUndefined();

    // ?? Used variable is NOT flagged ??
    const usedVarFlagged = allFindings.find(
      f => f.category === 'unused-variable' && f.symbolName?.toLowerCase() === 'musedvar',
    );
    expect(usedVarFlagged).toBeUndefined();

    // ?? Unreachable code is found (after Exit Sub in CallerSub) ??
    const unreachableFinding = allFindings.find(
      f => f.category === 'unreachable-code',
    );
    expect(unreachableFinding).toBeDefined();

    // ?? Commented-out block is found ??
    const commentedOutFinding = allFindings.find(
      f => f.category === 'commented-out-block',
    );
    expect(commentedOutFinding).toBeDefined();

    // ?? Duplicate block is found ??
    expect(duplicates.length).toBeGreaterThanOrEqual(1);
    const dupPair = duplicates[0];
    expect(dupPair.lineCount).toBeGreaterThanOrEqual(10);

    // ?? Unused Const is found ??
    const unusedConstFinding = allFindings.find(
      f => f.category === 'unused-const' && f.symbolName?.toLowerCase() === 'unused_const',
    );
    expect(unusedConstFinding).toBeDefined();

    // ?? Unused Type is found ??
    const unusedTypeFinding = allFindings.find(
      f => f.category === 'unused-type' && f.symbolName?.toLowerCase() === 'unusedtype',
    );
    expect(unusedTypeFinding).toBeDefined();

    // ?? Unused Function is found ??
    const unusedFuncFinding = allFindings.find(
      f => f.category === 'unused-procedure' && f.symbolName?.toLowerCase() === 'unusedfunction',
    );
    expect(unusedFuncFinding).toBeDefined();
  });

  it('safe removal creates backup, removes lines, and keeps valid syntax', () => {
    const root = createTempProject();
    const filePath = writeFixture(root, 'Module1.bas', MODULE1_BAS);

    const { allFindings } = runPipeline(root);

    // discoverFiles returns relative paths; patch to absolute for applyBatch
    for (const f of allFindings) {
      if (f.filePath === 'Module1.bas') {
        f.filePath = filePath;
      }
    }

    // Filter to only Module1 findings that are confirmed + removable
    const module1Findings = allFindings.filter(
      f => f.filePath === filePath && f.confidence === 'confirmed' && f.removable,
    );
    expect(module1Findings.length).toBeGreaterThan(0);

    const batches = createBatches(module1Findings);
    expect(batches.length).toBe(1);

    const batch = batches[0];
    const result = applyBatch(batch);

    // Backup was created
    const backupPath = filePath + '.bak';
    expect(fs.existsSync(backupPath)).toBe(true);

    // Removal succeeded
    expect(result.success).toBe(true);
    expect(result.syntaxValid).toBe(true);
    expect(result.reverted).toBe(false);

    // Verify removed content is gone
    const modifiedContent = fs.readFileSync(filePath, 'utf-8');

    // The unused sub body should be removed
    expect(modifiedContent.toLowerCase()).not.toContain('public sub unusedsub');

    // The used sub should still be present
    expect(modifiedContent.toLowerCase()).toContain('public sub usedsub');

    // Syntax should still be valid after removal
    const validation = validateVB6Syntax(modifiedContent);
    expect(validation.valid).toBe(true);
  });

  it('reverts removal when it would break syntax', () => {
    // Create a file where we'll craft a batch that removes End Sub but not Sub
    const root = createTempProject();
    const content = `Attribute VB_Name = "BrokenModule"
Option Explicit

Public Sub TestSub()
    Dim x As Long
    x = 1
End Sub
`;
    const filePath = writeFixture(root, 'BrokenModule.bas', content);

    // Create a batch that removes only the "End Sub" line (line 7), breaking syntax
    const brokenBatch = {
      filePath,
      moduleName: 'BrokenModule',
      findings: [
        {
          id: 'finding-break-1',
          category: 'unreachable-code' as const,
          confidence: 'confirmed' as const,
          filePath,
          startLine: 7,
          endLine: 7,
          reason: 'test',
          removable: true,
        },
      ],
      linesRemoved: 1,
    };

    const result = applyBatch(brokenBatch);

    // Should have been reverted due to syntax validation failure
    expect(result.success).toBe(false);
    expect(result.syntaxValid).toBe(false);
    expect(result.reverted).toBe(true);

    // Original content should be restored
    const restoredContent = fs.readFileSync(filePath, 'utf-8');
    expect(restoredContent).toContain('End Sub');
  });
});
