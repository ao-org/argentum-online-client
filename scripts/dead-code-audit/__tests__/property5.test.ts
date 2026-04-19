/**
 * Property 5: Unreachable code after unconditional exits
 *
 * For any procedure body containing an unconditional exit statement
 * (Exit Sub, Exit Function, Exit Property, End, or unconditional GoTo)
 * within a block, all executable statements after that exit and before
 * the next label or block-end keyword must be reported as unreachable.
 *
 * Validates: Requirements 4.1
 */
import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { detectUnreachableCode } from '../src/deadCodeDetector.js';
import type { ParsedModule, ParsedLine, SourceFile } from '../src/types.js';

// --- Helpers ---

/** Build a ParsedLine from text and line number. */
function makeLine(lineNumber: number, text: string): ParsedLine {
  const trimmed = text.trim();
  const isComment = /^('|Rem\s)/i.test(trimmed);
  const isPreprocessor = /^#/i.test(trimmed);
  const isBlank = trimmed === '';
  const isAttribute = /^Attribute\s/i.test(trimmed);
  const isExecutable = !isComment && !isPreprocessor && !isBlank && !isAttribute;

  return {
    lineNumber,
    text,
    isComment,
    isPreprocessor,
    isExecutable,
    originalLines: [lineNumber],
  };
}

/** Build a ParsedModule from an array of text lines (1-indexed). */
function buildModule(lines: string[], filePath = 'TestModule.bas'): ParsedModule {
  const source: SourceFile = {
    path: filePath,
    type: 'bas',
    moduleName: 'TestModule',
    content: lines.join('\n'),
  };

  return {
    source,
    lines: lines.map((text, i) => makeLine(i + 1, text)),
    attributeLines: [],
  };
}

// --- Arbitraries ---

/** Procedure wrapper types with matching start/end keywords. */
const procedureTypes = [
  { start: 'Sub TestProc()', end: 'End Sub', exit: 'Exit Sub' },
  { start: 'Function TestFunc() As Long', end: 'End Function', exit: 'Exit Function' },
  { start: 'Property Get TestProp() As Long', end: 'End Property', exit: 'Exit Property' },
] as const;

const procedureTypeArb = fc.constantFrom(...procedureTypes);

/** Unconditional exit statements. */
const exitStatementArb = fc.constantFrom(
  'Exit Sub',
  'Exit Function',
  'Exit Property',
  'End',
  'GoTo CleanUp',
);

/** Simple executable VB6 statements that are NOT exit/goto/label/block-end. */
const executableStatementArb = fc.constantFrom(
  'x = 1',
  'Call DoSomething',
  'y = x + 1',
  'Debug.Print "hello"',
  'MsgBox "test"',
  'z = Len(str)',
  'i = i + 1',
  'result = Calculate(a, b)',
  'Set obj = Nothing',
  'ReDim arr(10)',
);

/** Generate 1-5 executable statements to place before the exit. */
const preExitStatementsArb = fc.array(executableStatementArb, { minLength: 1, maxLength: 5 });

/** Generate 1-5 executable statements to place after the exit (these should be unreachable). */
const postExitStatementsArb = fc.array(executableStatementArb, { minLength: 1, maxLength: 5 });

// --- Property Tests ---

describe('Feature: dead-code-audit, Property 5: Unreachable code after unconditional exits', () => {
  it('all executable lines after an unconditional exit are reported as unreachable', () => {
    /**
     * Validates: Requirements 4.1
     *
     * Strategy:
     * 1. Generate a procedure with: start line, some executable lines,
     *    an unconditional exit, then 1+ trailing executable lines, then End Sub/Function
     * 2. Build ParsedModule from the generated lines
     * 3. Call detectUnreachableCode
     * 4. Verify that ALL trailing executable lines after the exit are reported
     */
    fc.assert(
      fc.property(
        procedureTypeArb,
        exitStatementArb,
        preExitStatementsArb,
        postExitStatementsArb,
        (procType, exitStmt, preStmts, postStmts) => {
          // Build the procedure body
          const lines: string[] = [];
          lines.push(procType.start);
          for (const stmt of preStmts) {
            lines.push(`  ${stmt}`);
          }
          lines.push(`  ${exitStmt}`);

          // Track which line numbers should be unreachable
          const unreachableLineNumbers: number[] = [];
          for (const stmt of postStmts) {
            lines.push(`  ${stmt}`);
            unreachableLineNumbers.push(lines.length); // 1-indexed
          }
          lines.push(procType.end);

          const mod = buildModule(lines);
          const findings = detectUnreachableCode([mod]);

          // Collect all line numbers reported as unreachable
          const reportedLines = new Set<number>();
          for (const f of findings) {
            if (f.category === 'unreachable-code') {
              for (let ln = f.startLine; ln <= f.endLine; ln++) {
                reportedLines.add(ln);
              }
            }
          }

          // Every trailing executable line must be reported
          for (const ln of unreachableLineNumbers) {
            expect(reportedLines.has(ln)).toBe(true);
          }
        },
      ),
      { numRuns: 100 },
    );
  });

  it('lines before the exit statement are NOT reported as unreachable', () => {
    /**
     * Validates: Requirements 4.1
     *
     * Strategy:
     * Verify that executable lines BEFORE the exit are not flagged.
     */
    fc.assert(
      fc.property(
        procedureTypeArb,
        exitStatementArb,
        preExitStatementsArb,
        postExitStatementsArb,
        (procType, exitStmt, preStmts, postStmts) => {
          const lines: string[] = [];
          lines.push(procType.start);

          const reachableLineNumbers: number[] = [];
          for (const stmt of preStmts) {
            lines.push(`  ${stmt}`);
            reachableLineNumbers.push(lines.length);
          }

          // The exit statement itself is reachable (it executes)
          lines.push(`  ${exitStmt}`);
          const exitLineNumber = lines.length;

          for (const stmt of postStmts) {
            lines.push(`  ${stmt}`);
          }
          lines.push(procType.end);

          const mod = buildModule(lines);
          const findings = detectUnreachableCode([mod]);

          // Collect all line numbers reported as unreachable
          const reportedLines = new Set<number>();
          for (const f of findings) {
            if (f.category === 'unreachable-code') {
              for (let ln = f.startLine; ln <= f.endLine; ln++) {
                reportedLines.add(ln);
              }
            }
          }

          // Lines before the exit must NOT be reported
          for (const ln of reachableLineNumbers) {
            expect(reportedLines.has(ln)).toBe(false);
          }

          // The exit statement itself must NOT be reported
          expect(reportedLines.has(exitLineNumber)).toBe(false);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('a label after the exit resets unreachable state', () => {
    /**
     * Validates: Requirements 4.1
     *
     * Strategy:
     * After an exit + unreachable lines, a label resets reachability.
     * Lines after the label should NOT be reported as unreachable.
     */
    fc.assert(
      fc.property(
        preExitStatementsArb,
        postExitStatementsArb,
        fc.array(executableStatementArb, { minLength: 1, maxLength: 3 }),
        (preStmts, unreachableStmts, afterLabelStmts) => {
          const lines: string[] = [];
          lines.push('Sub TestProc()');
          for (const stmt of preStmts) {
            lines.push(`  ${stmt}`);
          }
          lines.push('  Exit Sub');

          // Unreachable lines
          const unreachableLineNumbers: number[] = [];
          for (const stmt of unreachableStmts) {
            lines.push(`  ${stmt}`);
            unreachableLineNumbers.push(lines.length);
          }

          // Label resets reachability
          lines.push('CleanUp:');

          // Lines after label should be reachable
          const afterLabelLineNumbers: number[] = [];
          for (const stmt of afterLabelStmts) {
            lines.push(`  ${stmt}`);
            afterLabelLineNumbers.push(lines.length);
          }
          lines.push('End Sub');

          const mod = buildModule(lines);
          const findings = detectUnreachableCode([mod]);

          const reportedLines = new Set<number>();
          for (const f of findings) {
            if (f.category === 'unreachable-code') {
              for (let ln = f.startLine; ln <= f.endLine; ln++) {
                reportedLines.add(ln);
              }
            }
          }

          // Unreachable lines before label must be reported
          for (const ln of unreachableLineNumbers) {
            expect(reportedLines.has(ln)).toBe(true);
          }

          // Lines after label must NOT be reported
          for (const ln of afterLabelLineNumbers) {
            expect(reportedLines.has(ln)).toBe(false);
          }
        },
      ),
      { numRuns: 100 },
    );
  });
});
