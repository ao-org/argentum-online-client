/**
 * Property 6: Commented-out block detection threshold
 *
 * For any file containing consecutive comment lines, only blocks of more than
 * 5 consecutive comment lines that contain VB6 keywords are reported as
 * "commented-out code candidate for removal". Blocks of 5 or fewer consecutive
 * comment lines must not be reported.
 *
 * Validates: Requirements 4.4
 */
import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { detectCommentedOutBlocks } from '../src/deadCodeDetector.js';
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

/** VB6 keywords that the detector looks for in comment blocks. */
const vb6Keywords = [
  'Sub', 'Function', 'Dim', 'If', 'For', 'Call', 'End',
  'Select', 'Case', 'Do', 'Loop', 'While', 'Wend', 'Next',
  'Exit', 'GoTo', 'GoSub', 'Return', 'Set', 'Let', 'With',
  'Property', 'Public', 'Private', 'Const', 'Enum', 'Type',
  'Declare', 'ReDim', 'Erase', 'On Error', 'Resume',
];

/** Generate a comment line containing a VB6 keyword. */
const commentWithKeywordArb = fc.constantFrom(...vb6Keywords).map(
  (kw) => `' ${kw} SomeIdentifier`,
);

/** Generate a plain documentation comment (no VB6 keywords). */
const plainCommentArb = fc.constantFrom(
  "' This is a documentation comment",
  "' Author: developer",
  "' Date: 2024-01-01",
  "' Description of the module",
  "' TODO: refactor later",
  "' See also: other module",
  "' Version 1.0",
  "' Notes about the algorithm",
  "' Parameters explained here",
  "' Returns a value",
);

/** Generate a non-comment executable line to act as a separator. */
const executableLineArb = fc.constantFrom(
  'x = 1',
  'Call DoSomething',
  'y = x + 1',
  'Debug.Print "hello"',
);

/** Generate a block of comment lines of a specific length, with or without VB6 keywords. */
function commentBlockArb(
  length: number,
  withKeywords: boolean,
): fc.Arbitrary<string[]> {
  if (withKeywords) {
    // At least one line must have a keyword, rest can be mixed
    return fc.tuple(
      commentWithKeywordArb,
      fc.array(
        fc.oneof(commentWithKeywordArb, plainCommentArb),
        { minLength: length - 1, maxLength: length - 1 },
      ),
    ).map(([kwLine, rest]) => {
      // Shuffle the keyword line into the block
      const block = [kwLine, ...rest];
      return block;
    });
  } else {
    return fc.array(plainCommentArb, { minLength: length, maxLength: length });
  }
}

// --- Property Tests ---

describe('Feature: dead-code-audit, Property 6: Commented-out block detection threshold', () => {
  it('blocks of 6+ consecutive comment lines WITH VB6 keywords are reported', () => {
    /**
     * Validates: Requirements 4.4
     *
     * Strategy:
     * 1. Generate a comment block of 6-15 lines containing VB6 keywords
     * 2. Wrap it in a module with executable lines before/after as separators
     * 3. Call detectCommentedOutBlocks
     * 4. Verify the block IS reported
     */
    fc.assert(
      fc.property(
        fc.integer({ min: 6, max: 15 }).chain((blockLen) =>
          commentBlockArb(blockLen, true).map((block) => ({ blockLen, block })),
        ),
        ({ blockLen, block }) => {
          const lines: string[] = [];
          lines.push('x = 1'); // separator before
          const blockStartLine = lines.length + 1; // 1-indexed
          for (const commentLine of block) {
            lines.push(commentLine);
          }
          const blockEndLine = lines.length;
          lines.push('y = 2'); // separator after

          const mod = buildModule(lines);
          const findings = detectCommentedOutBlocks([mod]);

          // The block should be reported
          const matchingFindings = findings.filter(
            (f) =>
              f.category === 'commented-out-block' &&
              f.startLine >= blockStartLine &&
              f.endLine <= blockEndLine,
          );

          expect(matchingFindings.length).toBeGreaterThanOrEqual(1);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('blocks of 5 or fewer consecutive comment lines are NOT reported', () => {
    /**
     * Validates: Requirements 4.4
     *
     * Strategy:
     * 1. Generate a comment block of 1-5 lines (even with VB6 keywords)
     * 2. Wrap it in a module with executable lines before/after
     * 3. Call detectCommentedOutBlocks
     * 4. Verify the block is NOT reported
     */
    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 5 }).chain((blockLen) =>
          commentBlockArb(blockLen, true).map((block) => ({ blockLen, block })),
        ),
        ({ blockLen, block }) => {
          const lines: string[] = [];
          lines.push('x = 1'); // separator before
          for (const commentLine of block) {
            lines.push(commentLine);
          }
          lines.push('y = 2'); // separator after

          const mod = buildModule(lines);
          const findings = detectCommentedOutBlocks([mod]);

          // No findings should be reported for blocks <= 5 lines
          expect(findings.length).toBe(0);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('blocks of 6+ consecutive comment lines WITHOUT VB6 keywords are NOT reported', () => {
    /**
     * Validates: Requirements 4.4
     *
     * Strategy:
     * 1. Generate a comment block of 6-15 lines with NO VB6 keywords
     * 2. Wrap it in a module with executable lines before/after
     * 3. Call detectCommentedOutBlocks
     * 4. Verify the block is NOT reported
     */
    fc.assert(
      fc.property(
        fc.integer({ min: 6, max: 15 }).chain((blockLen) =>
          commentBlockArb(blockLen, false).map((block) => ({ blockLen, block })),
        ),
        ({ blockLen, block }) => {
          const lines: string[] = [];
          lines.push('x = 1'); // separator before
          for (const commentLine of block) {
            lines.push(commentLine);
          }
          lines.push('y = 2'); // separator after

          const mod = buildModule(lines);
          const findings = detectCommentedOutBlocks([mod]);

          // No findings should be reported for blocks without VB6 keywords
          expect(findings.length).toBe(0);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('mixed scenario: only qualifying blocks are reported', () => {
    /**
     * Validates: Requirements 4.4
     *
     * Strategy:
     * 1. Generate a file with multiple comment blocks of varying lengths
     * 2. Some blocks have VB6 keywords, some don't
     * 3. Verify only blocks of 6+ lines WITH keywords are reported
     */
    fc.assert(
      fc.property(
        fc.tuple(
          // A short block with keywords (should NOT be reported)
          fc.integer({ min: 1, max: 5 }).chain((len) => commentBlockArb(len, true)),
          // A long block without keywords (should NOT be reported)
          fc.integer({ min: 6, max: 10 }).chain((len) => commentBlockArb(len, false)),
          // A long block with keywords (SHOULD be reported)
          fc.integer({ min: 6, max: 10 }).chain((len) => commentBlockArb(len, true)),
        ),
        ([shortWithKw, longNoKw, longWithKw]) => {
          const lines: string[] = [];

          // Short block with keywords
          lines.push('x = 1');
          for (const c of shortWithKw) lines.push(c);

          // Separator
          lines.push('y = 2');

          // Long block without keywords
          for (const c of longNoKw) lines.push(c);

          // Separator
          lines.push('z = 3');

          // Long block with keywords — track its position
          const qualifyingStart = lines.length + 1;
          for (const c of longWithKw) lines.push(c);
          const qualifyingEnd = lines.length;

          // Separator
          lines.push('w = 4');

          const mod = buildModule(lines);
          const findings = detectCommentedOutBlocks([mod]);

          // Only the long block with keywords should be reported
          expect(findings.length).toBe(1);
          expect(findings[0].category).toBe('commented-out-block');
          expect(findings[0].startLine).toBe(qualifyingStart);
          expect(findings[0].endLine).toBe(qualifyingEnd);
        },
      ),
      { numRuns: 100 },
    );
  });
});
