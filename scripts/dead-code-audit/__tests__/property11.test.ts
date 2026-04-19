/**
 * Property 11: VB6 block balance validation
 *
 * For any VB6 source file, the syntax validator must report the file as valid
 * if and only if all block-opening keywords have matching block-closing keywords
 * and vice versa.
 *
 * Validates: Requirements 7.2
 */
import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { validateVB6Syntax } from '../src/removalEngine.js';

// --- Block pair definitions ---

interface BlockTemplate {
  open: string;
  close: string;
  label: string;
}

const BLOCK_TEMPLATES: BlockTemplate[] = [
  { open: 'Sub MySub()', close: 'End Sub', label: 'Sub' },
  { open: 'Function MyFunc() As Long', close: 'End Function', label: 'Function' },
  { open: 'Property Get MyProp() As String', close: 'End Property', label: 'Property' },
  { open: 'If x > 0 Then', close: 'End If', label: 'If' },
  { open: 'For i = 1 To 10', close: 'Next', label: 'For' },
  { open: 'Do While True', close: 'Loop', label: 'Do' },
  { open: 'While x > 0', close: 'Wend', label: 'While' },
  { open: 'Select Case x', close: 'End Select', label: 'Select Case' },
  { open: 'With obj', close: 'End With', label: 'With' },
  { open: 'Type MyType', close: 'End Type', label: 'Type' },
  { open: 'Enum MyEnum', close: 'End Enum', label: 'Enum' },
];

// --- Arbitraries ---

const blockTemplateArb = fc.constantFrom(...BLOCK_TEMPLATES);

/**
 * Generate a body line that won't be mistaken for a block open/close keyword.
 */
const bodyLineArb = fc.constantFrom(
  '  Dim x As Long',
  '  x = 1',
  '  Call DoSomething',
  '  Debug.Print "hello"',
  '  y = x + 1',
  "  ' This is a comment",
  '  MsgBox "test"',
);

/**
 * Generate a balanced VB6 block: open keyword, 0-3 body lines, close keyword.
 * Blocks can be nested by recursion.
 */
function balancedBlockArb(depth: number): fc.Arbitrary<string[]> {
  if (depth <= 0) {
    return fc.constant(['  x = 1']);
  }

  return blockTemplateArb.chain((template) =>
    fc.tuple(
      fc.array(bodyLineArb, { minLength: 0, maxLength: 3 }),
      fc.boolean(),
    ).chain(([bodyLines, hasNested]) => {
      if (!hasNested) {
        return fc.constant([template.open, ...bodyLines, template.close]);
      }
      // Nest an inner balanced block
      return balancedBlockArb(depth - 1).map((innerLines) => [
        template.open,
        ...bodyLines,
        ...innerLines.map(l => '  ' + l),
        template.close,
      ]);
    }),
  );
}

/**
 * Generate a full balanced VB6 source with 1-4 top-level blocks.
 */
const balancedSourceArb: fc.Arbitrary<string> = fc
  .array(balancedBlockArb(2), { minLength: 1, maxLength: 4 })
  .map((blocks) => blocks.flat().join('\n'));

/**
 * Generate an unbalanced VB6 source by taking a balanced source and
 * removing a closing keyword.
 */
const unbalancedMissingCloseArb: fc.Arbitrary<string> = fc
  .tuple(blockTemplateArb, fc.array(bodyLineArb, { minLength: 1, maxLength: 3 }))
  .map(([template, bodyLines]) =>
    // Open keyword with body but NO closing keyword
    [template.open, ...bodyLines].join('\n'),
  );

/**
 * Generate an unbalanced VB6 source with an extra closing keyword (no opener).
 */
const unbalancedExtraCloseArb: fc.Arbitrary<string> = fc
  .tuple(blockTemplateArb, fc.array(bodyLineArb, { minLength: 0, maxLength: 3 }))
  .map(([template, bodyLines]) =>
    // Closing keyword without matching open
    [...bodyLines, template.close].join('\n'),
  );

// --- Property Tests ---

describe('Feature: dead-code-audit, Property 11: VB6 block balance validation', () => {
  it('balanced VB6 blocks are reported as valid', () => {
    /**
     * Validates: Requirements 7.2
     *
     * Strategy:
     * 1. Generate VB6 source with properly balanced block structures
     * 2. Call validateVB6Syntax
     * 3. Verify result is valid with no errors
     */
    fc.assert(
      fc.property(balancedSourceArb, (source) => {
        const result = validateVB6Syntax(source);
        expect(result.valid).toBe(true);
        expect(result.errors).toHaveLength(0);
      }),
      { numRuns: 100 },
    );
  });

  it('unbalanced VB6 blocks (missing close) are reported as invalid', () => {
    /**
     * Validates: Requirements 7.2
     *
     * Strategy:
     * 1. Generate VB6 source with a missing closing keyword
     * 2. Call validateVB6Syntax
     * 3. Verify result is invalid with at least one error
     */
    fc.assert(
      fc.property(unbalancedMissingCloseArb, (source) => {
        const result = validateVB6Syntax(source);
        expect(result.valid).toBe(false);
        expect(result.errors.length).toBeGreaterThan(0);
      }),
      { numRuns: 100 },
    );
  });

  it('unbalanced VB6 blocks (extra close) are reported as invalid', () => {
    /**
     * Validates: Requirements 7.2
     *
     * Strategy:
     * 1. Generate VB6 source with an extra closing keyword (no opener)
     * 2. Call validateVB6Syntax
     * 3. Verify result is invalid with at least one error
     */
    fc.assert(
      fc.property(unbalancedExtraCloseArb, (source) => {
        const result = validateVB6Syntax(source);
        expect(result.valid).toBe(false);
        expect(result.errors.length).toBeGreaterThan(0);
      }),
      { numRuns: 100 },
    );
  });
});
