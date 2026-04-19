import { describe, it, expect } from 'vitest';
import { generateReport, computeSummary } from './reportGenerator.js';
import type { AuditReport, Finding, DuplicatePair } from './types.js';

function makeFinding(overrides: Partial<Finding> = {}): Finding {
  return {
    id: 'test::sym::1',
    category: 'unused-procedure',
    confidence: 'confirmed',
    filePath: 'Module1.bas',
    startLine: 10,
    endLine: 20,
    symbolName: 'TestSub',
    reason: 'Zero references',
    removable: true,
    ...overrides,
  };
}

function makeDuplicate(overrides: Partial<DuplicatePair> = {}): DuplicatePair {
  return {
    fileA: 'ModA.bas',
    startLineA: 1,
    endLineA: 15,
    fileB: 'ModB.bas',
    startLineB: 30,
    endLineB: 44,
    lineCount: 15,
    type: 'exact',
    ...overrides,
  };
}

function makeReport(overrides: Partial<AuditReport> = {}): AuditReport {
  const findings = overrides.findings ?? [];
  const duplicates = overrides.duplicates ?? [];
  return {
    codebase: 'client',
    timestamp: '2025-01-15T12:00:00Z',
    summary: computeSummary(findings, duplicates),
    findings,
    duplicates,
    ...overrides,
  };
}

describe('computeSummary', () => {
  it('returns all zeros for empty inputs', () => {
    const summary = computeSummary([], []);
    expect(summary).toEqual({
      unusedProcedures: 0,
      unusedVariables: 0,
      unusedConstsEnumsTypes: 0,
      unreachableCode: 0,
      commentedOutBlocks: 0,
      duplicateBlocks: 0,
    });
  });

  it('counts each category correctly', () => {
    const findings: Finding[] = [
      makeFinding({ category: 'unused-procedure' }),
      makeFinding({ category: 'unused-procedure' }),
      makeFinding({ category: 'unused-variable' }),
      makeFinding({ category: 'write-only-variable' }),
      makeFinding({ category: 'unused-const' }),
      makeFinding({ category: 'unused-enum' }),
      makeFinding({ category: 'unused-type' }),
      makeFinding({ category: 'unused-declare' }),
      makeFinding({ category: 'unreachable-code' }),
      makeFinding({ category: 'dead-branch' }),
      makeFinding({ category: 'commented-out-block' }),
    ];
    const duplicates = [makeDuplicate(), makeDuplicate()];
    const summary = computeSummary(findings, duplicates);

    expect(summary.unusedProcedures).toBe(2);
    expect(summary.unusedVariables).toBe(2); // unused-variable + write-only-variable
    expect(summary.unusedConstsEnumsTypes).toBe(4); // const + enum + type + declare
    expect(summary.unreachableCode).toBe(2); // unreachable-code + dead-branch
    expect(summary.commentedOutBlocks).toBe(1);
    expect(summary.duplicateBlocks).toBe(2);
  });
});

describe('generateReport', () => {
  it('includes title with codebase label', () => {
    const md = generateReport(makeReport({ codebase: 'client' }));
    expect(md).toContain('# Dead Code Audit Report — Client');

    const md2 = generateReport(makeReport({ codebase: 'server' }));
    expect(md2).toContain('# Dead Code Audit Report — Server');
  });

  it('includes timestamp', () => {
    const md = generateReport(makeReport());
    expect(md).toContain('**Audit performed:** 2025-01-15T12:00:00Z');
  });

  it('includes summary table with correct counts', () => {
    const findings: Finding[] = [
      makeFinding({ category: 'unused-procedure' }),
      makeFinding({ category: 'unused-variable' }),
    ];
    const report = makeReport({ findings });
    const md = generateReport(report);

    expect(md).toContain('| Unused Procedures | 1 |');
    expect(md).toContain('| Unused Variables | 1 |');
    expect(md).toContain('| Unused Constants/Enums/Types | 0 |');
  });

  it('renders finding rows with file, lines, symbol, confidence, reason', () => {
    const findings: Finding[] = [
      makeFinding({
        filePath: 'CODIGO/Protocol.bas',
        startLine: 42,
        endLine: 42,
        symbolName: 'OldHandler',
        confidence: 'review-needed',
        reason: 'Dynamic dispatch detected',
      }),
    ];
    const md = generateReport(makeReport({ findings }));
    expect(md).toContain('| CODIGO/Protocol.bas | 42 | OldHandler | review-needed | Dynamic dispatch detected |');
  });

  it('renders line range for multi-line findings', () => {
    const findings: Finding[] = [
      makeFinding({
        category: 'unreachable-code',
        startLine: 100,
        endLine: 110,
        symbolName: undefined,
      }),
    ];
    const md = generateReport(makeReport({ findings }));
    expect(md).toContain('| Module1.bas | 100-110 | — | confirmed |');
  });

  it('renders duplicate code section', () => {
    const duplicates: DuplicatePair[] = [
      makeDuplicate({
        fileA: 'A.bas',
        startLineA: 5,
        endLineA: 20,
        fileB: 'B.bas',
        startLineB: 50,
        endLineB: 65,
        lineCount: 16,
        type: 'near-duplicate',
      }),
    ];
    const md = generateReport(makeReport({ duplicates }));
    expect(md).toContain('| A.bas | 5-20 | B.bas | 50-65 | 16 | near-duplicate |');
  });

  it('shows "No findings." for empty sections', () => {
    const md = generateReport(makeReport());
    expect(md).toContain('## Unused Procedures\n\nNo findings.');
    expect(md).toContain('## Duplicate Code\n\nNo findings.');
  });

  it('has all required section headers', () => {
    const md = generateReport(makeReport());
    expect(md).toContain('## Summary');
    expect(md).toContain('## Unused Procedures');
    expect(md).toContain('## Unused Variables');
    expect(md).toContain('## Unused Constants/Enums/Types');
    expect(md).toContain('## Unreachable Code');
    expect(md).toContain('## Commented-Out Code');
    expect(md).toContain('## Duplicate Code');
  });
});
