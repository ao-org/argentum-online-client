/**
 * Report Generator — produces structured markdown audit reports.
 * 
 * Generates a markdown string from an AuditReport with sections for each
 * finding category, per-finding confidence levels, summary counts, and timestamp.
 */

import type { AuditReport, Finding, DuplicatePair, FindingCategory } from './types.js';

/**
 * Computes summary counts from findings and duplicates.
 */
export function computeSummary(
  findings: Finding[],
  duplicates: DuplicatePair[]
): AuditReport['summary'] {
  let unusedProcedures = 0;
  let unusedVariables = 0;
  let unusedConstsEnumsTypes = 0;
  let unreachableCode = 0;
  let commentedOutBlocks = 0;

  for (const f of findings) {
    switch (f.category) {
      case 'unused-procedure':
        unusedProcedures++;
        break;
      case 'unused-variable':
      case 'write-only-variable':
        unusedVariables++;
        break;
      case 'unused-const':
      case 'unused-enum':
      case 'unused-type':
      case 'unused-declare':
        unusedConstsEnumsTypes++;
        break;
      case 'unreachable-code':
      case 'dead-branch':
        unreachableCode++;
        break;
      case 'commented-out-block':
        commentedOutBlocks++;
        break;
    }
  }

  return {
    unusedProcedures,
    unusedVariables,
    unusedConstsEnumsTypes,
    unreachableCode,
    commentedOutBlocks,
    duplicateBlocks: duplicates.length,
  };
}


/** Category labels used as section headers in the report. */
const CATEGORY_SECTIONS: Record<string, FindingCategory[]> = {
  'Unused Procedures': ['unused-procedure'],
  'Unused Variables': ['unused-variable', 'write-only-variable'],
  'Unused Constants/Enums/Types': ['unused-const', 'unused-enum', 'unused-type', 'unused-declare'],
  'Unreachable Code': ['unreachable-code', 'dead-branch'],
  'Commented-Out Code': ['commented-out-block'],
};

/**
 * Generates a markdown audit report string from an AuditReport.
 * Does NOT write to disk — the CLI entry point handles file I/O.
 */
export function generateReport(report: AuditReport): string {
  const codebaseLabel = report.codebase === 'client' ? 'Client' : 'Server';
  const lines: string[] = [];

  // Title
  lines.push(`# Dead Code Audit Report — ${codebaseLabel}`);
  lines.push('');
  lines.push(`**Audit performed:** ${report.timestamp}`);
  lines.push('');

  // Summary table
  lines.push('## Summary');
  lines.push('');
  lines.push('| Category | Count |');
  lines.push('|---|---|');
  lines.push(`| Unused Procedures | ${report.summary.unusedProcedures} |`);
  lines.push(`| Unused Variables | ${report.summary.unusedVariables} |`);
  lines.push(`| Unused Constants/Enums/Types | ${report.summary.unusedConstsEnumsTypes} |`);
  lines.push(`| Unreachable Code | ${report.summary.unreachableCode} |`);
  lines.push(`| Commented-Out Code | ${report.summary.commentedOutBlocks} |`);
  lines.push(`| Duplicate Code | ${report.summary.duplicateBlocks} |`);
  lines.push('');

  // Finding sections
  for (const [sectionTitle, categories] of Object.entries(CATEGORY_SECTIONS)) {
    const sectionFindings = report.findings.filter(f => categories.includes(f.category));
    lines.push(`## ${sectionTitle}`);
    lines.push('');

    if (sectionFindings.length === 0) {
      lines.push('No findings.');
      lines.push('');
      continue;
    }

    lines.push('| File | Line(s) | Symbol | Confidence | Reason |');
    lines.push('|---|---|---|---|---|');

    for (const f of sectionFindings) {
      const lineRange = f.startLine === f.endLine
        ? `${f.startLine}`
        : `${f.startLine}-${f.endLine}`;
      const symbol = f.symbolName ?? '—';
      lines.push(`| ${f.filePath} | ${lineRange} | ${symbol} | ${f.confidence} | ${f.reason} |`);
    }

    lines.push('');
  }

  // Duplicate Code section
  lines.push('## Duplicate Code');
  lines.push('');

  if (report.duplicates.length === 0) {
    lines.push('No findings.');
    lines.push('');
  } else {
    lines.push('| File A | Lines A | File B | Lines B | Line Count | Type |');
    lines.push('|---|---|---|---|---|---|');

    for (const d of report.duplicates) {
      lines.push(
        `| ${d.fileA} | ${d.startLineA}-${d.endLineA} | ${d.fileB} | ${d.startLineB}-${d.endLineB} | ${d.lineCount} | ${d.type} |`
      );
    }

    lines.push('');
  }

  return lines.join('\n');
}
