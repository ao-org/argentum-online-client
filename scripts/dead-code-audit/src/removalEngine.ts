/**
 * Safe Removal Engine for the Dead Code Audit tool.
 *
 * Groups confirmed dead-code findings into per-module batches, applies
 * removals with syntax verification, and generates removal summaries.
 */

import * as fs from 'node:fs';
import * as path from 'node:path';
import type { Finding, RemovalBatch, RemovalResult } from './types.js';

/* ------------------------------------------------------------------ */
/*  12.1 – createBatches                                              */
/* ------------------------------------------------------------------ */

/**
 * Group confirmed, removable findings by file path into per-module batches.
 * `linesRemoved` is the count of unique source lines covered by the union
 * of all [startLine, endLine] ranges in the batch.
 *
 * Validates: Requirement 7.1
 */
export function createBatches(findings: Finding[]): RemovalBatch[] {
  // Only include confirmed + removable findings
  const eligible = findings.filter(
    (f) => f.confidence === 'confirmed' && f.removable === true,
  );

  // Group by filePath
  const grouped = new Map<string, Finding[]>();
  for (const f of eligible) {
    const list = grouped.get(f.filePath) ?? [];
    list.push(f);
    grouped.set(f.filePath, list);
  }

  const batches: RemovalBatch[] = [];
  for (const [filePath, batchFindings] of grouped) {
    // Calculate unique lines removed (union of all ranges)
    const lineSet = new Set<number>();
    for (const f of batchFindings) {
      for (let line = f.startLine; line <= f.endLine; line++) {
        lineSet.add(line);
      }
    }

    const moduleName = path.basename(filePath, path.extname(filePath));

    batches.push({
      filePath,
      moduleName,
      findings: batchFindings,
      linesRemoved: lineSet.size,
    });
  }

  return batches;
}


/* ------------------------------------------------------------------ */
/*  12.2 – validateVB6Syntax                                          */
/* ------------------------------------------------------------------ */

/**
 * Block-pair definitions: [openPattern, closePattern, label].
 *
 * Order matters – more specific patterns (e.g. "Select Case" / "End Select")
 * must be tested before generic ones ("Select").
 *
 * Validates: Requirement 7.2
 */

interface BlockPair {
  /** Regex that matches the opening keyword at the start of a trimmed line */
  open: RegExp;
  /** Regex that matches the closing keyword at the start of a trimmed line */
  close: RegExp;
  /** Human-readable label for error messages */
  label: string;
}

const BLOCK_PAIRS: BlockPair[] = [
  // Sub / End Sub
  {
    open: /^(?:(?:Public|Private|Friend)\s+)?Sub\s+/i,
    close: /^End\s+Sub\b/i,
    label: 'Sub',
  },
  // Function / End Function
  {
    open: /^(?:(?:Public|Private|Friend)\s+)?Function\s+/i,
    close: /^End\s+Function\b/i,
    label: 'Function',
  },
  // Property / End Property
  {
    open: /^(?:(?:Public|Private|Friend)\s+)?Property\s+(?:Get|Let|Set)\s+/i,
    close: /^End\s+Property\b/i,
    label: 'Property',
  },
  // Select Case / End Select
  {
    open: /^Select\s+Case\b/i,
    close: /^End\s+Select\b/i,
    label: 'Select Case',
  },
  // Type / End Type
  {
    open: /^(?:(?:Public|Private)\s+)?Type\s+/i,
    close: /^End\s+Type\b/i,
    label: 'Type',
  },
  // Enum / End Enum
  {
    open: /^(?:(?:Public|Private)\s+)?Enum\s+/i,
    close: /^End\s+Enum\b/i,
    label: 'Enum',
  },
  // If / End If  (only block-If, i.e. If ... Then on its own line)
  {
    open: /^If\b.+\bThen\s*$/i,
    close: /^End\s+If\b/i,
    label: 'If',
  },
  // For / Next
  {
    open: /^For\s+/i,
    close: /^Next\b/i,
    label: 'For',
  },
  // Do / Loop
  {
    open: /^Do\b/i,
    close: /^Loop\b/i,
    label: 'Do',
  },
  // While / Wend
  {
    open: /^While\s+/i,
    close: /^Wend\b/i,
    label: 'While',
  },
  // With / End With
  {
    open: /^With\s+/i,
    close: /^End\s+With\b/i,
    label: 'With',
  },
];

/**
 * Validate that a VB6 source string has balanced block structures.
 *
 * Returns `{ valid: true, errors: [] }` when all blocks are balanced,
 * or `{ valid: false, errors: [...] }` with descriptive messages otherwise.
 */
export function validateVB6Syntax(
  content: string,
): { valid: boolean; errors: string[] } {
  const lines = content.split(/\r?\n/);
  const stack: { label: string; line: number }[] = [];
  const errors: string[] = [];

  for (let i = 0; i < lines.length; i++) {
    const trimmed = lines[i].trim();

    // Skip blank lines and comments
    if (trimmed === '' || trimmed.startsWith("'") || /^Rem\b/i.test(trimmed)) {
      continue;
    }

    // Check closing keywords first (more specific patterns first)
    let matched = false;
    for (const pair of BLOCK_PAIRS) {
      if (pair.close.test(trimmed)) {
        matched = true;
        if (stack.length === 0) {
          errors.push(
            `Line ${i + 1}: Found '${pair.label}' closing keyword without matching opening`,
          );
        } else {
          const top = stack[stack.length - 1];
          if (top.label === pair.label) {
            stack.pop();
          } else {
            errors.push(
              `Line ${i + 1}: Expected 'End ${top.label}' but found '${pair.label}' closing keyword (opened at line ${top.line})`,
            );
            // Pop anyway to avoid cascading errors
            stack.pop();
          }
        }
        break;
      }
    }

    if (matched) continue;

    // Check opening keywords
    for (const pair of BLOCK_PAIRS) {
      if (pair.open.test(trimmed)) {
        stack.push({ label: pair.label, line: i + 1 });
        break;
      }
    }
  }

  // Any remaining unclosed blocks
  for (const unclosed of stack) {
    errors.push(
      `Unclosed '${unclosed.label}' block opened at line ${unclosed.line}`,
    );
  }

  return { valid: errors.length === 0, errors };
}


/* ------------------------------------------------------------------ */
/*  12.3 – applyBatch                                                 */
/* ------------------------------------------------------------------ */

/**
 * Apply a removal batch to a single module file.
 *
 * 1. Read the original file and create a `.bak` backup.
 * 2. Compute the union of all line ranges to remove.
 * 3. Write the modified content (lines not in the removal set).
 * 4. Validate VB6 syntax; revert from backup if invalid.
 * 5. VB6 projects have no standard test runner, so `testsPass` is `null`.
 *
 * Validates: Requirements 7.2, 7.3, 7.4, 7.5
 */
export function applyBatch(batch: RemovalBatch): RemovalResult {
  const backupPath = batch.filePath + '.bak';

  // Read original content
  let originalContent: string;
  try {
    originalContent = fs.readFileSync(batch.filePath, 'utf-8');
  } catch {
    return {
      batch,
      success: false,
      syntaxValid: false,
      testsPass: null,
      reverted: false,
    };
  }

  // Create .bak backup
  try {
    fs.writeFileSync(backupPath, originalContent, 'utf-8');
  } catch {
    return {
      batch,
      success: false,
      syntaxValid: false,
      testsPass: null,
      reverted: false,
    };
  }

  // Build set of 1-based line numbers to remove
  const linesToRemove = new Set<number>();
  for (const f of batch.findings) {
    for (let line = f.startLine; line <= f.endLine; line++) {
      linesToRemove.add(line);
    }
  }

  // Remove identified lines
  const originalLines = originalContent.split(/\r?\n/);
  const modifiedLines = originalLines.filter(
    (_, idx) => !linesToRemove.has(idx + 1),
  );
  const modifiedContent = modifiedLines.join('\n');

  // Write modified content
  try {
    fs.writeFileSync(batch.filePath, modifiedContent, 'utf-8');
  } catch {
    // Revert from backup
    try {
      fs.writeFileSync(batch.filePath, originalContent, 'utf-8');
    } catch { /* best effort */ }
    return {
      batch,
      success: false,
      syntaxValid: false,
      testsPass: null,
      reverted: true,
    };
  }

  // Validate syntax after removal
  const validation = validateVB6Syntax(modifiedContent);
  if (!validation.valid) {
    // Revert from backup
    try {
      const backup = fs.readFileSync(backupPath, 'utf-8');
      fs.writeFileSync(batch.filePath, backup, 'utf-8');
    } catch { /* best effort */ }

    // Reclassify findings as review-needed
    for (const f of batch.findings) {
      f.confidence = 'review-needed';
    }

    return {
      batch,
      success: false,
      syntaxValid: false,
      testsPass: null,
      reverted: true,
    };
  }

  // VB6 projects don't have standard test runners – set testsPass to null
  return {
    batch,
    success: true,
    syntaxValid: true,
    testsPass: null,
    reverted: false,
  };
}


/* ------------------------------------------------------------------ */
/*  12.4 – generateRemovalSummary                                     */
/* ------------------------------------------------------------------ */

/**
 * Produce a markdown summary of removal results.
 *
 * Lists lines removed per module and a total at the end.
 *
 * Validates: Requirement 7.6
 */
export function generateRemovalSummary(results: RemovalResult[]): string {
  const lines: string[] = [];
  lines.push('# Removal Summary');
  lines.push('');
  lines.push('| Module | Lines Removed | Status |');
  lines.push('|--------|--------------|--------|');

  let totalLinesRemoved = 0;

  for (const r of results) {
    const status = r.success ? 'Removed' : r.reverted ? 'Reverted' : 'Failed';
    const removed = r.success ? r.batch.linesRemoved : 0;
    totalLinesRemoved += removed;
    lines.push(
      `| ${r.batch.moduleName} | ${removed} | ${status} |`,
    );
  }

  lines.push('');
  lines.push(`**Total lines removed: ${totalLinesRemoved}**`);
  lines.push('');

  return lines.join('\n');
}
