/**
 * Dead Code Detector — identifies unused symbols, unreachable code,
 * and commented-out code blocks in VB6 source files.
 *
 * Three detection passes:
 * 1. detectUnusedSymbols: uses cross-reference usage data to find unused/write-only symbols
 * 2. detectUnreachableCode: scans procedure bodies for code after unconditional exits
 * 3. detectCommentedOutBlocks: finds large comment blocks containing VB6 keywords
 */

import type {
  SymbolUsage,
  Finding,
  FindingCategory,
  Confidence,
  ParsedModule,
  ParsedLine,
  SymbolKind,
} from './types.js';

/**
 * Map a SymbolKind to the appropriate FindingCategory for unused symbols.
 */
function unusedCategoryForKind(kind: SymbolKind): FindingCategory {
  switch (kind) {
    case 'Sub':
    case 'Function':
    case 'Property':
      return 'unused-procedure';
    case 'Variable':
      return 'unused-variable';
    case 'Const':
      return 'unused-const';
    case 'Enum':
      return 'unused-enum';
    case 'Type':
      return 'unused-type';
    case 'Declare':
      return 'unused-declare';
    default:
      return 'unused-variable';
  }
}

/**
 * Detect unused symbols and write-only variables from cross-reference usage data.
 *
 * - Skips event handlers (they are invoked by the VB6 runtime)
 * - Classifies write-only variables (writeCount > 0, readCount === 0)
 * - Sets confidence to 'review-needed' for dynamically referenced symbols
 * - Sets confidence to 'confirmed' for statically verifiable dead code
 */
export function detectUnusedSymbols(usages: SymbolUsage[]): Finding[] {
  const findings: Finding[] = [];

  for (const usage of usages) {
    const { definition } = usage;

    // Skip event handlers — they are invoked by the VB6 runtime
    if (definition.isEventHandler) continue;

    const isWriteOnly =
      definition.kind === 'Variable' &&
      usage.writeCount > 0 &&
      usage.readCount === 0;

    const isUnused = usage.totalReferences === 0;

    if (!isUnused && !isWriteOnly) continue;

    const category: FindingCategory = isWriteOnly
      ? 'write-only-variable'
      : unusedCategoryForKind(definition.kind);

    const confidence: Confidence = usage.isDynamicRef
      ? 'review-needed'
      : 'confirmed';

    const removable = confidence === 'confirmed';

    const startLine = definition.lineNumber;
    const endLine = definition.endLineNumber ?? definition.lineNumber;

    const reason = isWriteOnly
      ? `Variable '${definition.name}' is assigned but never read`
      : `${definition.kind} '${definition.name}' is declared but never referenced`;

    findings.push({
      id: `finding-${category}-${definition.filePath}-${startLine}`,
      category,
      confidence,
      filePath: definition.filePath,
      startLine,
      endLine,
      symbolName: definition.name,
      kind: definition.kind,
      visibility: definition.visibility,
      scope: definition.scope,
      dataType: definition.dataType,
      reason,
      removable,
    });
  }

  return findings;
}

// ?? Unreachable code detection ??????????????????????????????????????????

/** Regex for unconditional exit statements. */
const EXIT_PATTERN = /^(exit\s+sub|exit\s+function|exit\s+property|end)\s*$/i;

/** Regex for unconditional GoTo (not inside an If or other conditional). */
const GOTO_PATTERN = /^goto\s+\w+\s*$/i;

/** Regex for label lines (identifier followed by colon). */
const LABEL_PATTERN = /^\w+:\s*$/;

/** Regex for procedure start. */
const PROC_START = /^(public\s+|private\s+|friend\s+)?(sub|function|property\s+(get|let|set))\s+/i;

/** Regex for procedure end. */
const PROC_END = /^end\s+(sub|function|property)\s*$/i;

/** Block-end keywords that terminate unreachable regions. */
const BLOCK_END_PATTERN = /^(end\s+if|end\s+select|else|elseif\s|next\s|next$|loop|wend|end\s+with)\b/i;

/** Detect `If False Then` or `If 0 Then` dead branches. */
const DEAD_BRANCH_PATTERN = /^if\s+(false|0)\s+then\s*$/i;

/** Detect end of If block for dead branch tracking. */
const END_IF_PATTERN = /^end\s+if\s*$/i;
const ELSE_PATTERN = /^(else|elseif\s)/i;

/**
 * Detect unreachable code and dead branches in parsed modules.
 *
 * Rules:
 * - After Exit Sub/Function/Property, End, or unconditional GoTo,
 *   subsequent executable lines are unreachable until a label or block-end keyword.
 * - `If False Then` / `If 0 Then` blocks are flagged as dead branches.
 */
export function detectUnreachableCode(modules: ParsedModule[]): Finding[] {
  const findings: Finding[] = [];

  for (const mod of modules) {
    detectUnreachableInModule(mod, findings);
    detectDeadBranches(mod, findings);
  }

  return findings;
}

function detectUnreachableInModule(mod: ParsedModule, findings: Finding[]): void {
  let insideProcedure = false;
  let unreachableStart: number | null = null;
  let unreachableEnd: number | null = null;
  let unreachableReason = '';
  let afterExit = false;

  for (const line of mod.lines) {
    const trimmed = line.text.trim();

    // Track procedure boundaries
    if (PROC_START.test(trimmed)) {
      insideProcedure = true;
      afterExit = false;
      flushUnreachable(mod, findings, unreachableStart, unreachableEnd, unreachableReason);
      unreachableStart = null;
      unreachableEnd = null;
      continue;
    }

    if (PROC_END.test(trimmed)) {
      flushUnreachable(mod, findings, unreachableStart, unreachableEnd, unreachableReason);
      unreachableStart = null;
      unreachableEnd = null;
      insideProcedure = false;
      afterExit = false;
      continue;
    }

    if (!insideProcedure) continue;

    // Labels reset unreachable state
    if (LABEL_PATTERN.test(trimmed)) {
      flushUnreachable(mod, findings, unreachableStart, unreachableEnd, unreachableReason);
      unreachableStart = null;
      unreachableEnd = null;
      afterExit = false;
      continue;
    }

    // Block-end keywords reset unreachable state
    if (BLOCK_END_PATTERN.test(trimmed)) {
      flushUnreachable(mod, findings, unreachableStart, unreachableEnd, unreachableReason);
      unreachableStart = null;
      unreachableEnd = null;
      afterExit = false;
      continue;
    }

    if (!line.isExecutable) continue;

    // Check for unconditional exit
    if (EXIT_PATTERN.test(trimmed) || GOTO_PATTERN.test(trimmed)) {
      // Flush any pending unreachable block first
      flushUnreachable(mod, findings, unreachableStart, unreachableEnd, unreachableReason);
      unreachableStart = null;
      unreachableEnd = null;

      afterExit = true;
      unreachableReason = EXIT_PATTERN.test(trimmed)
        ? `Code after '${trimmed}' is unreachable`
        : `Code after unconditional '${trimmed}' is unreachable`;
      continue;
    }

    // If we're after an exit, this line is unreachable
    if (afterExit) {
      if (unreachableStart === null) {
        unreachableStart = line.lineNumber;
      }
      unreachableEnd = line.lineNumber;
    }
  }

  // Flush any remaining unreachable block
  flushUnreachable(mod, findings, unreachableStart, unreachableEnd, unreachableReason);
}

function flushUnreachable(
  mod: ParsedModule,
  findings: Finding[],
  start: number | null,
  end: number | null,
  reason: string,
): void {
  if (start !== null && end !== null) {
    findings.push({
      id: `finding-unreachable-code-${mod.source.path}-${start}`,
      category: 'unreachable-code',
      confidence: 'confirmed',
      filePath: mod.source.path,
      startLine: start,
      endLine: end,
      reason,
      removable: true,
    });
  }
}

function detectDeadBranches(mod: ParsedModule, findings: Finding[]): void {
  for (let i = 0; i < mod.lines.length; i++) {
    const line = mod.lines[i];
    if (!line.isExecutable) continue;

    const trimmed = line.text.trim();
    if (!DEAD_BRANCH_PATTERN.test(trimmed)) continue;

    // Found `If False Then` or `If 0 Then` — find the matching End If / Else
    const branchStart = line.lineNumber;
    let branchEnd = branchStart;
    let depth = 1;

    for (let j = i + 1; j < mod.lines.length; j++) {
      const inner = mod.lines[j].text.trim();

      // Track nested If blocks
      if (/^if\s+.+\s+then\s*$/i.test(inner)) {
        depth++;
      } else if (END_IF_PATTERN.test(inner)) {
        depth--;
        if (depth === 0) {
          branchEnd = mod.lines[j].lineNumber;
          break;
        }
      } else if (depth === 1 && ELSE_PATTERN.test(inner)) {
        // Else/ElseIf at our level ends the dead branch
        branchEnd = mod.lines[j - 1]?.lineNumber ?? branchStart;
        break;
      }

      branchEnd = mod.lines[j].lineNumber;
    }

    findings.push({
      id: `finding-dead-branch-${mod.source.path}-${branchStart}`,
      category: 'dead-branch',
      confidence: 'confirmed',
      filePath: mod.source.path,
      startLine: branchStart,
      endLine: branchEnd,
      reason: `Dead branch: '${trimmed}' condition is always false`,
      removable: true,
    });
  }
}

// ?? Commented-out code block detection ??????????????????????????????????

/** VB6 keywords that indicate commented-out code rather than documentation. */
const VB6_KEYWORDS = [
  'sub', 'function', 'dim', 'if', 'for', 'call', 'end',
  'select', 'case', 'do', 'loop', 'while', 'wend', 'next',
  'exit', 'goto', 'gosub', 'return', 'set', 'let', 'with',
  'property', 'public', 'private', 'const', 'enum', 'type',
  'declare', 'redim', 'erase', 'on error', 'resume',
];

/** Minimum consecutive comment lines to flag as commented-out code. */
const COMMENT_BLOCK_THRESHOLD = 6;

/**
 * Detect blocks of 6+ consecutive comment lines containing VB6 keywords.
 * These are likely commented-out code rather than documentation comments.
 */
export function detectCommentedOutBlocks(modules: ParsedModule[]): Finding[] {
  const findings: Finding[] = [];

  for (const mod of modules) {
    detectCommentedBlocksInModule(mod, findings);
  }

  return findings;
}

function detectCommentedBlocksInModule(mod: ParsedModule, findings: Finding[]): void {
  let blockStart: number | null = null;
  let blockLines: ParsedLine[] = [];

  for (const line of mod.lines) {
    if (line.isComment) {
      if (blockStart === null) {
        blockStart = line.lineNumber;
      }
      blockLines.push(line);
    } else {
      // End of comment block — check if it qualifies
      if (blockLines.length >= COMMENT_BLOCK_THRESHOLD) {
        checkAndEmitCommentBlock(mod, findings, blockStart!, blockLines);
      }
      blockStart = null;
      blockLines = [];
    }
  }

  // Check trailing block at end of file
  if (blockLines.length >= COMMENT_BLOCK_THRESHOLD && blockStart !== null) {
    checkAndEmitCommentBlock(mod, findings, blockStart, blockLines);
  }
}

function checkAndEmitCommentBlock(
  mod: ParsedModule,
  findings: Finding[],
  blockStart: number,
  blockLines: ParsedLine[],
): void {
  // Check if the block contains VB6 keywords
  const blockText = blockLines.map(l => l.text).join('\n').toLowerCase();

  const hasVB6Keywords = VB6_KEYWORDS.some(kw => {
    // Use word boundary matching to avoid false positives
    const pattern = new RegExp(`\\b${kw}\\b`, 'i');
    return pattern.test(blockText);
  });

  if (!hasVB6Keywords) return;

  const endLine = blockLines[blockLines.length - 1].lineNumber;

  findings.push({
    id: `finding-commented-out-block-${mod.source.path}-${blockStart}`,
    category: 'commented-out-block',
    confidence: 'confirmed',
    filePath: mod.source.path,
    startLine: blockStart,
    endLine,
    reason: 'Commented-out code candidate for removal',
    removable: true,
  });
}
