/**
 * Duplicate Detector — finds exact and near-duplicate code blocks
 * using Rabin-Karp rolling hash comparison.
 *
 * Requirements: 5.1, 5.2, 5.3, 5.4
 */

import type { ParsedModule, ParsedLine, DuplicatePair } from './types.js';

/** VB6 keywords that should NOT be treated as identifiers during near-duplicate normalization. */
const VB6_KEYWORDS = new Set([
  'and', 'as', 'boolean', 'byref', 'byte', 'byval', 'call', 'case', 'cbool',
  'cbyte', 'ccur', 'cdate', 'cdbl', 'cint', 'clng', 'close', 'const',
  'csng', 'cstr', 'currency', 'cvar', 'date', 'debug', 'declare', 'dim',
  'do', 'double', 'each', 'else', 'elseif', 'empty', 'end', 'enum', 'erase',
  'error', 'event', 'exit', 'false', 'for', 'friend', 'function', 'get',
  'global', 'gosub', 'goto', 'if', 'imp', 'implements', 'in', 'input',
  'integer', 'is', 'let', 'lib', 'like', 'long', 'loop', 'lset', 'me',
  'mod', 'new', 'next', 'not', 'nothing', 'null', 'object', 'on', 'open',
  'option', 'optional', 'or', 'output', 'paramarray', 'preserve', 'print',
  'private', 'property', 'public', 'raiseevent', 'redim', 'rem', 'resume',
  'return', 'rset', 'select', 'set', 'single', 'static', 'step', 'stop',
  'string', 'sub', 'then', 'to', 'true', 'type', 'typeof', 'until',
  'variant', 'wend', 'while', 'with', 'withevents', 'xor',
]);

/**
 * Normalize whitespace: collapse runs of whitespace to a single space, trim.
 */
function normalizeWhitespace(line: string): string {
  return line.replace(/\s+/g, ' ').trim();
}

/**
 * Check if a token is a numeric literal (integer, hex, float).
 */
function isNumericLiteral(token: string): boolean {
  return /^(&H[0-9A-Fa-f]+|&O[0-7]+|\d+(\.\d*)?([eE][+-]?\d+)?#?)$/.test(token);
}

/**
 * Normalize identifiers to positional placeholders for near-duplicate detection.
 * Replaces identifier tokens (not VB6 keywords, not numeric literals) with
 * positional placeholders like $1, $2, etc.
 */
function normalizeIdentifiers(lines: string[]): string[] {
  const identifierMap = new Map<string, string>();
  let nextId = 1;

  return lines.map(line => {
    return line.replace(/\b(\w+)\b/g, (match) => {
      const lower = match.toLowerCase();
      // Keep VB6 keywords and numeric literals as-is
      if (VB6_KEYWORDS.has(lower) || isNumericLiteral(match)) {
        return match.toLowerCase();
      }
      // Replace identifiers with positional placeholders
      if (!identifierMap.has(lower)) {
        identifierMap.set(lower, `$${nextId++}`);
      }
      return identifierMap.get(lower)!;
    });
  });
}


/** Info about an executable line with its original line number. */
interface ExecutableLine {
  normalized: string;
  originalLineNumber: number;
}

/** A block fingerprint for a sliding window of consecutive executable lines. */
interface BlockFingerprint {
  hash: number;
  filePath: string;
  /** Index into the executable lines array for this module */
  startIdx: number;
}

/** A structural fingerprint using identifier-normalized lines. */
interface StructuralFingerprint {
  hash: number;
  filePath: string;
  startIdx: number;
}

// Rabin-Karp rolling hash constants
// Use a large prime and base for good distribution
const HASH_BASE = 31;
const HASH_MOD = 1_000_000_007;

/**
 * Compute a simple polynomial hash for an array of strings.
 */
function computeHash(lines: string[]): number {
  let h = 0;
  for (const line of lines) {
    for (let i = 0; i < line.length; i++) {
      h = (h * HASH_BASE + line.charCodeAt(i)) % HASH_MOD;
    }
    // Add a separator character hash to distinguish line boundaries
    h = (h * HASH_BASE + 10) % HASH_MOD;
  }
  return h;
}

/**
 * Create a unique key for a pair of locations to avoid reporting duplicates.
 * Always orders by filePath then startLine so (A,B) and (B,A) produce the same key.
 */
function pairKey(
  fileA: string, startA: number,
  fileB: string, startB: number,
): string {
  if (fileA < fileB || (fileA === fileB && startA <= startB)) {
    return `${fileA}:${startA}|${fileB}:${startB}`;
  }
  return `${fileB}:${startB}|${fileA}:${startA}`;
}

/**
 * Extract executable lines from a parsed module, normalizing whitespace.
 */
function extractExecutableLines(module: ParsedModule): ExecutableLine[] {
  const result: ExecutableLine[] = [];
  for (const line of module.lines) {
    if (line.isExecutable) {
      const normalized = normalizeWhitespace(line.text);
      if (normalized.length > 0) {
        result.push({
          normalized,
          originalLineNumber: line.lineNumber,
        });
      }
    }
  }
  return result;
}

/**
 * Detect exact and near-duplicate code blocks across all modules.
 *
 * Algorithm:
 * 1. Extract and normalize executable lines from each module
 * 2. Compute rolling hashes over sliding windows of `minLines` consecutive lines
 * 3. Group windows by hash to find collision candidates
 * 4. Verify exact duplicates by character comparison
 * 5. For non-exact matches, normalize identifiers and re-compare for near-duplicates
 *
 * @param modules - Parsed VB6 modules to scan
 * @param minLines - Minimum consecutive executable lines to consider (default 10)
 * @returns Array of duplicate pairs found
 */
export function detectDuplicates(
  modules: ParsedModule[],
  minLines: number = 10,
): DuplicatePair[] {
  // Step 1: Extract executable lines per module
  const moduleLines = new Map<string, ExecutableLine[]>();
  for (const mod of modules) {
    const execLines = extractExecutableLines(mod);
    if (execLines.length >= minLines) {
      moduleLines.set(mod.source.path, execLines);
    }
  }

  // Step 2: Compute fingerprints for exact matches and structural (identifier-normalized) matches
  const exactBuckets = new Map<number, BlockFingerprint[]>();
  const structBuckets = new Map<number, StructuralFingerprint[]>();

  for (const [filePath, execLines] of moduleLines) {
    const windowCount = execLines.length - minLines + 1;
    for (let i = 0; i < windowCount; i++) {
      const windowLines = execLines.slice(i, i + minLines).map(l => l.normalized);
      const exactHash = computeHash(windowLines);

      const fp: BlockFingerprint = { hash: exactHash, filePath, startIdx: i };
      const bucket = exactBuckets.get(exactHash);
      if (bucket) {
        bucket.push(fp);
      } else {
        exactBuckets.set(exactHash, [fp]);
      }

      // Also compute structural hash for near-duplicate detection
      const structLines = normalizeIdentifiers(windowLines);
      const structHash = computeHash(structLines);
      const sfp: StructuralFingerprint = { hash: structHash, filePath, startIdx: i };
      const sBucket = structBuckets.get(structHash);
      if (sBucket) {
        sBucket.push(sfp);
      } else {
        structBuckets.set(structHash, [sfp]);
      }
    }
  }

  const results: DuplicatePair[] = [];
  const seenPairs = new Set<string>();

  /**
   * Helper: check a pair of fingerprints, classify as exact or near-duplicate.
   */
  function checkPair(
    a: { filePath: string; startIdx: number },
    b: { filePath: string; startIdx: number },
  ): void {
    // Skip self-overlapping blocks in the same file
    if (a.filePath === b.filePath) {
      const overlap =
        (a.startIdx < b.startIdx + minLines) &&
        (b.startIdx < a.startIdx + minLines);
      if (overlap) return;
    }

    const linesA = moduleLines.get(a.filePath)!;
    const linesB = moduleLines.get(b.filePath)!;

    const blockA = linesA.slice(a.startIdx, a.startIdx + minLines);
    const blockB = linesB.slice(b.startIdx, b.startIdx + minLines);

    const startLineA = blockA[0].originalLineNumber;
    const endLineA = blockA[blockA.length - 1].originalLineNumber;
    const startLineB = blockB[0].originalLineNumber;
    const endLineB = blockB[blockB.length - 1].originalLineNumber;

    const key = pairKey(a.filePath, startLineA, b.filePath, startLineB);
    if (seenPairs.has(key)) return;

    // Exact match: compare normalized text line by line
    const normalizedA = blockA.map(l => l.normalized);
    const normalizedB = blockB.map(l => l.normalized);

    const isExact = normalizedA.every((line, idx) => line === normalizedB[idx]);

    if (isExact) {
      seenPairs.add(key);
      results.push({
        fileA: a.filePath,
        startLineA,
        endLineA,
        fileB: b.filePath,
        startLineB,
        endLineB,
        lineCount: minLines,
        type: 'exact',
      });
      return;
    }

    // Near-duplicate: normalize identifiers and re-compare
    const identNormA = normalizeIdentifiers(normalizedA);
    const identNormB = normalizeIdentifiers(normalizedB);

    const isNearDuplicate = identNormA.every(
      (line, idx) => line === identNormB[idx],
    );

    if (isNearDuplicate) {
      seenPairs.add(key);
      results.push({
        fileA: a.filePath,
        startLineA,
        endLineA,
        fileB: b.filePath,
        startLineB,
        endLineB,
        lineCount: minLines,
        type: 'near-duplicate',
      });
    }
  }

  // Step 3: Check exact hash collisions
  for (const [, bucket] of exactBuckets) {
    if (bucket.length < 2) continue;
    for (let i = 0; i < bucket.length; i++) {
      for (let j = i + 1; j < bucket.length; j++) {
        checkPair(bucket[i], bucket[j]);
      }
    }
  }

  // Step 4: Check structural hash collisions for near-duplicates
  for (const [, bucket] of structBuckets) {
    if (bucket.length < 2) continue;
    for (let i = 0; i < bucket.length; i++) {
      for (let j = i + 1; j < bucket.length; j++) {
        checkPair(bucket[i], bucket[j]);
      }
    }
  }

  return results;
}
