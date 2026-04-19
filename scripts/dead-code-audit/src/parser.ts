/**
 * VB6 Parser — parses raw VB6 source into structured ParsedModule.
 *
 * Handles line continuations, form preamble stripping, comment/preprocessor
 * detection, string literal masking, and Attribute line classification.
 */

import type { SourceFile, ParsedLine, ParsedModule } from './types.js';

/**
 * Replace the content of string literals with a placeholder to prevent
 * false reference matches. Handles VB6 escaped quotes ("" inside strings).
 *
 * Example: `x = "hello ""world"""` ? `x = "___"`
 */
export function maskStringLiterals(line: string): string {
  // VB6 strings are delimited by double quotes.
  // Inside a string, "" is an escaped quote (literal ").
  // We walk character-by-character to handle this correctly.
  let result = '';
  let i = 0;
  while (i < line.length) {
    if (line[i] === '"') {
      // Start of a string literal — find the end
      result += '"';
      i++; // skip opening quote
      while (i < line.length) {
        if (line[i] === '"') {
          if (i + 1 < line.length && line[i + 1] === '"') {
            // Escaped quote "" — skip both
            i += 2;
          } else {
            // Closing quote
            break;
          }
        } else {
          i++;
        }
      }
      // Replace content with placeholder
      result += '___';
      if (i < line.length) {
        result += '"'; // closing quote
        i++; // skip closing quote
      }
    } else {
      result += line[i];
      i++;
    }
  }
  return result;
}

/** Preprocessor directive prefixes (case-insensitive). */
const PREPROCESSOR_PREFIXES = ['#if', '#else', '#elseif', '#end if', '#const'];

/**
 * Check if a trimmed line is a comment.
 * VB6 comments start with `'` or the keyword `Rem ` (case-insensitive).
 */
function isCommentLine(trimmed: string): boolean {
  if (trimmed.startsWith("'")) return true;
  if (/^rem(\s|$)/i.test(trimmed)) return true;
  return false;
}

/**
 * Check if a trimmed line is a preprocessor directive.
 */
function isPreprocessorLine(trimmed: string): boolean {
  const lower = trimmed.toLowerCase();
  return PREPROCESSOR_PREFIXES.some(p => lower.startsWith(p));
}

/**
 * Check if a trimmed line is an Attribute metadata line.
 */
function isAttributeLine(trimmed: string): boolean {
  return /^attribute\s/i.test(trimmed);
}

/**
 * Join physical lines connected by line continuations (`_` at end of line)
 * into logical lines, tracking original line numbers.
 *
 * Returns an array of { text, originalLines } where originalLines contains
 * all 1-based line numbers that were joined.
 */
function joinContinuations(rawLines: string[]): Array<{ text: string; originalLines: number[] }> {
  const result: Array<{ text: string; originalLines: number[] }> = [];
  let i = 0;

  while (i < rawLines.length) {
    let text = rawLines[i];
    const originalLines = [i + 1]; // 1-based

    // Check for continuation: line ends with ` _` (space then underscore)
    // or just `_` at the very end. VB6 continuation is `_` preceded by a space
    // at the end of the line.
    while (i < rawLines.length && isContinuationLine(text)) {
      // Remove the trailing ` _` and join with next line
      text = text.replace(/\s+_\s*$/, ' ');
      i++;
      if (i < rawLines.length) {
        originalLines.push(i + 1);
        text += rawLines[i].trimStart();
      }
    }

    result.push({ text, originalLines });
    i++;
  }

  return result;
}

/**
 * Check if a line ends with a continuation character.
 * VB6 continuation: line ends with ` _` (whitespace before underscore).
 */
function isContinuationLine(line: string): boolean {
  // Must end with `_` preceded by at least one whitespace character
  return /\s_\s*$/.test(line);
}

/**
 * Strip the form designer preamble from .frm files.
 * Everything before the first `Attribute VB_Name` line is form designer data.
 * Returns the lines starting from the `Attribute VB_Name` line.
 * If no `Attribute VB_Name` is found, returns all lines unchanged.
 */
function stripFormPreamble(rawLines: string[]): string[] {
  const idx = rawLines.findIndex(line =>
    /^\s*Attribute\s+VB_Name\s*=/i.test(line)
  );
  if (idx === -1) return rawLines;
  return rawLines.slice(idx);
}

/**
 * Parse a VB6 source file into a structured ParsedModule.
 */
export function parseModule(source: SourceFile): ParsedModule {
  let rawLines = source.content.split(/\r?\n/);

  // Strip form designer preamble for .frm files
  if (source.type === 'frm') {
    rawLines = stripFormPreamble(rawLines);
  }

  // Join line continuations
  const logicalLines = joinContinuations(rawLines);

  const lines: ParsedLine[] = [];
  const attributeLines: ParsedLine[] = [];

  for (const { text, originalLines } of logicalLines) {
    const masked = maskStringLiterals(text);
    const trimmed = masked.trim();

    const isComment = isCommentLine(trimmed);
    const isPreprocessor = !isComment && isPreprocessorLine(trimmed);
    const isAttribute = !isComment && !isPreprocessor && isAttributeLine(trimmed);
    const isBlank = trimmed.length === 0;
    const isExecutable = !isComment && !isPreprocessor && !isAttribute && !isBlank;

    const parsedLine: ParsedLine = {
      lineNumber: originalLines[0],
      text: masked,
      isComment,
      isPreprocessor,
      isExecutable,
      originalLines,
    };

    if (isAttribute) {
      attributeLines.push(parsedLine);
    }

    lines.push(parsedLine);
  }

  return { source, lines, attributeLines };
}
