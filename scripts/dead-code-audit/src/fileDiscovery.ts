/**
 * File Discovery module for the Dead Code Audit tool.
 *
 * Recursively walks a directory tree collecting VB6 source files
 * (.bas, .cls, .frm), reads their content with encoding fallback,
 * and extracts module names from Attribute VB_Name lines.
 */

import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, extname, basename, relative } from 'node:path';
import type { SourceFile } from './types.js';

const VB6_EXTENSIONS = new Set(['.bas', '.cls', '.frm']);

/** Regex to extract module name from `Attribute VB_Name = "ModuleName"` */
const VB_NAME_REGEX = /^Attribute\s+VB_Name\s*=\s*"([^"]+)"/m;

/**
 * Read a file's content, trying UTF-8 first and falling back to latin1
 * for Windows-1252 encoded files.
 */
function readFileContent(filePath: string): string {
  // Try UTF-8 first
  const utf8Content = readFileSync(filePath, 'utf-8');

  // Check for the UTF-8 replacement character which indicates decoding issues
  if (utf8Content.includes('\uFFFD')) {
    // Fall back to latin1 (superset of Windows-1252 for byte values 0x80-0xFF)
    return readFileSync(filePath, 'latin1');
  }

  return utf8Content;
}

/**
 * Extract the module name from file content by looking for the
 * `Attribute VB_Name = "..."` line. Falls back to filename without extension.
 */
function extractModuleName(content: string, filePath: string): string {
  const match = content.match(VB_NAME_REGEX);
  if (match) {
    return match[1];
  }
  return basename(filePath, extname(filePath));
}

/**
 * Determine the VB6 file type from its extension.
 */
function getFileType(filePath: string): 'bas' | 'cls' | 'frm' {
  const ext = extname(filePath).toLowerCase();
  switch (ext) {
    case '.bas': return 'bas';
    case '.cls': return 'cls';
    case '.frm': return 'frm';
    default: throw new Error(`Unexpected extension: ${ext}`);
  }
}

/**
 * Recursively walk a directory tree collecting all VB6 source files.
 *
 * @param rootDir - The root directory to scan
 * @returns Array of SourceFile objects for all discovered .bas, .cls, .frm files
 */
export function discoverFiles(rootDir: string): SourceFile[] {
  const results: SourceFile[] = [];
  walkDirectory(rootDir, rootDir, results);
  return results;
}

/**
 * Recursive directory walker that collects VB6 source files.
 */
function walkDirectory(currentDir: string, rootDir: string, results: SourceFile[]): void {
  let entries: string[];
  try {
    entries = readdirSync(currentDir);
  } catch (err) {
    console.warn(`Warning: Cannot read directory "${currentDir}": ${(err as Error).message}`);
    return;
  }

  for (const entry of entries) {
    const fullPath = join(currentDir, entry);

    let stat;
    try {
      stat = statSync(fullPath);
    } catch (err) {
      console.warn(`Warning: Cannot stat "${fullPath}": ${(err as Error).message}`);
      continue;
    }

    if (stat.isDirectory()) {
      walkDirectory(fullPath, rootDir, results);
    } else if (stat.isFile()) {
      const ext = extname(entry).toLowerCase();
      if (!VB6_EXTENSIONS.has(ext)) continue;

      try {
        const content = readFileContent(fullPath);
        const moduleName = extractModuleName(content, fullPath);
        const relativePath = relative(rootDir, fullPath);

        results.push({
          path: relativePath,
          type: getFileType(fullPath),
          moduleName,
          content,
        });
      } catch (err) {
        console.warn(`Warning: Cannot read file "${fullPath}": ${(err as Error).message}`);
      }
    }
  }
}
