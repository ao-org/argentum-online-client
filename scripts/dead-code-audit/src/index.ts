#!/usr/bin/env node
/**
 * Dead Code Audit CLI
 *
 * Static analysis tool for detecting dead code in VB6 codebases.
 * Usage: tsx src/index.ts --root <path-to-codebase> [--mode audit|remove] [--codebase client|server]
 */

import { writeFileSync, mkdirSync } from 'node:fs';
import * as path from 'node:path';

import { discoverFiles } from './fileDiscovery.js';
import { parseModule } from './parser.js';
import { extractSymbols, buildSymbolTable } from './symbolExtractor.js';
import { scanReferences } from './referenceScanner.js';
import { analyzeUsage } from './crossRefAnalyzer.js';
import {
  detectUnusedSymbols,
  detectUnreachableCode,
  detectCommentedOutBlocks,
} from './deadCodeDetector.js';
import { detectDuplicates } from './duplicateDetector.js';
import { computeSummary, generateReport } from './reportGenerator.js';
import {
  createBatches,
  applyBatch,
  generateRemovalSummary,
} from './removalEngine.js';
import type { AuditReport, Finding, ParsedModule, SymbolDefinition } from './types.js';

/* ------------------------------------------------------------------ */
/*  CLI argument parsing                                              */
/* ------------------------------------------------------------------ */

interface CliArgs {
  root: string | null;
  mode: 'audit' | 'remove';
  codebase: 'client' | 'server' | null;
}

function parseArgs(args: string[]): CliArgs {
  let root: string | null = null;
  let mode: 'audit' | 'remove' = 'audit';
  let codebase: 'client' | 'server' | null = null;

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--root' && i + 1 < args.length) {
      root = args[i + 1];
      i++;
    } else if (args[i] === '--mode' && i + 1 < args.length) {
      const val = args[i + 1];
      if (val === 'audit' || val === 'remove') {
        mode = val;
      }
      i++;
    } else if (args[i] === '--codebase' && i + 1 < args.length) {
      const val = args[i + 1];
      if (val === 'client' || val === 'server') {
        codebase = val;
      }
      i++;
    }
  }

  return { root, mode, codebase };
}


/* ------------------------------------------------------------------ */
/*  Codebase detection                                                */
/* ------------------------------------------------------------------ */

function detectCodebase(rootPath: string): 'client' | 'server' {
  const normalized = rootPath.replace(/\\/g, '/');
  if (normalized.includes('argentum-online-server')) return 'server';
  return 'client';
}

/* ------------------------------------------------------------------ */
/*  Report output path                                                */
/* ------------------------------------------------------------------ */

function getReportPath(codebase: 'client' | 'server', rootDir: string): string {
  // Walk up from the --root dir to find the workspace root (parent of CODIGO/Codigo)
  const parentDir = path.dirname(rootDir);
  return path.join(parentDir, '.kiro', 'specs', 'dead-code-audit', 'audit-report.md');
}

/* ------------------------------------------------------------------ */
/*  Main pipeline                                                     */
/* ------------------------------------------------------------------ */

function main(): void {
  const { root, mode, codebase: codebaseOverride } = parseArgs(process.argv.slice(2));

  if (!root) {
    console.error('Usage: tsx src/index.ts --root <path-to-codebase> [--mode audit|remove] [--codebase client|server]');
    process.exit(1);
  }

  const resolvedRoot = path.resolve(root);
  const codebase = codebaseOverride ?? detectCodebase(resolvedRoot);

  console.log(`Dead Code Audit`);
  console.log(`  Root: ${resolvedRoot}`);
  console.log(`  Mode: ${mode}`);
  console.log(`  Codebase: ${codebase}`);
  console.log('');

  // 1. Discover VB6 files
  console.log('Discovering VB6 files...');
  const sourceFiles = discoverFiles(resolvedRoot);
  console.log(`  Found ${sourceFiles.length} files`);

  if (sourceFiles.length === 0) {
    console.log('No VB6 files found. Nothing to audit.');
    return;
  }

  // 2. Parse each file
  console.log('Parsing modules...');
  const modules: ParsedModule[] = sourceFiles.map((sf) => parseModule(sf));
  console.log(`  Parsed ${modules.length} modules`);

  // 3. Extract symbols and build symbol table
  console.log('Extracting symbols...');
  const allDefinitions: SymbolDefinition[] = [];
  for (const mod of modules) {
    const defs = extractSymbols(mod);
    allDefinitions.push(...defs);
  }
  const symbolTable = buildSymbolTable(allDefinitions);
  console.log(`  Extracted ${allDefinitions.length} symbol definitions`);

  // 4. Scan references
  console.log('Scanning references...');
  const referenceMap = scanReferences(modules, symbolTable);
  const totalRefs = Array.from(referenceMap.references.values()).reduce(
    (sum, refs) => sum + refs.length,
    0,
  );
  console.log(`  Found ${totalRefs} references`);

  // 5. Analyze usage
  console.log('Analyzing usage...');
  const usages = analyzeUsage(symbolTable, referenceMap);

  // 6. Detect dead code
  console.log('Detecting dead code...');
  const unusedFindings = detectUnusedSymbols(usages);
  const unreachableFindings = detectUnreachableCode(modules);
  const commentedOutFindings = detectCommentedOutBlocks(modules);

  const allFindings: Finding[] = [
    ...unusedFindings,
    ...unreachableFindings,
    ...commentedOutFindings,
  ];
  console.log(`  Unused symbols: ${unusedFindings.length}`);
  console.log(`  Unreachable code: ${unreachableFindings.length}`);
  console.log(`  Commented-out blocks: ${commentedOutFindings.length}`);

  // 7. Detect duplicates
  console.log('Detecting duplicates...');
  const duplicates = detectDuplicates(modules, 15);
  console.log(`  Duplicate pairs: ${duplicates.length}`);

  // Cap duplicates in the report to the top 200 by line count
  const sortedDuplicates = duplicates
    .sort((a, b) => b.lineCount - a.lineCount)
    .slice(0, 200);
  if (duplicates.length > 200) {
    console.log(`  (showing top 200 of ${duplicates.length} in report)`);
  }

  // 8. Build report
  console.log('Generating report...');
  const summary = computeSummary(allFindings, sortedDuplicates);
  const report: AuditReport = {
    codebase,
    timestamp: new Date().toISOString(),
    summary,
    findings: allFindings,
    duplicates: sortedDuplicates,
  };

  const reportMarkdown = generateReport(report);

  // 9. Write report to disk
  const reportPath = getReportPath(codebase, resolvedRoot);
  mkdirSync(path.dirname(reportPath), { recursive: true });
  writeFileSync(reportPath, reportMarkdown, 'utf-8');
  console.log(`\nReport written to: ${reportPath}`);

  // 10. If remove mode, apply safe removal
  if (mode === 'remove') {
    console.log('\nRunning safe removal engine...');
    const batches = createBatches(allFindings);
    console.log(`  Created ${batches.length} removal batches`);

    const results = batches.map((batch) => {
      const result = applyBatch(batch);
      const status = result.success ? 'OK' : result.reverted ? 'Reverted' : 'Failed';
      console.log(`  ${batch.moduleName}: ${status} (${batch.linesRemoved} lines)`);
      return result;
    });

    const removalSummary = generateRemovalSummary(results);
    console.log('\n' + removalSummary);
  }

  // 11. Print summary to stdout
  console.log('\nAudit Summary:');
  console.log(`  Unused procedures: ${summary.unusedProcedures}`);
  console.log(`  Unused variables: ${summary.unusedVariables}`);
  console.log(`  Unused constants/enums/types: ${summary.unusedConstsEnumsTypes}`);
  console.log(`  Unreachable code: ${summary.unreachableCode}`);
  console.log(`  Commented-out blocks: ${summary.commentedOutBlocks}`);
  console.log(`  Duplicate blocks: ${summary.duplicateBlocks}`);
}

main();
