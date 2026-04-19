/**
 * Cross-Reference Analyzer — merges symbol definitions with references
 * to compute per-symbol usage statistics.
 *
 * Normalizes symbol names to lowercase for matching (VB6 is case-insensitive).
 * Handles ambiguous references (same name in multiple modules) by counting
 * the reference for all matching definitions.
 */

import type {
  SymbolTable,
  ReferenceMap,
  SymbolUsage,
  SymbolDefinition,
  SymbolReference,
} from './types.js';

/**
 * Analyze usage of all symbols by merging definitions with references.
 *
 * For each symbol definition in the symbol table, looks up references by
 * lowercase name in the reference map and computes:
 * - totalReferences: count of non-comment references
 * - intraModuleRefs: references from the same module as the definition
 * - crossModuleRefs: references from a different module
 * - commentOnlyRefs: references that appear only in comments
 * - writeCount: references with context === 'write'
 * - readCount: references with context === 'read'
 * - isDynamicRef: true if any reference has isDynamic === true
 *
 * Ambiguous references (same name defined in multiple modules) are counted
 * for ALL matching definitions.
 */
export function analyzeUsage(
  symbolTable: SymbolTable,
  referenceMap: ReferenceMap,
): SymbolUsage[] {
  const usages: SymbolUsage[] = [];

  // Iterate over every definition in the symbol table
  for (const [nameLower, definitions] of symbolTable.symbols.entries()) {
    // Look up all references for this symbol name (already lowercase-keyed)
    const refs = referenceMap.references.get(nameLower) ?? [];

    // For each definition with this name, compute usage stats
    for (const definition of definitions) {
      const usage = computeUsageForDefinition(definition, refs);
      usages.push(usage);
    }
  }

  return usages;
}


/**
 * Compute usage statistics for a single symbol definition given all
 * references that share its lowercase name.
 *
 * For ambiguous names (defined in multiple modules), each definition
 * gets the full set of references counted — this is intentional since
 * VB6 unqualified references could resolve to any matching definition.
 */
function computeUsageForDefinition(
  definition: SymbolDefinition,
  refs: SymbolReference[],
): SymbolUsage {
  let intraModuleRefs = 0;
  let crossModuleRefs = 0;
  let commentOnlyRefs = 0;
  let writeCount = 0;
  let readCount = 0;
  let isDynamicRef = false;

  for (const ref of refs) {
    if (ref.isInComment) {
      // Count comment-only references separately
      commentOnlyRefs++;
      continue;
    }

    // Non-comment (executable) reference
    if (ref.referencingModule === definition.moduleName) {
      intraModuleRefs++;
    } else {
      crossModuleRefs++;
    }

    if (ref.context === 'write') {
      writeCount++;
    } else if (ref.context === 'read') {
      readCount++;
    }

    if (ref.isDynamic) {
      isDynamicRef = true;
    }
  }

  // totalReferences counts only non-comment (executable) references
  const totalReferences = intraModuleRefs + crossModuleRefs;

  return {
    definition,
    totalReferences,
    intraModuleRefs,
    crossModuleRefs,
    commentOnlyRefs,
    writeCount,
    readCount,
    isDynamicRef,
  };
}
