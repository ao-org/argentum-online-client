/**
 * Shared TypeScript interfaces for the Dead Code Audit tool.
 * These types model the VB6 static analysis pipeline from file discovery
 * through report generation and safe removal.
 */

/** A discovered VB6 source file. */
export interface SourceFile {
  /** Relative path from project root */
  path: string;
  /** File type determined by extension */
  type: 'bas' | 'cls' | 'frm';
  /** Extracted from Attribute VB_Name or filename */
  moduleName: string;
  /** Raw file content */
  content: string;
}

/** A single logical line after continuation joining. */
export interface ParsedLine {
  /** Original 1-based line number */
  lineNumber: number;
  /** Logical line after continuation joining */
  text: string;
  isComment: boolean;
  isPreprocessor: boolean;
  /** Not a comment, not blank, not preprocessor, not Attribute */
  isExecutable: boolean;
  /** All original line numbers this logical line spans */
  originalLines: number[];
}

/** A fully parsed VB6 module. */
export interface ParsedModule {
  source: SourceFile;
  lines: ParsedLine[];
  /** Attribute VB_Name, etc. */
  attributeLines: ParsedLine[];
}

export type SymbolKind = 'Sub' | 'Function' | 'Property' | 'Variable' | 'Const' | 'Enum' | 'Type' | 'Declare' | 'Event';
export type Visibility = 'Public' | 'Private' | 'Friend';
export type VariableScope = 'module' | 'local';


/** A single symbol definition extracted from VB6 source. */
export interface SymbolDefinition {
  /** Unique: `${moduleName}::${name}::${lineNumber}` */
  id: string;
  name: string;
  kind: SymbolKind;
  visibility: Visibility;
  moduleName: string;
  filePath: string;
  lineNumber: number;
  /** For multi-line symbols (Type, Enum, Sub, Function) */
  endLineNumber?: number;
  scope: VariableScope;
  /** For variables and constants */
  dataType?: string;
  /** For local variables, the containing procedure name */
  parentProcedure?: string;
  /** True for Form_Load, cmdX_Click, etc. */
  isEventHandler: boolean;
}

/** Symbol table mapping names and modules to definitions. */
export interface SymbolTable {
  /** name ? definitions (may have duplicates across modules) */
  symbols: Map<string, SymbolDefinition[]>;
  /** moduleName ? definitions in that module */
  byModule: Map<string, SymbolDefinition[]>;
}

/** A single reference to a symbol found in source code. */
export interface SymbolReference {
  /** Case-normalized identifier */
  symbolName: string;
  /** Module where the reference occurs */
  referencingModule: string;
  filePath: string;
  lineNumber: number;
  /** True if reference is in a commented-out line */
  isInComment: boolean;
  /** How the symbol is used */
  context: 'call' | 'read' | 'write' | 'type-usage';
  /** True if via CallByName or similar */
  isDynamic: boolean;
}

/** Map of symbol names to all their references. */
export interface ReferenceMap {
  /** symbolName ? all references */
  references: Map<string, SymbolReference[]>;
}

/** Per-symbol usage statistics after cross-reference analysis. */
export interface SymbolUsage {
  definition: SymbolDefinition;
  totalReferences: number;
  intraModuleRefs: number;
  crossModuleRefs: number;
  /** References that appear only in comments */
  commentOnlyRefs: number;
  /** Times written to (for variables) */
  writeCount: number;
  /** Times read from (for variables) */
  readCount: number;
  /** Referenced via CallByName or similar */
  isDynamicRef: boolean;
}

export type FindingCategory =
  | 'unused-procedure'
  | 'unused-variable'
  | 'unused-const'
  | 'unused-enum'
  | 'unused-type'
  | 'unused-declare'
  | 'write-only-variable'
  | 'unreachable-code'
  | 'commented-out-block'
  | 'dead-branch';

export type Confidence = 'confirmed' | 'review-needed';

/** A single dead code finding. */
export interface Finding {
  /** Unique finding identifier */
  id: string;
  category: FindingCategory;
  confidence: Confidence;
  filePath: string;
  startLine: number;
  endLine: number;
  symbolName?: string;
  kind?: SymbolKind;
  visibility?: Visibility;
  scope?: VariableScope;
  dataType?: string;
  /** Human-readable explanation */
  reason: string;
  /** Whether safe auto-removal is possible */
  removable: boolean;
}

/** A pair of duplicate or near-duplicate code blocks. */
export interface DuplicatePair {
  fileA: string;
  startLineA: number;
  endLineA: number;
  fileB: string;
  startLineB: number;
  endLineB: number;
  lineCount: number;
  type: 'exact' | 'near-duplicate';
}

/** The full audit report for a codebase. */
export interface AuditReport {
  codebase: 'client' | 'server';
  timestamp: string;
  summary: {
    unusedProcedures: number;
    unusedVariables: number;
    unusedConstsEnumsTypes: number;
    unreachableCode: number;
    commentedOutBlocks: number;
    duplicateBlocks: number;
  };
  findings: Finding[];
  duplicates: DuplicatePair[];
}

/** A batch of removals for a single module file. */
export interface RemovalBatch {
  filePath: string;
  moduleName: string;
  findings: Finding[];
  linesRemoved: number;
}

/** Result of applying a removal batch. */
export interface RemovalResult {
  batch: RemovalBatch;
  success: boolean;
  syntaxValid: boolean;
  /** null if no tests exist for this module */
  testsPass: boolean | null;
  reverted: boolean;
}

/** Tracks parser state while processing a VB6 module. */
export interface ProcessingState {
  /** Tracks which Sub/Function we're inside */
  currentProcedure: string | null;
  /** Stack of open blocks (Sub, If, For, etc.) */
  blockStack: string[];
  /** Inside a Type...End Type block */
  insideType: boolean;
  /** Inside an Enum...End Enum block */
  insideEnum: boolean;
}
