/**
 * Reference Scanner (Pass 2) — scans all parsed VB6 modules for symbol references.
 *
 * Tokenizes each line by splitting on VB6 operators and delimiters, matches
 * tokens against the symbol table (case-insensitive), classifies reference
 * context (call/read/write/type-usage), and tracks comment-only and dynamic
 * references.
 */

import type {
  ParsedModule,
  ParsedLine,
  SymbolTable,
  SymbolReference,
  ReferenceMap,
  SymbolDefinition,
} from './types.js';

// ?? VB6 keywords to ignore during reference scanning ?????????????????????????

const VB6_KEYWORDS = new Set([
  'as', 'new', 'byval', 'byref', 'optional', 'paramarray',
  'sub', 'function', 'property', 'end', 'if', 'then', 'else',
  'elseif', 'for', 'next', 'do', 'loop', 'while', 'wend',
  'select', 'case', 'with', 'each', 'in', 'to', 'step',
  'exit', 'goto', 'gosub', 'return', 'on', 'error', 'resume',
  'call', 'let', 'set', 'get', 'rem', 'type', 'enum',
  'declare', 'lib', 'alias', 'const', 'dim', 'redim',
  'public', 'private', 'friend', 'global', 'static',
  'attribute', 'option', 'explicit', 'compare', 'base',
  'implements', 'withevents', 'nothing', 'true', 'false',
  'not', 'and', 'or', 'xor', 'eqv', 'imp', 'mod', 'is', 'like',
  'me', 'print', 'debug', 'stop', 'boolean', 'byte', 'integer',
  'long', 'single', 'double', 'currency', 'string', 'date',
  'object', 'variant', 'any', 'preserve', 'begin', 'load',
  'unload', 'open', 'close', 'input', 'output', 'append',
  'binary', 'random', 'read', 'write', 'lock', 'unlock',
  'seek', 'put', 'name', 'kill', 'mkdir', 'rmdir', 'chdir',
  'chdrive', 'filecopy', 'mid', 'left', 'right', 'trim',
  'ltrim', 'rtrim', 'len', 'lcase', 'ucase', 'space', 'chr',
  'asc', 'val', 'str', 'cint', 'clng', 'csng', 'cdbl', 'cbool',
  'cbyte', 'cstr', 'cdate', 'cvar', 'hex', 'oct', 'fix', 'int',
  'abs', 'sgn', 'sqr', 'log', 'exp', 'rnd', 'timer',
  'msgbox', 'inputbox', 'doevents', 'typeof', 'addressof',
  'raise', 'event', 'raiseevent', 'err', 'array', 'lbound',
  'ubound', 'erase', 'isempty', 'isnull', 'isnumeric',
  'isarray', 'isobject', 'isdate', 'ismissing', 'format',
  'instr', 'instrrev', 'replace', 'split', 'join',
]);


// ?? Tokenizer ????????????????????????????????????????????????????????????????

/**
 * VB6 operator and delimiter characters used to split lines into tokens.
 * Includes: arithmetic, comparison, assignment, parentheses, commas, colons,
 * semicolons, dots (handled specially for qualified names), and whitespace.
 */
const TOKEN_SPLIT_RE = /[+\-*\/\\^&<>=(),;:\s!#]+/;

/**
 * Tokenize a VB6 line into identifier tokens.
 * Handles module-qualified references (ModuleName.SymbolName) by preserving
 * the dot-separated pair as well as emitting individual parts.
 *
 * Returns an array of tokens (identifiers only, no operators/delimiters).
 */
export function tokenizeLine(line: string): string[] {
  const tokens: string[] = [];
  // First, split by everything except dots and word chars to preserve qualified names
  // Then handle dots within those segments
  const segments = line.split(/[+\-*\/\\^&<>=(),;:\s!#]+/);

  for (const seg of segments) {
    if (!seg) continue;
    // A segment might be "ModuleName.SymbolName" or just "SymbolName"
    if (seg.includes('.')) {
      // Emit the full qualified form and individual parts
      const parts = seg.split('.');
      tokens.push(seg); // full qualified name
      for (const part of parts) {
        if (part && /^\w+$/.test(part)) {
          tokens.push(part);
        }
      }
    } else if (/^\w+$/.test(seg)) {
      tokens.push(seg);
    }
  }

  return tokens;
}

// ?? Declaration line detection ???????????????????????????????????????????????

/**
 * Regex patterns that match VB6 declaration lines.
 * Used to detect when a token on a line is the symbol being declared
 * (so we can skip counting it as a reference).
 */
const DECLARATION_PATTERNS = [
  /^(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function)\s+(\w+)/i,
  /^(Public\s+|Private\s+|Friend\s+)?Property\s+(Get|Let|Set)\s+(\w+)/i,
  /^(Public\s+|Private\s+|Friend\s+)?Declare\s+(Sub|Function)\s+(\w+)\s+Lib\s/i,
  /^(Public\s+|Private\s+|Friend\s+)?Const\s+(\w+)/i,
  /^(Public\s+|Private\s+|Friend\s+)?Enum\s+(\w+)/i,
  /^(Public\s+|Private\s+|Friend\s+)?Type\s+(\w+)/i,
];

const RE_VAR_DECL = /^(Public|Private|Dim|Global|Friend)\s+(.*)/i;

/**
 * Get the set of symbol names being declared on this line (lowercase).
 * Returns an empty set if the line is not a declaration.
 */
function getDeclaredNames(trimmedLine: string): Set<string> {
  const names = new Set<string>();

  for (const pattern of DECLARATION_PATTERNS) {
    const m = trimmedLine.match(pattern);
    if (m) {
      // The name is always the last captured group
      const groups = m.filter((_, i) => i > 0 && m[i] !== undefined);
      const name = groups[groups.length - 1];
      if (name) names.add(name.toLowerCase());
      return names;
    }
  }

  // Variable declarations (may declare multiple vars)
  const varMatch = trimmedLine.match(RE_VAR_DECL);
  if (varMatch) {
    const rest = varMatch[2];
    // Skip if it's actually a Const/Sub/Function/etc.
    if (/^(Const|Sub|Function|Declare|Enum|Type|Property|Static)\s/i.test(rest.trim())) {
      return names;
    }
    // Remove WithEvents
    const cleaned = rest.replace(/\bWithEvents\s+/gi, '');
    // Split by top-level commas
    const parts = splitByTopLevelComma(cleaned);
    for (const part of parts) {
      const nameMatch = part.trim().match(/^(\w+)/);
      if (nameMatch) {
        names.add(nameMatch[1].toLowerCase());
      }
    }
  }

  return names;
}

/**
 * Split a string by commas not inside parentheses.
 */
function splitByTopLevelComma(s: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let current = '';
  for (const ch of s) {
    if (ch === '(') { depth++; current += ch; }
    else if (ch === ')') { depth = Math.max(0, depth - 1); current += ch; }
    else if (ch === ',' && depth === 0) { parts.push(current); current = ''; }
    else { current += ch; }
  }
  if (current) parts.push(current);
  return parts;
}

// ?? Context classification ???????????????????????????????????????????????????

/**
 * Classify the reference context for a token on a given line.
 *
 * - `type-usage`: appears after `As` keyword
 * - `write`: appears on the left side of `=` (assignment) or after `Set`/`Let`
 * - `call`: appears as a procedure call (after `Call`, at start of statement, or with `(`)
 * - `read`: default — appears on right side of `=`, as argument, etc.
 */
export function classifyContext(
  token: string,
  line: string,
  symbolDefs: SymbolDefinition[] | undefined,
): 'call' | 'read' | 'write' | 'type-usage' {
  const tokenLower = token.toLowerCase();
  const lineTrimmed = line.trim();
  const lineLower = lineTrimmed.toLowerCase();

  // Type-usage: token appears after `As` keyword
  // Pattern: As <Token> or As New <Token>
  const asPattern = new RegExp(`\\bAs\\s+(?:New\\s+)?${escapeRegex(token)}\\b`, 'i');
  if (asPattern.test(lineTrimmed)) {
    return 'type-usage';
  }

  // Check if the symbol is a callable (Sub, Function, Property, Declare)
  const isCallable = symbolDefs?.some(d =>
    ['Sub', 'Function', 'Property', 'Declare'].includes(d.kind)
  ) ?? false;

  // Write: token is on the left side of `=` (but not `==`, `<=`, `>=`, `<>`)
  // Pattern: <Token> = ... (assignment)
  // Also: Set <Token> = ... or Let <Token> = ...
  const setLetPattern = new RegExp(`^(?:Set|Let)\\s+(?:\\w+\\.)?${escapeRegex(token)}\\s*=`, 'i');
  if (setLetPattern.test(lineTrimmed)) {
    return 'write';
  }

  // Simple assignment: Token = value (not comparison inside If/While/etc.)
  // Only if the token is at the start of the statement (possibly after module qualifier)
  const assignPattern = new RegExp(`^(?:\\w+\\.)?${escapeRegex(token)}\\s*\\([^)]*\\)\\s*=|^(?:\\w+\\.)?${escapeRegex(token)}\\s*=`, 'i');
  if (assignPattern.test(lineTrimmed) && !isCallable) {
    return 'write';
  }

  // Call: explicit `Call Token` pattern
  const callPattern = new RegExp(`\\bCall\\s+(?:\\w+\\.)?${escapeRegex(token)}\\b`, 'i');
  if (callPattern.test(lineTrimmed)) {
    return 'call';
  }

  // Call: token is a callable and appears at the start of a statement
  if (isCallable) {
    // Token at start of line (possibly after module qualifier)
    const startPattern = new RegExp(`^(?:\\w+\\.)?${escapeRegex(token)}\\b`, 'i');
    if (startPattern.test(lineTrimmed)) {
      return 'call';
    }

    // Token followed by opening paren (function call)
    const funcCallPattern = new RegExp(`\\b${escapeRegex(token)}\\s*\\(`, 'i');
    if (funcCallPattern.test(lineTrimmed)) {
      return 'call';
    }

    // Token used as a statement (Sub call without parens): after a colon separator
    const colonCallPattern = new RegExp(`:\\s*(?:\\w+\\.)?${escapeRegex(token)}\\b`, 'i');
    if (colonCallPattern.test(lineTrimmed)) {
      return 'call';
    }
  }

  // Default: read
  return 'read';
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ?? CallByName detection ?????????????????????????????????????????????????????

/**
 * Detect CallByName patterns in the ORIGINAL (unmasked) source content.
 * Pattern: CallByName(obj, "ProcName", ...)
 * Returns the procedure names found.
 */
export function detectCallByName(originalContent: string): string[] {
  const names: string[] = [];
  // Match CallByName(anything, "ProcName", ...) — extract the string literal
  const re = /\bCallByName\s*\([^,]+,\s*"([^"]+)"/gi;
  let match: RegExpExecArray | null;
  while ((match = re.exec(originalContent)) !== null) {
    names.push(match[1]);
  }
  return names;
}


// ?? Build declaration line index ?????????????????????????????????????????????

/**
 * Build a set of (moduleName, lineNumber) pairs for all symbol declarations.
 * Used to skip counting a symbol's own declaration line as a reference.
 */
function buildDeclarationIndex(
  symbolTable: SymbolTable,
): Map<string, Set<number>> {
  const index = new Map<string, Set<number>>();
  for (const defs of symbolTable.symbols.values()) {
    for (const def of defs) {
      let lineSet = index.get(def.moduleName);
      if (!lineSet) {
        lineSet = new Set<number>();
        index.set(def.moduleName, lineSet);
      }
      lineSet.add(def.lineNumber);
    }
  }
  return index;
}

// ?? Main scanner ?????????????????????????????????????????????????????????????

/**
 * Scan all parsed modules for symbol references, building a ReferenceMap.
 *
 * For each line (executable and comment), tokenizes the line, matches tokens
 * against the symbol table, classifies the reference context, and records
 * the reference. Skips a symbol's own declaration line.
 *
 * Also detects CallByName dynamic dispatch patterns from the original source.
 */
export function scanReferences(
  modules: ParsedModule[],
  symbolTable: SymbolTable,
): ReferenceMap {
  const references = new Map<string, SymbolReference[]>();
  const declIndex = buildDeclarationIndex(symbolTable);

  // Build a set of known module names (lowercase) for qualified reference resolution
  const moduleNames = new Set<string>();
  for (const mod of modules) {
    moduleNames.add(mod.source.moduleName.toLowerCase());
  }

  // First pass: detect CallByName patterns from original source content
  const dynamicNames = new Set<string>();
  for (const mod of modules) {
    const callByNames = detectCallByName(mod.source.content);
    for (const name of callByNames) {
      dynamicNames.add(name.toLowerCase());
    }
  }

  // Add dynamic references for CallByName targets
  for (const mod of modules) {
    const callByNameRefs = detectCallByNameWithLocation(mod);
    for (const ref of callByNameRefs) {
      const key = ref.symbolName.toLowerCase();
      // Only add if the symbol exists in the symbol table
      if (symbolTable.symbols.has(key)) {
        addReference(references, ref);
      }
    }
  }

  // Second pass: scan each line for token references
  for (const mod of modules) {
    const { source, lines } = mod;
    const moduleName = source.moduleName;
    const filePath = source.path;
    const moduleDeclLines = declIndex.get(moduleName);

    for (const line of lines) {
      // Process executable lines and comment lines
      if (!line.isExecutable && !line.isComment) continue;
      if (line.text.trim().length === 0) continue;

      const isInComment = line.isComment;

      // For comment lines, strip the leading comment marker before tokenizing
      let textToTokenize = line.text;
      if (isInComment) {
        const trimmed = line.text.trim();
        if (trimmed.startsWith("'")) {
          textToTokenize = trimmed.substring(1);
        } else if (/^rem\s/i.test(trimmed)) {
          textToTokenize = trimmed.substring(4);
        }
      }

      // Get declared names on this line to skip self-references
      const declaredOnLine = getDeclaredNames(line.text.trim());

      // Check if this line is a declaration line for any symbol
      const isDeclarationLine = moduleDeclLines?.has(line.lineNumber) ?? false;

      const tokens = tokenizeLine(textToTokenize);
      const seen = new Set<string>(); // avoid duplicate refs per line

      for (const token of tokens) {
        // Handle qualified references: ModuleName.SymbolName
        if (token.includes('.')) {
          const dotIdx = token.indexOf('.');
          const qualifier = token.substring(0, dotIdx);
          const member = token.substring(dotIdx + 1);

          if (qualifier && member && moduleNames.has(qualifier.toLowerCase())) {
            const memberLower = member.toLowerCase();
            const qualKey = `${qualifier.toLowerCase()}.${memberLower}`;

            if (!seen.has(qualKey) && symbolTable.symbols.has(memberLower)) {
              // Check it's not a self-declaration
              if (!(isDeclarationLine && declaredOnLine.has(memberLower))) {
                seen.add(qualKey);
                const defs = symbolTable.symbols.get(memberLower);
                const context = classifyContext(member, textToTokenize, defs);
                addReference(references, {
                  symbolName: memberLower,
                  referencingModule: moduleName,
                  filePath,
                  lineNumber: line.lineNumber,
                  isInComment,
                  context,
                  isDynamic: false,
                });
              }
            }
          }
          continue; // qualified token already handled; individual parts emitted separately
        }

        const tokenLower = token.toLowerCase();

        // Skip VB6 keywords
        if (VB6_KEYWORDS.has(tokenLower)) continue;

        // Skip if not in symbol table
        if (!symbolTable.symbols.has(tokenLower)) continue;

        // Skip self-declaration
        if (isDeclarationLine && declaredOnLine.has(tokenLower)) continue;

        // Skip duplicate references on the same line for the same symbol
        if (seen.has(tokenLower)) continue;
        seen.add(tokenLower);

        const defs = symbolTable.symbols.get(tokenLower);
        const context = classifyContext(token, textToTokenize, defs);
        const isDynamic = dynamicNames.has(tokenLower);

        addReference(references, {
          symbolName: tokenLower,
          referencingModule: moduleName,
          filePath,
          lineNumber: line.lineNumber,
          isInComment,
          context,
          isDynamic,
        });
      }
    }
  }

  return { references };
}

/**
 * Detect CallByName patterns with location info from original source content.
 * Returns SymbolReference entries for each CallByName target found.
 */
function detectCallByNameWithLocation(mod: ParsedModule): SymbolReference[] {
  const refs: SymbolReference[] = [];
  const rawLines = mod.source.content.split(/\r?\n/);

  for (let i = 0; i < rawLines.length; i++) {
    const rawLine = rawLines[i];
    const re = /\bCallByName\s*\([^,]+,\s*"([^"]+)"/gi;
    let match: RegExpExecArray | null;
    while ((match = re.exec(rawLine)) !== null) {
      refs.push({
        symbolName: match[1].toLowerCase(),
        referencingModule: mod.source.moduleName,
        filePath: mod.source.path,
        lineNumber: i + 1,
        isInComment: false,
        context: 'call',
        isDynamic: true,
      });
    }
  }

  return refs;
}

/**
 * Add a reference to the reference map.
 */
function addReference(
  references: Map<string, SymbolReference[]>,
  ref: SymbolReference,
): void {
  const key = ref.symbolName;
  const existing = references.get(key);
  if (existing) {
    existing.push(ref);
  } else {
    references.set(key, [ref]);
  }
}
