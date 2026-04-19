/**
 * Symbol Extractor (Pass 1) — extracts all symbol definitions from parsed VB6 modules.
 *
 * Recognizes: Sub, Function, Property Get/Let/Set, module-level variables,
 * local variables, Const, Enum, Type, API Declare, and event handlers.
 * Tracks ProcessingState to distinguish module-level vs local scope.
 */

import type {
  ParsedModule,
  ParsedLine,
  SymbolDefinition,
  SymbolKind,
  Visibility,
  VariableScope,
  SymbolTable,
  ProcessingState,
} from './types.js';

// ?? Regex patterns for VB6 declarations ??????????????????????????????????????

/**
 * Sub / Function declaration.
 * Captures: [1] visibility (optional), [2] Static (optional), [3] Sub|Function, [4] name, [5] return type (optional)
 * Examples:
 *   Public Sub Init()
 *   Private Function GetValue() As Long
 *   Sub DoStuff()
 *   Private Static Function Calc() As Long
 *   Friend Sub Helper()
 */
const RE_PROC = /^(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function)\s+(\w+)/i;

/**
 * Property Get/Let/Set declaration.
 * Captures: [1] visibility (optional), [2] Get|Let|Set, [3] name
 * Examples:
 *   Public Property Get Name() As String
 *   Property Let Name(val As String)
 *   Friend Property Set Obj(val As Object)
 */
const RE_PROPERTY = /^(Public\s+|Private\s+|Friend\s+)?Property\s+(Get|Let|Set)\s+(\w+)/i;

/**
 * API Declare statement.
 * Captures: [1] visibility (optional), [2] Sub|Function, [3] name
 * Examples:
 *   Private Declare Function SendMessage Lib "user32" ...
 *   Public Declare Sub Sleep Lib "kernel32" ...
 *   Declare Function CallNextHookEx Lib "user32" ...
 */
const RE_DECLARE = /^(Public\s+|Private\s+|Friend\s+)?Declare\s+(Sub|Function)\s+(\w+)\s+Lib\s/i;

/**
 * Const declaration.
 * Captures: [1] visibility (optional), [2] name, [3] type (optional via As)
 * Examples:
 *   Public Const NO_WEAPON As Byte = 2
 *   Private Const MAX_ITEMS = 100
 *   Const MY_VAL As Long = 5
 */
const RE_CONST = /^(Public\s+|Private\s+|Friend\s+)?Const\s+(\w+)(?:\s+As\s+(\w+))?/i;

/**
 * Enum declaration.
 * Captures: [1] visibility (optional), [2] name
 * Examples:
 *   Public Enum tMacro
 *   Private Enum eDirection
 *   Enum SomeEnum
 */
const RE_ENUM = /^(Public\s+|Private\s+|Friend\s+)?Enum\s+(\w+)/i;

/**
 * Type (UDT) declaration.
 * Captures: [1] visibility (optional), [2] name
 * Examples:
 *   Private Type Position
 *   Public Type tUser
 *   Type SomeType
 */
const RE_TYPE = /^(Public\s+|Private\s+|Friend\s+)?Type\s+(\w+)/i;

/**
 * Variable declaration at module level or local scope.
 * Handles Public, Private, Dim, Global, and Friend.
 * We parse the rest of the line to extract individual variables
 * (VB6 allows `Dim x As Long, y As String`).
 *
 * Captures: [1] keyword (Public|Private|Dim|Global|Friend), [2] rest of line
 * Examples:
 *   Public bSkins As Boolean
 *   Private mCount As Long
 *   Dim gData As String
 *   Dim x As Long, y As String
 *   Public WithEvents tmrTimer As Timer
 *   Global SomeVar As Integer
 */
const RE_VAR_LINE = /^(Public|Private|Dim|Global|Friend)\s+(.*)/i;

/** End-block patterns for tracking procedure/block boundaries */
const RE_END_SUB = /^End\s+Sub\b/i;
const RE_END_FUNCTION = /^End\s+Function\b/i;
const RE_END_PROPERTY = /^End\s+Property\b/i;
const RE_END_TYPE = /^End\s+Type\b/i;
const RE_END_ENUM = /^End\s+Enum\b/i;

/**
 * Event handler pattern — matches `ControlName_EventName` in a Private Sub.
 * Common VB6 events: Click, DblClick, Change, Load, Unload, Activate, Deactivate,
 * MouseMove, MouseDown, MouseUp, KeyPress, KeyDown, KeyUp, GotFocus, LostFocus,
 * Timer, Resize, Paint, Initialize, Terminate, Scroll, etc.
 *
 * Pattern: name contains underscore and the part after the last underscore
 * is a known VB6 event name, OR the prefix is a known control-like pattern
 * (Form_, cmd_, txt_, lst_, img_, pic_, tmr_, Timer, etc.)
 */
const VB6_EVENT_NAMES = new Set([
  'click', 'dblclick', 'change', 'load', 'unload', 'activate', 'deactivate',
  'mousedown', 'mousemove', 'mouseup', 'keypress', 'keydown', 'keyup',
  'gotfocus', 'lostfocus', 'timer', 'resize', 'paint', 'initialize',
  'terminate', 'scroll', 'validate', 'dropdown', 'itemclick', 'collapse',
  'expand', 'nodeclicked', 'beforeupdate', 'afterupdate', 'enter', 'exit',
  'dragdrop', 'dragover', 'linkclose', 'linkopen', 'linkerror',
  'linkexecute', 'linknotify', 'oledragdrop', 'oledragover',
  'olegivefeedback', 'olesetdata', 'olestartdrag', 'olecompletedrag',
  'queryunload', 'close', 'connect', 'dataarrival', 'error',
  'sendcomplete', 'sendprogress', 'connectionrequest', 'statechanged',
]);


// ?? Helpers ??????????????????????????????????????????????????????????????????

/**
 * Determine if a Sub name looks like a VB6 event handler.
 * Event handlers follow the pattern: `ControlName_EventName`
 * where EventName is a known VB6 event.
 */
function isEventHandlerName(name: string): boolean {
  const underscoreIdx = name.lastIndexOf('_');
  if (underscoreIdx <= 0 || underscoreIdx === name.length - 1) return false;

  const eventPart = name.substring(underscoreIdx + 1).toLowerCase();
  return VB6_EVENT_NAMES.has(eventPart);
}

/**
 * Parse the visibility keyword from a captured group.
 * Returns the normalized visibility and whether it was explicitly specified.
 */
function parseVisibility(raw: string | undefined): { visibility: Visibility; explicit: boolean } {
  if (!raw) return { visibility: 'Public', explicit: false };
  const trimmed = raw.trim().toLowerCase();
  if (trimmed === 'private') return { visibility: 'Private', explicit: true };
  if (trimmed === 'friend') return { visibility: 'Friend', explicit: true };
  return { visibility: 'Public', explicit: true };
}

/**
 * Parse the variable keyword to determine default visibility.
 * - `Public` / `Global` ? Public
 * - `Private` ? Private
 * - `Dim` at module level ? Private
 * - `Dim` inside procedure ? local (handled by caller)
 * - `Friend` ? Friend
 */
function varKeywordVisibility(keyword: string): Visibility {
  const lower = keyword.toLowerCase();
  if (lower === 'public' || lower === 'global') return 'Public';
  if (lower === 'private') return 'Private';
  if (lower === 'friend') return 'Friend';
  // Dim at module level defaults to Private
  return 'Private';
}

/**
 * Generate a unique symbol ID.
 */
function makeId(moduleName: string, name: string, lineNumber: number): string {
  return `${moduleName}::${name}::${lineNumber}`;
}

/**
 * Parse a variable declaration list (the part after Public/Private/Dim/etc.).
 * Handles:
 *   - `x As Long`
 *   - `x As Long, y As String`
 *   - `WithEvents tmr As Timer`
 *   - `x(1 To 10) As Byte`
 *   - `x As New Collection`
 *   - `x` (no type specified)
 *   - `x$` (type suffix)
 *
 * Returns array of { name, dataType } for each variable found.
 */
function parseVariableList(rest: string): Array<{ name: string; dataType?: string }> {
  const results: Array<{ name: string; dataType?: string }> = [];

  // Remove WithEvents keyword if present
  let cleaned = rest.replace(/\bWithEvents\s+/gi, '');

  // Split by comma, but be careful about commas inside parentheses (array bounds)
  const parts = splitByTopLevelComma(cleaned);

  for (const part of parts) {
    const trimmed = part.trim();
    if (!trimmed) continue;

    // Match: varName[(bounds)] [As [New] TypeName]
    const match = trimmed.match(/^(\w+)\s*(?:\([^)]*\))?\s*(?:As\s+(?:New\s+)?(\w+))?/i);
    if (match) {
      const name = match[1];
      const dataType = match[2];
      // Skip if the "name" is a keyword that leaked through
      if (!isVB6Keyword(name)) {
        results.push({ name, dataType });
      }
    }
  }

  return results;
}

/**
 * Split a string by commas that are not inside parentheses.
 */
function splitByTopLevelComma(s: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let current = '';

  for (const ch of s) {
    if (ch === '(') {
      depth++;
      current += ch;
    } else if (ch === ')') {
      depth = Math.max(0, depth - 1);
      current += ch;
    } else if (ch === ',' && depth === 0) {
      parts.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  if (current) parts.push(current);
  return parts;
}

/** VB6 keywords that should not be treated as variable names. */
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
]);

function isVB6Keyword(name: string): boolean {
  return VB6_KEYWORDS.has(name.toLowerCase());
}


// ?? Main extraction logic ????????????????????????????????????????????????????

/**
 * Extract all symbol definitions from a parsed VB6 module.
 *
 * Uses ProcessingState to track whether we're inside a procedure (for local
 * vs module scope) and to detect end-line numbers for multi-line blocks.
 */
export function extractSymbols(module: ParsedModule): SymbolDefinition[] {
  const { source, lines } = module;
  const { moduleName, path: filePath, type: fileType } = source;
  const definitions: SymbolDefinition[] = [];

  const state: ProcessingState = {
    currentProcedure: null,
    blockStack: [],
    insideType: false,
    insideEnum: false,
  };

  // Track procedure start indices so we can set endLineNumber
  // Key: index in definitions array, Value: block stack depth when opened
  const procStartIndices: Array<{ defIndex: number; blockDepth: number }> = [];

  for (const line of lines) {
    if (!line.isExecutable) continue;

    const trimmed = line.text.trim();
    if (!trimmed) continue;

    // ?? Check for block-end keywords first ??????????????????????????????

    if (RE_END_SUB.test(trimmed) || RE_END_FUNCTION.test(trimmed) || RE_END_PROPERTY.test(trimmed)) {
      // Close the current procedure
      if (state.currentProcedure !== null) {
        // Find the definition for this procedure and set endLineNumber
        const lastProc = procStartIndices.pop();
        if (lastProc !== undefined) {
          definitions[lastProc.defIndex].endLineNumber = line.lineNumber;
        }
        state.currentProcedure = null;
        // Pop the block from the stack
        if (state.blockStack.length > 0) {
          state.blockStack.pop();
        }
      }
      continue;
    }

    if (RE_END_TYPE.test(trimmed)) {
      state.insideType = false;
      // Set endLineNumber on the Type definition
      for (let i = definitions.length - 1; i >= 0; i--) {
        if (definitions[i].kind === 'Type' && definitions[i].endLineNumber === undefined) {
          definitions[i].endLineNumber = line.lineNumber;
          break;
        }
      }
      if (state.blockStack.length > 0) {
        state.blockStack.pop();
      }
      continue;
    }

    if (RE_END_ENUM.test(trimmed)) {
      state.insideEnum = false;
      // Set endLineNumber on the Enum definition
      for (let i = definitions.length - 1; i >= 0; i--) {
        if (definitions[i].kind === 'Enum' && definitions[i].endLineNumber === undefined) {
          definitions[i].endLineNumber = line.lineNumber;
          break;
        }
      }
      if (state.blockStack.length > 0) {
        state.blockStack.pop();
      }
      continue;
    }

    // Skip lines inside Type or Enum blocks (members are not separate symbols)
    if (state.insideType || state.insideEnum) continue;

    // ?? Try matching declaration patterns ????????????????????????????????

    // API Declare (must check before Sub/Function since Declare contains Sub/Function keyword)
    const declareMatch = trimmed.match(RE_DECLARE);
    if (declareMatch) {
      const { visibility } = parseVisibility(declareMatch[1]);
      const name = declareMatch[3];
      definitions.push({
        id: makeId(moduleName, name, line.lineNumber),
        name,
        kind: 'Declare',
        visibility,
        moduleName,
        filePath,
        lineNumber: line.lineNumber,
        scope: 'module',
        isEventHandler: false,
      });
      continue;
    }

    // Property Get/Let/Set
    const propMatch = trimmed.match(RE_PROPERTY);
    if (propMatch) {
      const { visibility } = parseVisibility(propMatch[1]);
      const name = propMatch[3];
      state.currentProcedure = name;
      state.blockStack.push('Property');

      const defIndex = definitions.length;
      procStartIndices.push({ defIndex, blockDepth: state.blockStack.length });

      definitions.push({
        id: makeId(moduleName, name, line.lineNumber),
        name,
        kind: 'Property',
        visibility,
        moduleName,
        filePath,
        lineNumber: line.lineNumber,
        scope: 'module',
        isEventHandler: false,
      });
      continue;
    }

    // Sub / Function
    const procMatch = trimmed.match(RE_PROC);
    if (procMatch) {
      const { visibility, explicit: visExplicit } = parseVisibility(procMatch[1]);
      const keyword = procMatch[3]; // Sub or Function
      const name = procMatch[4];
      const kind: SymbolKind = keyword.toLowerCase() === 'sub' ? 'Sub' : 'Function';

      // Default visibility: Public in .bas files if not explicitly specified
      const effectiveVisibility: Visibility = visExplicit
        ? visibility
        : (fileType === 'bas' ? 'Public' : 'Private');

      // Detect event handlers
      const eventHandler = kind === 'Sub' && isEventHandlerName(name);

      state.currentProcedure = name;
      state.blockStack.push(keyword);

      const defIndex = definitions.length;
      procStartIndices.push({ defIndex, blockDepth: state.blockStack.length });

      // Extract return type for Functions
      let dataType: string | undefined;
      if (kind === 'Function') {
        const retMatch = trimmed.match(/\)\s*As\s+(\w+)/i);
        if (retMatch) dataType = retMatch[1];
      }

      definitions.push({
        id: makeId(moduleName, name, line.lineNumber),
        name,
        kind,
        visibility: effectiveVisibility,
        moduleName,
        filePath,
        lineNumber: line.lineNumber,
        scope: 'module',
        dataType,
        isEventHandler: eventHandler,
      });
      continue;
    }

    // Const
    const constMatch = trimmed.match(RE_CONST);
    if (constMatch) {
      const scope: VariableScope = state.currentProcedure ? 'local' : 'module';
      const { visibility } = parseVisibility(constMatch[1]);
      // Local consts are effectively Private
      const effectiveVisibility: Visibility = scope === 'local' ? 'Private' : visibility;
      const name = constMatch[2];
      const dataType = constMatch[3];

      definitions.push({
        id: makeId(moduleName, name, line.lineNumber),
        name,
        kind: 'Const',
        visibility: effectiveVisibility,
        moduleName,
        filePath,
        lineNumber: line.lineNumber,
        scope,
        dataType,
        parentProcedure: state.currentProcedure ?? undefined,
        isEventHandler: false,
      });
      continue;
    }

    // Enum
    const enumMatch = trimmed.match(RE_ENUM);
    if (enumMatch) {
      const { visibility } = parseVisibility(enumMatch[1]);
      const name = enumMatch[2];
      state.insideEnum = true;
      state.blockStack.push('Enum');

      definitions.push({
        id: makeId(moduleName, name, line.lineNumber),
        name,
        kind: 'Enum',
        visibility,
        moduleName,
        filePath,
        lineNumber: line.lineNumber,
        scope: 'module',
        isEventHandler: false,
      });
      continue;
    }

    // Type (UDT)
    const typeMatch = trimmed.match(RE_TYPE);
    if (typeMatch) {
      const { visibility } = parseVisibility(typeMatch[1]);
      const name = typeMatch[2];
      state.insideType = true;
      state.blockStack.push('Type');

      definitions.push({
        id: makeId(moduleName, name, line.lineNumber),
        name,
        kind: 'Type',
        visibility,
        moduleName,
        filePath,
        lineNumber: line.lineNumber,
        scope: 'module',
        isEventHandler: false,
      });
      continue;
    }

    // Variable declarations (Public/Private/Dim/Global/Friend)
    const varMatch = trimmed.match(RE_VAR_LINE);
    if (varMatch) {
      const keyword = varMatch[1];
      const rest = varMatch[2];

      // Skip if this is actually a Const, Sub, Function, Declare, Enum, Type, Property, or Event
      // (those are handled above, but `Public Const ...` would match RE_VAR_LINE too)
      const restTrimmed = rest.trim();
      if (/^(Const|Sub|Function|Declare|Enum|Type|Property)\s/i.test(restTrimmed)) {
        continue;
      }
      // Also skip `Static` modifier followed by Sub/Function
      if (/^Static\s+(Sub|Function)\s/i.test(restTrimmed)) {
        continue;
      }

      const scope: VariableScope = state.currentProcedure ? 'local' : 'module';

      // Dim inside a procedure is always local
      // Public/Global cannot appear inside a procedure (VB6 rule), but we handle gracefully
      const baseVisibility = varKeywordVisibility(keyword);
      const effectiveVisibility: Visibility = scope === 'local' ? 'Private' : baseVisibility;

      const variables = parseVariableList(rest);
      for (const v of variables) {
        definitions.push({
          id: makeId(moduleName, v.name, line.lineNumber),
          name: v.name,
          kind: 'Variable',
          visibility: effectiveVisibility,
          moduleName,
          filePath,
          lineNumber: line.lineNumber,
          scope,
          dataType: v.dataType,
          parentProcedure: state.currentProcedure ?? undefined,
          isEventHandler: false,
        });
      }
      continue;
    }
  }

  return definitions;
}


// ?? Symbol Table Builder ?????????????????????????????????????????????????????

/**
 * Build a SymbolTable from an array of symbol definitions.
 *
 * Creates two maps:
 * - `symbols`: lowercase name ? definitions (for cross-module lookup)
 * - `byModule`: moduleName ? definitions (for per-module analysis)
 */
export function buildSymbolTable(definitions: SymbolDefinition[]): SymbolTable {
  const symbols = new Map<string, SymbolDefinition[]>();
  const byModule = new Map<string, SymbolDefinition[]>();

  for (const def of definitions) {
    // Index by lowercase name (VB6 is case-insensitive)
    const key = def.name.toLowerCase();
    const existing = symbols.get(key);
    if (existing) {
      existing.push(def);
    } else {
      symbols.set(key, [def]);
    }

    // Index by module name
    const moduleList = byModule.get(def.moduleName);
    if (moduleList) {
      moduleList.push(def);
    } else {
      byModule.set(def.moduleName, [def]);
    }
  }

  return { symbols, byModule };
}
