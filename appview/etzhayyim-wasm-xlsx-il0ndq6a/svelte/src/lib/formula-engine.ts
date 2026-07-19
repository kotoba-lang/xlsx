/**
 * Formula Engine — parse, evaluate, and track dependencies for spreadsheet formulas.
 *
 * Supports common Excel functions (SUM, AVERAGE, IF, VLOOKUP, etc.) with
 * dependency graph tracking for incremental recalculation. Circular reference
 * detection via topological sort.
 */

import type { XlsxSheet, XlsxCell, CellRef } from "./ooxml-parser";
import { parseRef, buildRef, letterToCol, colToLetter } from "./ooxml-parser";

// ---------------------------------------------------------------------------
// Token types
// ---------------------------------------------------------------------------

type TokenType = "number" | "string" | "cell" | "range" | "func" | "op" | "paren" | "comma" | "bool" | "error";

interface Token {
  type: TokenType;
  value: string;
}

// ---------------------------------------------------------------------------
// Tokenizer
// ---------------------------------------------------------------------------

/** Tokenize a formula string (without leading "="). */
function tokenize(formula: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const s = formula.trim();

  while (i < s.length) {
    const ch = s[i];

    // Whitespace
    if (ch === " " || ch === "\t") { i++; continue; }

    // String literal
    if (ch === '"') {
      let j = i + 1;
      while (j < s.length && s[j] !== '"') j++;
      tokens.push({ type: "string", value: s.slice(i + 1, j) });
      i = j + 1;
      continue;
    }

    // Number
    if (/\d/.test(ch) || (ch === "." && i + 1 < s.length && /\d/.test(s[i + 1]))) {
      let j = i;
      while (j < s.length && /[\d.]/.test(s[j])) j++;
      tokens.push({ type: "number", value: s.slice(i, j) });
      i = j;
      continue;
    }

    // Parentheses
    if (ch === "(" || ch === ")") {
      tokens.push({ type: "paren", value: ch });
      i++;
      continue;
    }

    // Comma
    if (ch === ",") {
      tokens.push({ type: "comma", value: "," });
      i++;
      continue;
    }

    // Operators
    if ("+-*/^%".includes(ch)) {
      tokens.push({ type: "op", value: ch });
      i++;
      continue;
    }
    if (ch === ">" || ch === "<" || ch === "=") {
      let op = ch;
      if (i + 1 < s.length && (s[i + 1] === "=" || (ch !== "=" && s[i + 1] === ">"))) {
        op += s[i + 1];
        i++;
      }
      tokens.push({ type: "op", value: op });
      i++;
      continue;
    }
    if (ch === "&") {
      tokens.push({ type: "op", value: "&" });
      i++;
      continue;
    }

    // Identifier (cell ref, range, function, boolean)
    if (/[A-Za-z$_]/.test(ch)) {
      let j = i;
      while (j < s.length && /[A-Za-z0-9$_:!.]/.test(s[j])) j++;
      const word = s.slice(i, j);
      const upper = word.toUpperCase();

      if (upper === "TRUE" || upper === "FALSE") {
        tokens.push({ type: "bool", value: upper });
      } else if (j < s.length && s[j] === "(") {
        tokens.push({ type: "func", value: upper });
      } else if (word.includes(":")) {
        tokens.push({ type: "range", value: upper });
      } else if (/^[A-Z]{1,3}\d+$/i.test(word.replace(/\$/g, ""))) {
        tokens.push({ type: "cell", value: upper.replace(/\$/g, "") });
      } else {
        tokens.push({ type: "func", value: upper });
      }
      i = j;
      continue;
    }

    // Skip unknown
    i++;
  }

  return tokens;
}

// ---------------------------------------------------------------------------
// Evaluator
// ---------------------------------------------------------------------------

type CellValue = string | number | boolean | null;

/** Evaluate a formula string in the context of a sheet. */
export function evaluateFormula(formula: string, sheet: XlsxSheet): CellValue {
  const tokens = tokenize(formula);
  const ctx: EvalContext = { sheet, pos: 0, tokens };

  try {
    const result = evalExpression(ctx);
    return result;
  } catch {
    return "#ERROR!";
  }
}

interface EvalContext {
  sheet: XlsxSheet;
  pos: number;
  tokens: Token[];
}

function peek(ctx: EvalContext): Token | null {
  return ctx.pos < ctx.tokens.length ? ctx.tokens[ctx.pos] : null;
}

function consume(ctx: EvalContext): Token {
  return ctx.tokens[ctx.pos++];
}

function evalExpression(ctx: EvalContext): CellValue {
  return evalComparison(ctx);
}

function evalComparison(ctx: EvalContext): CellValue {
  let left = evalConcat(ctx);
  while (peek(ctx)?.type === "op" && ["=", "<>", "!=", "<", ">", "<=", ">="].includes(peek(ctx)!.value)) {
    const op = consume(ctx).value;
    const right = evalConcat(ctx);
    const l = toNumber(left);
    const r = toNumber(right);
    switch (op) {
      case "=": left = left === right; break;
      case "<>": case "!=": left = left !== right; break;
      case "<": left = l < r; break;
      case ">": left = l > r; break;
      case "<=": left = l <= r; break;
      case ">=": left = l >= r; break;
    }
  }
  return left;
}

function evalConcat(ctx: EvalContext): CellValue {
  let left = evalAddSub(ctx);
  while (peek(ctx)?.value === "&") {
    consume(ctx);
    const right = evalAddSub(ctx);
    left = String(left ?? "") + String(right ?? "");
  }
  return left;
}

function evalAddSub(ctx: EvalContext): CellValue {
  let left = evalMulDiv(ctx);
  while (peek(ctx)?.type === "op" && (peek(ctx)!.value === "+" || peek(ctx)!.value === "-")) {
    const op = consume(ctx).value;
    const right = evalMulDiv(ctx);
    if (op === "+") left = toNumber(left) + toNumber(right);
    else left = toNumber(left) - toNumber(right);
  }
  return left;
}

function evalMulDiv(ctx: EvalContext): CellValue {
  let left = evalUnary(ctx);
  while (peek(ctx)?.type === "op" && "*/^%".includes(peek(ctx)!.value)) {
    const op = consume(ctx).value;
    const right = evalUnary(ctx);
    switch (op) {
      case "*": left = toNumber(left) * toNumber(right); break;
      case "/": { const d = toNumber(right); left = d === 0 ? "#DIV/0!" : toNumber(left) / d; break; }
      case "^": left = Math.pow(toNumber(left), toNumber(right)); break;
      case "%": left = toNumber(left) / 100; break;
    }
  }
  return left;
}

function evalUnary(ctx: EvalContext): CellValue {
  if (peek(ctx)?.value === "-") {
    consume(ctx);
    return -toNumber(evalPrimary(ctx));
  }
  if (peek(ctx)?.value === "+") {
    consume(ctx);
    return toNumber(evalPrimary(ctx));
  }
  return evalPrimary(ctx);
}

function evalPrimary(ctx: EvalContext): CellValue {
  const tok = peek(ctx);
  if (!tok) return null;

  if (tok.type === "number") {
    consume(ctx);
    return parseFloat(tok.value);
  }

  if (tok.type === "string") {
    consume(ctx);
    return tok.value;
  }

  if (tok.type === "bool") {
    consume(ctx);
    return tok.value === "TRUE";
  }

  if (tok.type === "cell") {
    consume(ctx);
    return getCellValue(ctx.sheet, tok.value as CellRef);
  }

  if (tok.type === "range") {
    consume(ctx);
    return null;
  }

  if (tok.type === "func") {
    return evalFunction(ctx);
  }

  if (tok.type === "paren" && tok.value === "(") {
    consume(ctx);
    const result = evalExpression(ctx);
    if (peek(ctx)?.value === ")") consume(ctx);
    return result;
  }

  consume(ctx);
  return null;
}

// ---------------------------------------------------------------------------
// Built-in functions
// ---------------------------------------------------------------------------

function evalFunction(ctx: EvalContext): CellValue {
  const name = consume(ctx).value;
  if (peek(ctx)?.value === "(") consume(ctx);

  const args: CellValue[] = [];
  const rangeArgs: string[] = [];

  while (peek(ctx) && peek(ctx)!.value !== ")") {
    if (peek(ctx)?.type === "range") {
      rangeArgs.push(consume(ctx).value);
      args.push(null);
    } else {
      args.push(evalExpression(ctx));
    }
    if (peek(ctx)?.value === ",") consume(ctx);
  }
  if (peek(ctx)?.value === ")") consume(ctx);

  /** Expand a range "A1:C3" into numeric cell values. */
  const expandRange = (rangeStr: string): number[] => {
    const [startRef, endRef] = rangeStr.split(":") as [CellRef, CellRef];
    const start = parseRef(startRef);
    const end = parseRef(endRef);
    const vals: number[] = [];
    for (let r = start.row; r <= end.row; r++) {
      for (let c = start.col; c <= end.col; c++) {
        const v = getCellValue(ctx.sheet, buildRef(c, r));
        if (typeof v === "number") vals.push(v);
      }
    }
    return vals;
  };

  /** Expand a range into all cell values (any type). */
  const expandRangeAll = (rangeStr: string): CellValue[] => {
    const [startRef, endRef] = rangeStr.split(":") as [CellRef, CellRef];
    const start = parseRef(startRef);
    const end = parseRef(endRef);
    const vals: CellValue[] = [];
    for (let r = start.row; r <= end.row; r++) {
      for (let c = start.col; c <= end.col; c++) {
        vals.push(getCellValue(ctx.sheet, buildRef(c, r)));
      }
    }
    return vals;
  };

  /** Expand a range as 2D array (for VLOOKUP/INDEX). */
  const expandRange2D = (rangeStr: string): CellValue[][] => {
    const [startRef, endRef] = rangeStr.split(":") as [CellRef, CellRef];
    const start = parseRef(startRef);
    const end = parseRef(endRef);
    const rows: CellValue[][] = [];
    for (let r = start.row; r <= end.row; r++) {
      const row: CellValue[] = [];
      for (let c = start.col; c <= end.col; c++) {
        row.push(getCellValue(ctx.sheet, buildRef(c, r)));
      }
      rows.push(row);
    }
    return rows;
  };

  const allNums = rangeArgs.length > 0 ? rangeArgs.flatMap(expandRange) : args.filter((a) => typeof a === "number") as number[];
  const allVals = rangeArgs.length > 0 ? rangeArgs.flatMap(expandRangeAll) : args;

  switch (name) {
    // ── Math & Trig (20) ──
    case "SUM": return allNums.reduce((a, b) => a + b, 0);
    case "SUMPRODUCT": {
      if (rangeArgs.length < 2) return 0;
      const arrays = rangeArgs.map(expandRange);
      const len = Math.min(...arrays.map(a => a.length));
      let total = 0;
      for (let i = 0; i < len; i++) { let prod = 1; for (const arr of arrays) prod *= arr[i]; total += prod; }
      return total;
    }
    case "AVERAGE": return allNums.length > 0 ? allNums.reduce((a, b) => a + b, 0) / allNums.length : 0;
    case "MIN": return allNums.length > 0 ? Math.min(...allNums) : 0;
    case "MAX": return allNums.length > 0 ? Math.max(...allNums) : 0;
    case "COUNT": return allNums.length;
    case "COUNTA": return allVals.filter(v => v != null && v !== "").length;
    case "COUNTBLANK": return allVals.filter(v => v == null || v === "").length;
    case "ROUND": return Math.round(toNumber(args[0]) * Math.pow(10, toNumber(args[1]))) / Math.pow(10, toNumber(args[1]));
    case "ROUNDUP": { const f = Math.pow(10, toNumber(args[1])); return Math.ceil(toNumber(args[0]) * f) / f; }
    case "ROUNDDOWN": { const f = Math.pow(10, toNumber(args[1])); return Math.floor(toNumber(args[0]) * f) / f; }
    case "CEILING": { const sig = toNumber(args[1]) || 1; return Math.ceil(toNumber(args[0]) / sig) * sig; }
    case "FLOOR": { const sig = toNumber(args[1]) || 1; return Math.floor(toNumber(args[0]) / sig) * sig; }
    case "INT": return Math.floor(toNumber(args[0]));
    case "SIGN": { const n = toNumber(args[0]); return n > 0 ? 1 : n < 0 ? -1 : 0; }
    case "ABS": return Math.abs(toNumber(args[0]));
    case "MOD": return toNumber(args[0]) % toNumber(args[1]);
    case "SQRT": return Math.sqrt(toNumber(args[0]));
    case "POWER": return Math.pow(toNumber(args[0]), toNumber(args[1]));
    case "LOG": return args.length >= 2 ? Math.log(toNumber(args[0])) / Math.log(toNumber(args[1])) : Math.log10(toNumber(args[0]));
    case "LOG10": return Math.log10(toNumber(args[0]));
    case "LN": return Math.log(toNumber(args[0]));
    case "EXP": return Math.exp(toNumber(args[0]));
    case "PI": return Math.PI;
    case "RAND": return Math.random();
    case "RANDBETWEEN": return Math.floor(Math.random() * (toNumber(args[1]) - toNumber(args[0]) + 1)) + toNumber(args[0]);
    case "PRODUCT": return allNums.length > 0 ? allNums.reduce((a, b) => a * b, 1) : 0;
    case "SUBTOTAL": {
      // Simplified: function_num 1-11 = AVERAGE..VAR, 101-111 = same ignoring hidden
      const fn = toNumber(args[0]);
      const fnMod = fn > 100 ? fn - 100 : fn;
      switch (fnMod) {
        case 1: return allNums.length > 0 ? allNums.reduce((a, b) => a + b, 0) / allNums.length : 0;
        case 2: return allNums.length;
        case 3: return allVals.filter(v => v != null && v !== "").length;
        case 4: return allNums.length > 0 ? Math.max(...allNums) : 0;
        case 5: return allNums.length > 0 ? Math.min(...allNums) : 0;
        case 6: return allNums.length > 0 ? allNums.reduce((a, b) => a * b, 1) : 0;
        case 9: return allNums.reduce((a, b) => a + b, 0);
        default: return allNums.reduce((a, b) => a + b, 0);
      }
    }

    // ── Logical (9) ──
    case "IF": return toBool(args[0]) ? args[1] : (args[2] ?? false);
    case "IFS": {
      for (let i = 0; i < args.length - 1; i += 2) { if (toBool(args[i])) return args[i + 1]; }
      return "#N/A";
    }
    case "SWITCH": {
      const expr = args[0];
      for (let i = 1; i < args.length - 1; i += 2) { if (expr === args[i]) return args[i + 1]; }
      return args.length % 2 === 0 ? args[args.length - 1] : "#N/A";
    }
    case "AND": return args.every(toBool);
    case "OR": return args.some(toBool);
    case "XOR": return args.filter(toBool).length % 2 === 1;
    case "NOT": return !toBool(args[0]);
    case "IFERROR": { const v = args[0]; return typeof v === "string" && v.startsWith("#") ? args[1] : v; }
    case "IFNA": { const v = args[0]; return v === "#N/A" ? args[1] : v; }
    case "TRUE": return true;
    case "FALSE": return false;

    // ── Text (22) ──
    case "CONCATENATE": case "CONCAT": return args.map(a => String(a ?? "")).join("");
    case "TEXTJOIN": {
      const delim = String(args[0] ?? "");
      const ignoreEmpty = toBool(args[1]);
      const vals = args.slice(2).map(a => String(a ?? ""));
      return (ignoreEmpty ? vals.filter(v => v !== "") : vals).join(delim);
    }
    case "LEFT": return String(args[0] ?? "").slice(0, toNumber(args[1]) || 1);
    case "RIGHT": { const s = String(args[0] ?? ""); const n = toNumber(args[1]) || 1; return s.slice(-n); }
    case "MID": return String(args[0] ?? "").slice(toNumber(args[1]) - 1, toNumber(args[1]) - 1 + toNumber(args[2]));
    case "LEN": return String(args[0] ?? "").length;
    case "FIND": { const idx = String(args[1] ?? "").indexOf(String(args[0] ?? ""), (toNumber(args[2]) || 1) - 1); return idx >= 0 ? idx + 1 : "#VALUE!"; }
    case "SEARCH": { const idx = String(args[1] ?? "").toLowerCase().indexOf(String(args[0] ?? "").toLowerCase(), (toNumber(args[2]) || 1) - 1); return idx >= 0 ? idx + 1 : "#VALUE!"; }
    case "SUBSTITUTE": {
      const s = String(args[0] ?? ""), old = String(args[1] ?? ""), rep = String(args[2] ?? "");
      const inst = args[3] != null ? toNumber(args[3]) : 0;
      if (inst <= 0) return s.replaceAll(old, rep);
      let count = 0, result = s;
      let pos = 0;
      while (true) { const idx = result.indexOf(old, pos); if (idx < 0) break; count++; if (count === inst) { result = result.slice(0, idx) + rep + result.slice(idx + old.length); break; } pos = idx + old.length; }
      return result;
    }
    case "REPLACE": { const s = String(args[0] ?? ""); const start = toNumber(args[1]) - 1; const n = toNumber(args[2]); return s.slice(0, start) + String(args[3] ?? "") + s.slice(start + n); }
    case "REPT": return String(args[0] ?? "").repeat(Math.max(0, toNumber(args[1])));
    case "EXACT": return String(args[0] ?? "") === String(args[1] ?? "");
    case "TRIM": return String(args[0] ?? "").trim().replace(/\s+/g, " ");
    case "CLEAN": return String(args[0] ?? "").replace(/[\x00-\x1F]/g, "");
    case "UPPER": return String(args[0] ?? "").toUpperCase();
    case "LOWER": return String(args[0] ?? "").toLowerCase();
    case "PROPER": return String(args[0] ?? "").replace(/\b\w/g, c => c.toUpperCase());
    case "TEXT": { const v = toNumber(args[0]); const fmt = String(args[1] ?? ""); if (fmt.includes("0")) return v.toFixed((fmt.split(".")[1] || "").length); return String(v); }
    case "VALUE": return parseFloat(String(args[0] ?? "")) || 0;
    case "CHAR": return String.fromCharCode(toNumber(args[0]));
    case "CODE": return (String(args[0] ?? "").charCodeAt(0)) || 0;
    case "T": return typeof args[0] === "string" ? args[0] : "";

    // ── Lookup & Reference (8) ──
    case "VLOOKUP": {
      const lookup = args[0];
      const table = rangeArgs.length > 0 ? expandRange2D(rangeArgs[0]) : [];
      const colIdx = toNumber(args[2]) - 1;
      const exact = args[3] === false || args[3] === 0;
      for (const row of table) { if (exact ? row[0] === lookup : toNumber(row[0]) <= toNumber(lookup)) { if (colIdx >= 0 && colIdx < row.length) return row[colIdx]; } }
      if (!exact && table.length > 0) { const last = table[table.length - 1]; if (colIdx >= 0 && colIdx < last.length) return last[colIdx]; }
      return "#N/A";
    }
    case "HLOOKUP": {
      const lookup = args[0];
      const table = rangeArgs.length > 0 ? expandRange2D(rangeArgs[0]) : [];
      const rowIdx = toNumber(args[2]) - 1;
      if (table.length === 0) return "#N/A";
      const cols = table[0]?.length ?? 0;
      for (let c = 0; c < cols; c++) { if (table[0][c] === lookup) { return rowIdx < table.length ? table[rowIdx][c] : "#N/A"; } }
      return "#N/A";
    }
    case "INDEX": {
      const table = rangeArgs.length > 0 ? expandRange2D(rangeArgs[0]) : [];
      const rowIdx = toNumber(args[1]) - 1;
      const colIdx = args.length >= 3 ? toNumber(args[2]) - 1 : 0;
      if (rowIdx >= 0 && rowIdx < table.length && colIdx >= 0 && colIdx < (table[rowIdx]?.length ?? 0)) return table[rowIdx][colIdx];
      return "#REF!";
    }
    case "MATCH": {
      const lookup = args[0];
      const vals = rangeArgs.length > 0 ? expandRangeAll(rangeArgs[0]) : [];
      const matchType = args.length >= 3 ? toNumber(args[2]) : 1;
      if (matchType === 0) { const idx = vals.indexOf(lookup); return idx >= 0 ? idx + 1 : "#N/A"; }
      if (matchType === 1) { let last = -1; for (let i = 0; i < vals.length; i++) { if (toNumber(vals[i]) <= toNumber(lookup)) last = i; else break; } return last >= 0 ? last + 1 : "#N/A"; }
      if (matchType === -1) { let last = -1; for (let i = 0; i < vals.length; i++) { if (toNumber(vals[i]) >= toNumber(lookup)) last = i; else break; } return last >= 0 ? last + 1 : "#N/A"; }
      return "#N/A";
    }
    case "XLOOKUP": {
      const lookup = args[0];
      const lookupArr = rangeArgs.length > 0 ? expandRangeAll(rangeArgs[0]) : [];
      const returnArr = rangeArgs.length > 1 ? expandRangeAll(rangeArgs[1]) : [];
      const notFound = args.length >= 4 ? args[3] : "#N/A";
      const idx = lookupArr.indexOf(lookup);
      return idx >= 0 && idx < returnArr.length ? returnArr[idx] : notFound;
    }
    case "CHOOSE": { const idx = toNumber(args[0]); return idx >= 1 && idx < args.length ? args[idx] : "#VALUE!"; }
    case "ROW": { if (args[0] && typeof args[0] === "string") { const p = parseRef(String(args[0])); return p.row + 1; } return 0; }
    case "COLUMN": { if (args[0] && typeof args[0] === "string") { const p = parseRef(String(args[0])); return p.col + 1; } return 0; }
    case "ROWS": return rangeArgs.length > 0 ? expandRange2D(rangeArgs[0]).length : 0;
    case "COLUMNS": return rangeArgs.length > 0 ? (expandRange2D(rangeArgs[0])[0]?.length ?? 0) : 0;
    case "INDIRECT": { const ref = String(args[0] ?? ""); return getCellValue(ctx.sheet, ref as CellRef); }

    // ── Date & Time (18) ──
    case "NOW": return new Date().toISOString();
    case "TODAY": return new Date().toISOString().slice(0, 10);
    case "DATE": { const d = new Date(toNumber(args[0]), toNumber(args[1]) - 1, toNumber(args[2])); return excelSerial(d); }
    case "TIME": return (toNumber(args[0]) * 3600 + toNumber(args[1]) * 60 + toNumber(args[2])) / 86400;
    case "YEAR": return dateFromArg(args[0]).getFullYear();
    case "MONTH": return dateFromArg(args[0]).getMonth() + 1;
    case "DAY": return dateFromArg(args[0]).getDate();
    case "HOUR": return dateFromArg(args[0]).getHours();
    case "MINUTE": return dateFromArg(args[0]).getMinutes();
    case "SECOND": return dateFromArg(args[0]).getSeconds();
    case "WEEKDAY": { const d = dateFromArg(args[0]); return d.getDay() + 1; }
    case "DATEVALUE": return excelSerial(new Date(String(args[0])));
    case "DATEDIF": {
      const d1 = dateFromArg(args[0]), d2 = dateFromArg(args[1]), unit = String(args[2] ?? "D").toUpperCase();
      if (unit === "D") return Math.floor((d2.getTime() - d1.getTime()) / 86400000);
      if (unit === "M") return (d2.getFullYear() - d1.getFullYear()) * 12 + d2.getMonth() - d1.getMonth();
      if (unit === "Y") return d2.getFullYear() - d1.getFullYear();
      return "#VALUE!";
    }
    case "EDATE": { const d = dateFromArg(args[0]); d.setMonth(d.getMonth() + toNumber(args[1])); return excelSerial(d); }
    case "EOMONTH": { const d = dateFromArg(args[0]); d.setMonth(d.getMonth() + toNumber(args[1]) + 1, 0); return excelSerial(d); }
    case "NETWORKDAYS": {
      const d1 = dateFromArg(args[0]), d2 = dateFromArg(args[1]);
      let count = 0, cur = new Date(d1);
      while (cur <= d2) { const dow = cur.getDay(); if (dow !== 0 && dow !== 6) count++; cur.setDate(cur.getDate() + 1); }
      return count;
    }
    case "WORKDAY": {
      const d = dateFromArg(args[0]); let days = toNumber(args[1]);
      while (days > 0) { d.setDate(d.getDate() + 1); const dow = d.getDay(); if (dow !== 0 && dow !== 6) days--; }
      return excelSerial(d);
    }
    case "ISOWEEKNUM": { const d = dateFromArg(args[0]); const jan4 = new Date(d.getFullYear(), 0, 4); const diff = d.getTime() - jan4.getTime(); return Math.ceil((diff / 86400000 + jan4.getDay() + 1) / 7); }

    // ── Statistical (14) ──
    case "COUNTIF": {
      if (rangeArgs.length === 0) return 0;
      const vals = expandRange(rangeArgs[0]);
      const criteria = toNumber(args[1]);
      return vals.filter(v => v === criteria).length;
    }
    case "COUNTIFS": {
      // Simplified: single range+criteria pair
      if (rangeArgs.length === 0) return 0;
      const vals = expandRange(rangeArgs[0]);
      const criteria = toNumber(args[1]);
      return vals.filter(v => v === criteria).length;
    }
    case "SUMIF": {
      if (rangeArgs.length === 0) return 0;
      const vals = expandRange(rangeArgs[0]);
      const criteria = toNumber(args[1]);
      return vals.filter(v => v === criteria).reduce((a, b) => a + b, 0);
    }
    case "SUMIFS": {
      if (rangeArgs.length === 0) return 0;
      const sumVals = expandRange(rangeArgs[0]);
      const criteriaVals = rangeArgs.length > 1 ? expandRange(rangeArgs[1]) : sumVals;
      const criteria = toNumber(args[2] ?? args[1]);
      let total = 0;
      for (let i = 0; i < sumVals.length; i++) { if (i < criteriaVals.length && criteriaVals[i] === criteria) total += sumVals[i]; }
      return total;
    }
    case "AVERAGEIF": {
      if (rangeArgs.length === 0) return 0;
      const vals = expandRange(rangeArgs[0]);
      const criteria = toNumber(args[1]);
      const matched = vals.filter(v => v === criteria);
      return matched.length > 0 ? matched.reduce((a, b) => a + b, 0) / matched.length : 0;
    }
    case "MEDIAN": {
      const sorted = [...allNums].sort((a, b) => a - b);
      if (sorted.length === 0) return 0;
      const mid = Math.floor(sorted.length / 2);
      return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
    }
    case "MODE": {
      const freq = new Map<number, number>();
      for (const n of allNums) freq.set(n, (freq.get(n) ?? 0) + 1);
      let maxF = 0, mode = 0;
      for (const [v, f] of freq) { if (f > maxF) { maxF = f; mode = v; } }
      return maxF > 1 ? mode : "#N/A";
    }
    case "STDEV": case "STDEV.S": {
      if (allNums.length < 2) return 0;
      const mean = allNums.reduce((a, b) => a + b, 0) / allNums.length;
      return Math.sqrt(allNums.reduce((s, v) => s + (v - mean) ** 2, 0) / (allNums.length - 1));
    }
    case "STDEVP": case "STDEV.P": {
      if (allNums.length === 0) return 0;
      const mean = allNums.reduce((a, b) => a + b, 0) / allNums.length;
      return Math.sqrt(allNums.reduce((s, v) => s + (v - mean) ** 2, 0) / allNums.length);
    }
    case "VAR": case "VAR.S": {
      if (allNums.length < 2) return 0;
      const mean = allNums.reduce((a, b) => a + b, 0) / allNums.length;
      return allNums.reduce((s, v) => s + (v - mean) ** 2, 0) / (allNums.length - 1);
    }
    case "VARP": case "VAR.P": {
      if (allNums.length === 0) return 0;
      const mean = allNums.reduce((a, b) => a + b, 0) / allNums.length;
      return allNums.reduce((s, v) => s + (v - mean) ** 2, 0) / allNums.length;
    }
    case "LARGE": { const sorted = [...allNums].sort((a, b) => b - a); const k = toNumber(args[1] ?? args[0]) - 1; return k >= 0 && k < sorted.length ? sorted[k] : "#NUM!"; }
    case "SMALL": { const sorted = [...allNums].sort((a, b) => a - b); const k = toNumber(args[1] ?? args[0]) - 1; return k >= 0 && k < sorted.length ? sorted[k] : "#NUM!"; }
    case "RANK": case "RANK.EQ": {
      const val = toNumber(args[0]);
      const sorted = [...allNums].sort((a, b) => b - a);
      const idx = sorted.indexOf(val);
      return idx >= 0 ? idx + 1 : "#N/A";
    }
    case "PERCENTILE": case "PERCENTILE.INC": {
      if (allNums.length === 0) return 0;
      const k = toNumber(args[1] ?? args[0]);
      const sorted = [...allNums].sort((a, b) => a - b);
      const idx = k * (sorted.length - 1);
      const lo = Math.floor(idx), hi = Math.ceil(idx);
      return lo === hi ? sorted[lo] : sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
    }
    case "MAXIFS": {
      if (rangeArgs.length < 2) return allNums.length > 0 ? Math.max(...allNums) : 0;
      const maxVals = expandRange(rangeArgs[0]);
      const criteriaVals = expandRange(rangeArgs[1]);
      const criteria = toNumber(args[2]);
      const matched = maxVals.filter((_, i) => i < criteriaVals.length && criteriaVals[i] === criteria);
      return matched.length > 0 ? Math.max(...matched) : 0;
    }
    case "MINIFS": {
      if (rangeArgs.length < 2) return allNums.length > 0 ? Math.min(...allNums) : 0;
      const minVals = expandRange(rangeArgs[0]);
      const criteriaVals = expandRange(rangeArgs[1]);
      const criteria = toNumber(args[2]);
      const matched = minVals.filter((_, i) => i < criteriaVals.length && criteriaVals[i] === criteria);
      return matched.length > 0 ? Math.min(...matched) : 0;
    }

    // ── Information (10) ──
    case "ISNUMBER": return typeof args[0] === "number";
    case "ISTEXT": return typeof args[0] === "string" && !String(args[0]).startsWith("#");
    case "ISBLANK": return args[0] == null || args[0] === "";
    case "ISERROR": return typeof args[0] === "string" && String(args[0]).startsWith("#");
    case "ISNA": return args[0] === "#N/A";
    case "ISLOGICAL": return typeof args[0] === "boolean";
    case "ISFORMULA": return false; // simplified — would need cell ref context
    case "TYPE": { const v = args[0]; if (typeof v === "number") return 1; if (typeof v === "string") return 2; if (typeof v === "boolean") return 4; if (v == null) return 1; return 16; }
    case "N": { const v = args[0]; if (typeof v === "number") return v; if (typeof v === "boolean") return v ? 1 : 0; return 0; }
    case "NA": return "#N/A";
    case "ERROR.TYPE": {
      const v = String(args[0] ?? "");
      const map: Record<string, number> = { "#NULL!": 1, "#DIV/0!": 2, "#VALUE!": 3, "#REF!": 4, "#NAME?": 5, "#NUM!": 6, "#N/A": 7 };
      return map[v] ?? "#N/A";
    }

    // ── Financial (8) ──
    case "PMT": {
      const rate = toNumber(args[0]), nper = toNumber(args[1]), pv = toNumber(args[2]), fv = toNumber(args[3] ?? 0);
      if (rate === 0) return -(pv + fv) / nper;
      return -(pv * rate * Math.pow(1 + rate, nper) + fv * rate) / (Math.pow(1 + rate, nper) - 1);
    }
    case "FV": {
      const rate = toNumber(args[0]), nper = toNumber(args[1]), pmt = toNumber(args[2]), pv = toNumber(args[3] ?? 0);
      if (rate === 0) return -(pv + pmt * nper);
      return -(pv * Math.pow(1 + rate, nper) + pmt * (Math.pow(1 + rate, nper) - 1) / rate);
    }
    case "PV": {
      const rate = toNumber(args[0]), nper = toNumber(args[1]), pmt = toNumber(args[2]), fv = toNumber(args[3] ?? 0);
      if (rate === 0) return -(fv + pmt * nper);
      return -(fv / Math.pow(1 + rate, nper) + pmt * (1 - Math.pow(1 + rate, -nper)) / rate);
    }
    case "NPV": {
      const rate = toNumber(args[0]);
      let npv = 0;
      const cashflows = allNums.length > 0 ? allNums : args.slice(1).map(toNumber);
      for (let i = 0; i < cashflows.length; i++) npv += cashflows[i] / Math.pow(1 + rate, i + 1);
      return npv;
    }
    case "IRR": {
      const cashflows = allNums.length > 0 ? allNums : args.slice(0).map(toNumber);
      let guess = toNumber(args[1] ?? 0.1);
      for (let iter = 0; iter < 100; iter++) {
        let npv = 0, dnpv = 0;
        for (let i = 0; i < cashflows.length; i++) { npv += cashflows[i] / Math.pow(1 + guess, i); dnpv -= i * cashflows[i] / Math.pow(1 + guess, i + 1); }
        if (Math.abs(npv) < 1e-10) return guess;
        guess -= npv / dnpv;
      }
      return guess;
    }
    case "RATE": {
      const nper = toNumber(args[0]), pmt = toNumber(args[1]), pv = toNumber(args[2]), fv = toNumber(args[3] ?? 0);
      let rate = toNumber(args[4] ?? 0.1);
      for (let i = 0; i < 100; i++) {
        const t = Math.pow(1 + rate, nper);
        const f = pv * t + pmt * (t - 1) / rate + fv;
        const df = pv * nper * Math.pow(1 + rate, nper - 1) + pmt * ((nper * Math.pow(1 + rate, nper - 1) * rate - (t - 1)) / (rate * rate));
        rate -= f / df;
        if (Math.abs(f) < 1e-10) break;
      }
      return rate;
    }
    case "NPER": {
      const rate = toNumber(args[0]), pmt = toNumber(args[1]), pv = toNumber(args[2]), fv = toNumber(args[3] ?? 0);
      if (rate === 0) return -(pv + fv) / pmt;
      return Math.log((-fv * rate + pmt) / (pv * rate + pmt)) / Math.log(1 + rate);
    }
    case "SLN": { const cost = toNumber(args[0]), salvage = toNumber(args[1]), life = toNumber(args[2]); return (cost - salvage) / life; }

    default: return `#NAME?`;
  }
}

// ---------------------------------------------------------------------------
// Dependency graph
// ---------------------------------------------------------------------------

/** Extract all cell references from a formula string. */
export function getFormulaDependencies(formula: string): CellRef[] {
  const deps: CellRef[] = [];
  const tokens = tokenize(formula);
  for (const tok of tokens) {
    if (tok.type === "cell") {
      deps.push(tok.value as CellRef);
    } else if (tok.type === "range") {
      const [startRef, endRef] = tok.value.split(":") as [CellRef, CellRef];
      const start = parseRef(startRef);
      const end = parseRef(endRef);
      for (let r = start.row; r <= end.row; r++) {
        for (let c = start.col; c <= end.col; c++) {
          deps.push(buildRef(c, r));
        }
      }
    }
  }
  return deps;
}

/** Detect circular references starting from a cell. */
export function detectCircular(sheet: XlsxSheet, startRef: CellRef): boolean {
  const visited = new Set<CellRef>();
  const stack = [startRef];

  while (stack.length > 0) {
    const ref = stack.pop()!;
    if (visited.has(ref)) continue;
    visited.add(ref);

    const cell = sheet.cells.get(ref);
    if (!cell?.formula) continue;

    const deps = getFormulaDependencies(cell.formula);
    for (const dep of deps) {
      if (dep === startRef) return true;
      stack.push(dep);
    }
  }

  return false;
}

/** Recalculate all formula cells in a sheet (topological order). */
export function recalculateSheet(sheet: XlsxSheet): void {
  // Build dependency graph: ref → set of formula cells that depend on it
  const formulaCells: CellRef[] = [];
  const depMap = new Map<CellRef, CellRef[]>();
  // dependents: for each cell, which formula cells depend on it
  const dependents = new Map<CellRef, CellRef[]>();

  for (const [ref, cell] of sheet.cells) {
    if (cell.formula) {
      formulaCells.push(ref);
      const deps = getFormulaDependencies(cell.formula);
      depMap.set(ref, deps);
      for (const dep of deps) {
        if (!dependents.has(dep)) dependents.set(dep, []);
        dependents.get(dep)!.push(ref);
      }
    }
  }

  // Topological sort (Kahn's algorithm)
  // in-degree = number of dependencies that are also formula cells
  const inDegree = new Map<CellRef, number>();
  for (const ref of formulaCells) {
    const deps = depMap.get(ref) ?? [];
    let deg = 0;
    for (const dep of deps) {
      if (depMap.has(dep)) deg++; // dep is also a formula cell
    }
    inDegree.set(ref, deg);
  }

  const queue: CellRef[] = [];
  for (const [ref, deg] of inDegree) {
    if (deg === 0) queue.push(ref);
  }

  const order: CellRef[] = [];
  while (queue.length > 0) {
    const ref = queue.shift()!;
    order.push(ref);
    // For each formula cell that depends on ref, decrement its in-degree
    const deptList = dependents.get(ref) ?? [];
    for (const dependent of deptList) {
      if (inDegree.has(dependent)) {
        const newDeg = (inDegree.get(dependent) ?? 1) - 1;
        inDegree.set(dependent, newDeg);
        if (newDeg === 0) queue.push(dependent);
      }
    }
  }

  // Evaluate in topological order (dependencies first)
  for (const ref of order) {
    const cell = sheet.cells.get(ref);
    if (!cell?.formula) continue;

    if (detectCircular(sheet, ref)) {
      cell.calculatedValue = "#REF!";
    } else {
      cell.calculatedValue = evaluateFormula(cell.formula, sheet);
    }
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Get the resolved value of a cell. */
function getCellValue(sheet: XlsxSheet, ref: CellRef): CellValue {
  const cell = sheet.cells.get(ref);
  if (!cell) return null;
  if (cell.type === "formula") return cell.calculatedValue;
  return cell.value;
}

function toNumber(v: CellValue): number {
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (typeof v === "string") return parseFloat(v) || 0;
  return 0;
}

function toBool(v: CellValue): boolean {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (typeof v === "string") return v.length > 0;
  return false;
}

/** Convert a CellValue to a Date. Handles Excel serial numbers and ISO strings. */
function dateFromArg(v: CellValue): Date {
  if (typeof v === "number") return excelSerialToDate(v);
  if (typeof v === "string") return new Date(v);
  return new Date();
}

/** Excel serial number epoch: 1900-01-01 = 1 (with Lotus 1-2-3 leap year bug). */
const EXCEL_EPOCH = new Date(1899, 11, 30).getTime();

/** Convert Date to Excel serial number. */
function excelSerial(d: Date): number {
  return Math.floor((d.getTime() - EXCEL_EPOCH) / 86400000);
}

/** Convert Excel serial number to Date. */
function excelSerialToDate(serial: number): Date {
  return new Date(EXCEL_EPOCH + serial * 86400000);
}
