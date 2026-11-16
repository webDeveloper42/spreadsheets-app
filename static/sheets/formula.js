"use strict";
// The spreadsheet engine: formulas, functions, number formats and typed input, with no page code in it, so the same
// file runs in the browser and in Node for the tests. Written from how Excel and Google Sheets behave; no code from either.
//
// A workbook is plain JSON: { sheets: [{ name, rows, cols, cells: { A1: { v, s } } }] }. v is what was typed, already
// parsed: a number, true/false, text, or text starting with "=" for a formula. Text starting with ' is shown without it.
(function (root) {

// ---------------------------------------------------------------- values
class FErr {
  constructor(code, msg) { this.code = code; this.msg = msg || ""; }
  toString() { return this.code; }
}
const ERRS = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!", "#ERROR!"];
const isErr = (v) => v instanceof FErr;
const E = {
  na: (m) => new FErr("#N/A", m || "Value not found"), div: (m) => new FErr("#DIV/0!", m || "Divided by zero"),
  value: (m) => new FErr("#VALUE!", m || "Wrong type of value"), ref: (m) => new FErr("#REF!", m || "Reference isn't valid"),
  name: (m) => new FErr("#NAME?", m || "Unknown name"), num: (m) => new FErr("#NUM!", m || "Number out of range"),
  spill: (m) => new FErr("#SPILL!", m || "Something is in the way of the results"), calc: (m) => new FErr("#CALC!", m || "Can't calculate"),
  parse: (m) => new FErr("#ERROR!", m || "Formula parse error"), nul: (m) => new FErr("#NULL!", m),
};
// A 2-D block of values, from a range, an array literal or a function like SEQUENCE.
class Arr {
  constructor(rows) { this.rows = rows; }
  get h() { return this.rows.length; }
  get w() { return this.rows.length ? this.rows[0].length : 0; }
  at(r, c) { const row = this.rows[r]; return row ? (row[c] === undefined ? null : row[c]) : null; }
}
// A reference to cells; functions like ROW, OFFSET and INDEX need the place, not just the values.
class Ref {
  constructor(sheet, r1, c1, r2, c2) { this.sheet = sheet; this.r1 = Math.min(r1, r2); this.c1 = Math.min(c1, c2); this.r2 = Math.max(r1, r2); this.c2 = Math.max(c1, c2); }
  get h() { return this.r2 - this.r1 + 1; }
  get w() { return this.c2 - this.c1 + 1; }
  get single() { return this.r1 === this.r2 && this.c1 === this.c2; }
}

// ---------------------------------------------------------------- addresses
function colName(c) { let s = ""; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
function colIndex(name) { let n = 0; for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
const addr = (r, c) => colName(c) + (r + 1);
function parseAddr(a) { const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a); return m ? { r: Number(m[2]) - 1, c: colIndex(m[1]) } : null; }
const MAX_ROWS = 1048576, MAX_COLS = 16384;
function quoteSheet(name) { return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`; }

// ---------------------------------------------------------------- tokens
// Every token keeps its exact text, so formulas can be rewritten (copied, rows inserted) without changing anything else.
const SHEET = String.raw`(?:'((?:[^']|'')+)'|([A-Za-z_À-￿][\w.À-￿]*))!`;
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d+)`;
const RE_REF = new RegExp(String.raw`^(?:${SHEET})?(?:${CELL}(?::${CELL})?|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})|(\$?)(\d+):(\$?)(\d+))(?![\w(])`);
const RE_NUM = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/;
const RE_NAME = /^[A-Za-z_\\][\w.]*/;
function tokenize(src) {
  const out = [];
  let i = 0;
  const push = (t, text, extra) => { out.push({ t, text, ...extra }); i += text.length; };
  while (i < src.length) {
    const rest = src.slice(i), ch = src[i];
    if (/\s/.test(ch)) { push("ws", /^\s+/.exec(rest)[0]); continue; }
    if (ch === '"') {
      let j = i + 1, val = "";
      for (;;) {
        if (j >= src.length) throw E.parse("A text in quotes isn't closed");
        if (src[j] === '"') { if (src[j + 1] === '"') { val += '"'; j += 2; continue; } break; }
        val += src[j++];
      }
      push("str", src.slice(i, j + 1), { v: val }); continue;
    }
    if (ch === "#") {
      const code = ERRS.find((e) => rest.toUpperCase().startsWith(e));
      if (code) { push("err", rest.slice(0, code.length), { v: code }); continue; }
    }
    const ref = RE_REF.exec(rest);
    if (ref) { push("ref", ref[0], { m: ref }); continue; }
    const num = RE_NUM.exec(rest);
    if (num) { push("num", num[0], { v: Number(num[0]) }); continue; }
    const name = RE_NAME.exec(rest);
    if (name) {
      const after = src.slice(i + name[0].length).match(/^\s*\(/);
      const up = name[0].toUpperCase();
      if (after) push("func", name[0], { v: up.replace(/^_XLFN\.|^_XLWS\./, "") });
      else if (up === "TRUE" || up === "FALSE") push("bool", name[0], { v: up === "TRUE" });
      else push("name", name[0], { v: up });
      continue;
    }
    const two = rest.slice(0, 2);
    if (["<=", ">=", "<>"].includes(two)) { push("op", two); continue; }
    if ("+-*/^&=<>%:".includes(ch)) { push("op", ch); continue; }
    if (ch === "(") { push("(", ch); continue; }
    if (ch === ")") { push(")", ch); continue; }
    if (ch === "{") { push("{", ch); continue; }
    if (ch === "}") { push("}", ch); continue; }
    if (ch === ",") { push(",", ch); continue; }
    if (ch === ";") { push(";", ch); continue; }
    if (ch === "@") { push("ws", ch); continue; } // implicit intersection marker from newer Excel files
    throw E.parse(`Unexpected “${ch}”`);
  }
  return out;
}

// A ref token as sheet name plus coordinates. Whole columns and rows use open ends.
function refParts(tok) {
  const m = tok.m;
  const sheet = m[1] != null ? m[1].replace(/''/g, "'") : m[2] != null ? m[2] : null;
  if (m[4] != null) {
    const a = { r: Number(m[6]) - 1, c: colIndex(m[4]), ar: !!m[5], ac: !!m[3] };
    const b = m[8] != null ? { r: Number(m[10]) - 1, c: colIndex(m[8]), ar: !!m[9], ac: !!m[7] } : null;
    return { kind: b ? "range" : "cell", sheet, a, b };
  }
  if (m[12] != null) return { kind: "cols", sheet, a: { c: colIndex(m[12]), ac: !!m[11] }, b: { c: colIndex(m[14]), ac: !!m[13] } };
  return { kind: "rows", sheet, a: { r: Number(m[16]) - 1, ar: !!m[15] }, b: { r: Number(m[18]) - 1, ar: !!m[17] } };
}
function refText(p) {
  const sh = p.sheet != null ? quoteSheet(p.sheet) + "!" : "";
  const cell = (x) => `${x.ac ? "$" : ""}${colName(x.c)}${x.ar ? "$" : ""}${x.r + 1}`;
  if (p.kind === "cell") return sh + cell(p.a);
  if (p.kind === "range") return sh + cell(p.a) + ":" + cell(p.b);
  if (p.kind === "cols") return `${sh}${p.a.ac ? "$" : ""}${colName(p.a.c)}:${p.b.ac ? "$" : ""}${colName(p.b.c)}`;
  return `${sh}${p.a.ar ? "$" : ""}${p.a.r + 1}:${p.b.ar ? "$" : ""}${p.b.r + 1}`;
}

// ---------------------------------------------------------------- parser
// Precedence, lowest first: comparison, &, + -, * /, ^, unary minus, %, then values. Like both apps, -2^2 is 4.
function parse(src) {
  const tokens = tokenize(src).filter((t) => t.t !== "ws");
  let p = 0;
  const peek = () => tokens[p], next = () => tokens[p++];
  const isOp = (t, ...ops) => t && t.t === "op" && ops.includes(t.text);
  const expect = (type) => { const t = next(); if (!t || t.t !== type) throw E.parse(type === ")" ? "A bracket isn't closed" : `Expected ${type}`); return t; };
  function expr() { return comparison(); }
  function comparison() { let l = concat(); while (isOp(peek(), "=", "<>", "<", ">", "<=", ">=")) { const op = next().text; l = { k: "bin", op, l, r: concat() }; } return l; }
  function concat() { let l = additive(); while (isOp(peek(), "&")) { next(); l = { k: "bin", op: "&", l, r: additive() }; } return l; }
  function additive() { let l = mult(); while (isOp(peek(), "+", "-")) { const op = next().text; l = { k: "bin", op, l, r: mult() }; } return l; }
  function mult() { let l = power(); while (isOp(peek(), "*", "/")) { const op = next().text; l = { k: "bin", op, l, r: power() }; } return l; }
  function power() { let l = unary(); while (isOp(peek(), "^")) { next(); l = { k: "bin", op: "^", l, r: unary() }; } return l; }
  function unary() {
    if (isOp(peek(), "-")) { next(); return { k: "neg", x: unary() }; }
    if (isOp(peek(), "+")) { next(); return unary(); }
    return percent();
  }
  function percent() { let x = range(); while (isOp(peek(), "%")) { next(); x = { k: "pct", x }; } return x; }
  function range() {
    let x = primary();
    while (isOp(peek(), ":")) { next(); const y = primary(); x = { k: "join", l: x, r: y }; } // A1:INDEX(...) and the like
    return x;
  }
  function primary() {
    const t = next();
    if (!t) throw E.parse("The formula ends too early");
    switch (t.t) {
      case "num": return { k: "lit", v: t.v };
      case "str": return { k: "lit", v: t.v };
      case "bool": return { k: "lit", v: t.v };
      case "err": return { k: "lit", v: new FErr(t.v) };
      case "ref": return { k: "ref", p: refParts(t) };
      case "name": return { k: "name", v: t.v };
      case "(": { const x = expr(); expect(")"); return x; }
      case "{": {
        const rows = [[]];
        for (;;) {
          let neg = false;
          if (isOp(peek(), "-")) { next(); neg = true; }
          const v = next();
          if (!v || !["num", "str", "bool", "err"].includes(v.t)) throw E.parse("An array can only hold plain values");
          rows[rows.length - 1].push(v.t === "err" ? new FErr(v.v) : neg ? -v.v : v.v);
          const sep = next();
          if (!sep) throw E.parse("An array isn't closed");
          if (sep.t === "}") break;
          if (sep.t === ";") rows.push([]);
          else if (sep.t !== ",") throw E.parse("Separate array values with , and rows with ;");
        }
        if (rows.some((r) => r.length !== rows[0].length)) throw E.parse("Every row of an array needs the same length");
        return { k: "lit", v: new Arr(rows) };
      }
      case "func": {
        expect("(");
        const args = [];
        if (peek() && peek().t === ")") { next(); return { k: "call", name: t.v, args }; }
        for (;;) {
          if (peek() && (peek().t === "," || peek().t === ";")) args.push({ k: "empty" });
          else if (peek() && peek().t === ")") { args.push({ k: "empty" }); }
          else args.push(expr());
          const sep = next();
          if (!sep) throw E.parse("A bracket isn't closed");
          if (sep.t === ")") break;
          if (sep.t !== "," && sep.t !== ";") throw E.parse(`Unexpected “${sep.text}”`);
        }
        return { k: "call", name: t.v, args };
      }
      case "op": throw E.parse(`Unexpected “${t.text}”`);
      default: throw E.parse(`Unexpected “${t.text}”`);
    }
  }
  const body = expr();
  if (p < tokens.length) throw E.parse(`Unexpected “${tokens[p].text}”`);
  return body;
}

// ---------------------------------------------------------------- rewriting formulas
// Copying shifts relative references; inserting and deleting rows or columns moves every reference that points past them.
function rewrite(formula, fn) {
  let tokens;
  try { tokens = tokenize(formula.slice(1)); } catch (e) { return formula; }
  return "=" + tokens.map((t) => {
    if (t.t !== "ref") return t.text;
    const p = refParts(t);
    const out = fn(p);
    return out === null ? "#REF!" : typeof out === "string" ? out : refText(out);
  }).join("");
}
function shiftFormula(formula, dr, dc) {
  return rewrite(formula, (p) => {
    const move = (x) => {
      const y = { ...x };
      if (y.r != null && !y.ar) y.r += dr;
      if (y.c != null && !y.ac) y.c += dc;
      return y;
    };
    const q = { ...p, a: move(p.a), b: p.b ? move(p.b) : null };
    for (const x of [q.a, q.b]) if (x && ((x.r != null && (x.r < 0 || x.r >= MAX_ROWS)) || (x.c != null && (x.c < 0 || x.c >= MAX_COLS)))) return null;
    return q;
  });
}
// axis "r" or "c"; at: first index; n > 0 inserts n, n < 0 deletes -n. Only references into sheetName (or unqualified ones when
// the formula lives on that sheet) move.
function structFormula(formula, home, sheetName, axis, at, n) {
  const low = (s) => (s || "").toLowerCase();
  return rewrite(formula, (p) => {
    const target = p.sheet != null ? p.sheet : home;
    if (low(target) !== low(sheetName)) return p;
    if (axis === "r" && p.kind === "cols") return p;
    if (axis === "c" && p.kind === "rows") return p;
    const key = axis;
    const q = { ...p, a: { ...p.a }, b: p.b ? { ...p.b } : null };
    if (n > 0) {
      for (const x of [q.a, q.b]) if (x && x[key] >= at) x[key] += n;
      return q;
    }
    const del = -n, end = at + del - 1;
    if (!q.b) { if (q.a[key] >= at && q.a[key] <= end) return null; if (q.a[key] > end) q.a[key] -= del; return q; }
    let a = q.a[key], b = q.b[key];
    if (a >= at && b <= end) return null; // the whole range is gone
    const fix = (v, isEnd) => v < at ? v : v > end ? v - del : isEnd ? at - 1 : at;
    a = fix(a, false); b = fix(b, true);
    if (b < a) b = a;
    q.a[key] = a; q.b[key] = b;
    return q;
  });
}
function renameSheetIn(formula, from, to) {
  return rewrite(formula, (p) => (p.sheet != null && p.sheet.toLowerCase() === from.toLowerCase() ? { ...p, sheet: to } : p));
}
function dropSheetIn(formula, name) {
  return rewrite(formula, (p) => (p.sheet != null && p.sheet.toLowerCase() === name.toLowerCase() ? null : p));
}
// The references in a formula, for coloured boxes while you type it.
function refsIn(formula) {
  let tokens;
  try { tokens = tokenize(formula.replace(/^=/, "")); } catch (e) { return []; }
  const out = [];
  let pos = formula.startsWith("=") ? 1 : 0;
  for (const t of tokens) {
    if (t.t === "ref") {
      const p = refParts(t);
      const box = p.kind === "cell" ? { r1: p.a.r, c1: p.a.c, r2: p.a.r, c2: p.a.c } : p.kind === "range" ? { r1: p.a.r, c1: p.a.c, r2: p.b.r, c2: p.b.c }
        : p.kind === "cols" ? { r1: 0, c1: p.a.c, r2: MAX_ROWS - 1, c2: p.b.c } : { r1: p.a.r, c1: 0, r2: p.b.r, c2: MAX_COLS - 1 };
      out.push({ sheet: p.sheet, start: pos, end: pos + t.text.length, text: t.text, ...box });
    }
    pos += t.text.length;
  }
  return out;
}

// ---------------------------------------------------------------- coercion
const isBlank = (v) => v === null || v === undefined;
function numText(n) {
  if (!isFinite(n)) return "#NUM!";
  if (Number.isInteger(n) && Math.abs(n) < 1e21) return String(n);
  let s = String(Number(n.toPrecision(15)));
  if (s.includes("e")) { const [m, x] = s.split("e"); s = `${m}E${x[0] === "-" ? "-" : "+"}${x.replace(/^[+-]/, "").padStart(2, "0")}`; }
  return s;
}
function toNum(v) {
  if (typeof v === "number") return v;
  if (isBlank(v)) return 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (isErr(v)) throw v;
  if (typeof v === "string") {
    const t = v.trim();
    if (t === "") throw E.value("Expected a number but got empty text");
    const parsed = parseInput(t);
    if (typeof parsed.v === "number") return parsed.v;
    throw E.value(`“${v.length > 30 ? v.slice(0, 30) + "…" : v}” isn't a number`);
  }
  throw E.value();
}
function toStr(v) {
  if (typeof v === "string") return v;
  if (isBlank(v)) return "";
  if (typeof v === "number") return numText(v);
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (isErr(v)) throw v;
  return String(v);
}
function toBool(v) {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (isBlank(v)) return false;
  if (isErr(v)) throw v;
  if (typeof v === "string") { const u = v.trim().toUpperCase(); if (u === "TRUE") return true; if (u === "FALSE") return false; throw E.value(`“${v}” isn't TRUE or FALSE`); }
  throw E.value();
}
const toInt = (v) => Math.trunc(toNum(v));
// Ordering like both apps: numbers, then text (ignoring case), then TRUE/FALSE. A blank takes the other side's type.
function typeRank(v) { return typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3; }
function compare(a, b) {
  if (isBlank(a)) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (isBlank(b)) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const ta = typeRank(a), tb = typeRank(b);
  if (ta !== tb) return ta < tb ? -1 : 1;
  if (typeof a === "string") { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  if (typeof a === "boolean") return a === b ? 0 : a ? 1 : -1;
  if (a === b) return 0;
  const x = Number(a.toPrecision(15)), y = Number(b.toPrecision(15)); // like both apps, 0.1+0.2 equals 0.3
  return x < y ? -1 : x > y ? 1 : 0;
}

// ---------------------------------------------------------------- dates
// Day numbers like both apps: 1 = 1900-01-01 (with Excel's phantom 29 Feb 1900 kept, so later dates match Excel files).
const DAY = 86400000, EPOCH = Date.UTC(1899, 11, 30);
function dateSerial(y, m, d) { return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY); }
function fromSerial(n) {
  const whole = Math.floor(n + 1e-9);
  const dt = new Date(EPOCH + whole * DAY);
  const secs = Math.round((n - whole) * 86400);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), wd: dt.getUTCDay(),
    hh: Math.floor(secs / 3600) % 24, mi: Math.floor(secs / 60) % 60, ss: secs % 60, frac: n - whole };
}
function todaySerial(now = new Date()) { return dateSerial(now.getFullYear(), now.getMonth() + 1, now.getDate()); }
function nowSerial(now = new Date()) { return todaySerial(now) + (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) / 86400; }
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

// ---------------------------------------------------------------- typing into a cell
// What you type becomes a value plus, sometimes, a number format: "$1,200" is 1200 shown as money, "5%" is 0.05.
function parseInput(text) {
  if (text === null || text === undefined) return { v: null };
  if (typeof text !== "string") return { v: text };
  const t = text.trim();
  if (text === "") return { v: null };
  if (text.startsWith("=") && text.length > 1) return { v: text };
  if (text.startsWith("'")) return { v: text };
  if (/^(true|false)$/i.test(t)) return { v: t.toUpperCase() === "TRUE" };
  const num = /^([+-])?(\$)?([+-])?((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE]([+-]?\d+))?(%)?$/.exec(t);
  if (num) {
    let v = Number(num[4].replace(/,/g, "")) * (num[5] ? Math.pow(10, Number(num[5])) : 1);
    if (num[1] === "-" || num[3] === "-") v = -v;
    const dec = (num[4].split(".")[1] || "").length;
    if (num[6]) return { v: v / 100, fmt: dec ? `0.${"0".repeat(Math.min(dec, 10))}%` : "0%" };
    if (num[2]) return { v, fmt: dec ? "$#,##0.00" : "$#,##0" };
    if (num[5]) return { v, fmt: "0.00E+00" };
    if (num[4].includes(",")) return { v, fmt: dec ? "#,##0.00" : "#,##0" };
    return { v };
  }
  const time = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([ap]\.?m\.?)?$/i.exec(t);
  if (time) {
    let hh = Number(time[1]); const mi = Number(time[2]), ss = Number(time[3] || 0);
    if (time[4]) { const pm = /^p/i.test(time[4]); if (hh > 12 || hh < 1) return { v: text }; hh = (hh % 12) + (pm ? 12 : 0); }
    if (hh > 23 || mi > 59 || ss > 59) return { v: text };
    return { v: (hh * 3600 + mi * 60 + ss) / 86400, fmt: time[4] ? (time[3] ? "h:mm:ss AM/PM" : "h:mm AM/PM") : time[3] ? "h:mm:ss" : "h:mm" };
  }
  const date = parseDate(t);
  if (date) return date;
  return { v: text };
}
function parseDate(t) {
  let m, y, mo, d, rest = "";
  const year = (s) => { const n = Number(s); return s.length <= 2 ? (n < 30 ? 2000 + n : 1900 + n) : n; };
  const mon = (s) => { const i = MONTHS.findIndex((x) => x.slice(0, 3).toLowerCase() === s.slice(0, 3).toLowerCase()); return i < 0 ? null : i + 1; };
  let fmt;
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(.*)$/.exec(t))) { y = Number(m[1]); mo = Number(m[2]); d = Number(m[3]); rest = m[4]; fmt = "yyyy-mm-dd"; }
  else if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})(?=\s|$)(.*)$/.exec(t))) { mo = Number(m[1]); d = Number(m[2]); y = year(m[3]); rest = m[4]; fmt = "m/d/yyyy"; }
  else if ((m = /^(\d{1,2})\/(\d{1,2})$/.exec(t))) { mo = Number(m[1]); d = Number(m[2]); y = new Date().getFullYear(); fmt = "m/d/yyyy"; }
  else if ((m = /^([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})(.*)$/.exec(t)) && mon(m[1])) { mo = mon(m[1]); d = Number(m[2]); y = Number(m[3]); rest = m[4]; fmt = "mmm d, yyyy"; }
  else if ((m = /^(\d{1,2})\s+([A-Za-z]{3,9})\.?,?\s+(\d{4})(.*)$/.exec(t)) && mon(m[2])) { d = Number(m[1]); mo = mon(m[2]); y = Number(m[3]); rest = m[4]; fmt = "d mmm yyyy"; }
  else return null;
  if (mo < 1 || mo > 12 || d < 1 || d > 31 || y < 1900 || y > 9999) return null;
  const serial = dateSerial(y, mo, d);
  if (fromSerial(serial).d !== d) return null; // 31 April and the like
  rest = rest.trim();
  if (!rest) return { v: serial, fmt };
  const tm = parseInput(rest.replace(/^t/i, ""));
  if (typeof tm.v === "number" && tm.v < 1 && tm.fmt) return { v: serial + tm.v, fmt: `${fmt} ${tm.fmt.includes("AM") ? "h:mm AM/PM" : "h:mm"}` };
  return null;
}

// ---------------------------------------------------------------- number formats
// Excel-style format codes: up to four sections (positive;negative;zero;text), 0 # ? placeholders, thousands commas,
// %, E+00, quoted text, [Red] colours, and date codes (yyyy mm dd h mm ss AM/PM, [h] for durations).
const fmtCache = new Map();
function splitSections(code) {
  const out = []; let cur = "", q = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') q = !q;
    if (ch === "\\" && !q) { cur += ch + (code[i + 1] || ""); i++; continue; }
    if (ch === ";" && !q) { out.push(cur); cur = ""; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}
function compileSection(sec) {
  let color = null, cond = null;
  sec = sec.replace(/\[(red|blue|green|black|white|yellow|magenta|cyan|color\s*\d+)\]/gi, (_, c) => { color = c.toLowerCase(); return ""; });
  sec = sec.replace(/\[([<>=]{1,2})(-?[\d.]+)\]/, (_, op, n) => { cond = { op, n: Number(n) }; return ""; });
  sec = sec.replace(/\[\$([^\]-]*)(?:-[^\]]*)?\]/g, (_, sym) => `"${sym}"`); // [$€-x] currency
  const stripped = sec.replace(/"[^"]*"|\\.|_.|\*./g, "");
  const isDate = /[yd]|[hs]|am\/pm|a\/p|\[h+\]|\[m+\]|\[s+\]/i.test(stripped) && !/[0#?]/.test(stripped.replace(/\.0+/g, ""));
  const isText = stripped.includes("@");
  return { src: sec, color, cond, isDate, isText, hasDigits: /[0#?]/.test(stripped) };
}
function getFormat(code) {
  if (!fmtCache.has(code)) {
    const secs = splitSections(code).map(compileSection);
    fmtCache.set(code, secs);
  }
  return fmtCache.get(code);
}
function literalPieces(src) {
  // turns a section into a list of pieces: {lit:"text"} or {code:"..."} for the rest
  const out = []; let buf = "";
  const flush = () => { if (buf) { out.push({ code: buf }); buf = ""; } };
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (ch === '"') { flush(); const j = src.indexOf('"', i + 1); out.push({ lit: src.slice(i + 1, j < 0 ? undefined : j) }); i = j < 0 ? src.length : j; continue; }
    if (ch === "\\") { flush(); out.push({ lit: src[i + 1] || "" }); i++; continue; }
    if (ch === "_") { flush(); out.push({ lit: " " }); i++; continue; }
    if (ch === "*") { flush(); i++; continue; }
    buf += ch;
  }
  flush();
  return out;
}
function formatNumberSection(n, sec, useAbs) {
  const sign = !useAbs && n < 0 ? "-" : "";
  let v = Math.abs(n);
  const pieces = literalPieces(sec.src);
  const code = pieces.filter((p) => p.code).map((p) => p.code).join("");
  const pct = (code.match(/%/g) || []).length;
  v *= Math.pow(100, pct);
  const sci = /E[+-]/i.exec(code);
  const numCore = code.replace(/[^0#?.,E+\-]/gi, "");
  const [intPart, decPartRaw = ""] = numCore.split(/E[+-]/i)[0].split(".");
  const decPart = decPartRaw.replace(/,/g, "");
  const decMax = (decPart.match(/[0#?]/g) || []).length, decMin = (decPart.match(/0/g) || []).length;
  const intMin = (intPart.match(/0/g) || []).length;
  const group = /[0#?],[0#?]/.test(intPart);
  const trailingCommas = (intPart.match(/,+$/) || [""])[0].length;
  if (!sci) v /= Math.pow(1000, trailingCommas);
  let body;
  if (sci) {
    const expDigits = (code.slice(sci.index + 2).match(/0/g) || []).length || 2;
    let exp = v === 0 ? 0 : Math.floor(Math.log10(Math.abs(v)));
    let mant = v / Math.pow(10, exp);
    if (Math.abs(Number(mant.toFixed(decMax))) >= 10) { mant /= 10; exp += 1; }
    body = mant.toFixed(decMax) + "E" + (exp < 0 ? "-" : "+") + String(Math.abs(exp)).padStart(expDigits, "0");
  } else {
    const neg = false;
    let s = Math.abs(roundTo(v, decMax)).toFixed(decMax);
    let [ip, dp = ""] = s.split(".");
    while (dp.length > decMin && dp.endsWith("0")) dp = dp.slice(0, -1);
    if (ip === "0" && intMin === 0) ip = "";
    ip = ip.padStart(intMin, "0");
    if (group) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    body = (neg ? "-" : "") + ip + (decMax && (dp.length || decPart.includes("0")) ? "." + dp : decPart.length && !dp.length && /[#?]/.test(decPart) ? "." : "");
    if (body === "" || body === "-") body += "0".repeat(Math.max(intMin, 0)) || (decMax ? "" : "0");
  }
  // put the number where the placeholders were, keeping literals around it
  let out = "", placed = false;
  for (const p of pieces) {
    if (p.lit != null) { out += p.lit; continue; }
    const code2 = p.code;
    const first = code2.search(/[0#?]/);
    if (first < 0) { out += code2.replace(/%/g, "%").replace(/,/g, placed ? "" : ","); continue; }
    if (placed) { out += code2.replace(/[0#?.,]|E[+-]/gi, ""); continue; }
    const before = code2.slice(0, first), afterIdx = code2.search(/[^0#?.,E+\-](?!.*[0#?])/i);
    const lastPh = Math.max(code2.lastIndexOf("0"), code2.lastIndexOf("#"), code2.lastIndexOf("?"));
    const after = code2.slice(lastPh + 1).replace(/^[.,]*/, "").replace(/^E[+-]0*/i, "");
    out += before + body + after;
    placed = true;
    void afterIdx;
  }
  out = out.replace(/^\.$/, "0");
  return sign && /[1-9]/.test(out) ? sign + out : out;
}
function formatDateSection(n, sec) {
  if (n < 0) return "#".repeat(8);
  const t = fromSerial(n);
  const pieces = literalPieces(sec.src);
  const hasAmPm = /am\/pm|a\/p/i.test(sec.src);
  let out = "";
  for (const p of pieces) {
    if (p.lit != null) { out += p.lit; continue; }
    const re = /\[h+\]|\[m+\]|\[s+\]|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p|\.0+|./gi;
    const toks = p.code.match(re) || [];
    toks.forEach((tok, i) => {
      const low = tok.toLowerCase();
      const prevTime = () => toks.slice(0, i).reverse().find((x) => /^[hs\[]/i.test(x) || /^[ymd]/i.test(x));
      const nextTime = () => toks.slice(i + 1).find((x) => /^[hs]/i.test(x) || /^[ymd]/i.test(x));
      if (low.startsWith("[h")) out += String(Math.floor(n * 24)).padStart(low.length - 2, "0");
      else if (low.startsWith("[m")) out += String(Math.floor(n * 1440)).padStart(low.length - 2, "0");
      else if (low.startsWith("[s")) out += String(Math.round(n * 86400)).padStart(low.length - 2, "0");
      else if (low === "yyyy") out += t.y;
      else if (low === "yy") out += String(t.y).slice(-2);
      else if (low === "mmmmm") out += MONTHS[t.m - 1][0];
      else if (low === "mmmm") out += MONTHS[t.m - 1];
      else if (low === "mmm") out += MONTHS[t.m - 1].slice(0, 3);
      else if (low === "mm" || low === "m") {
        const p1 = prevTime(), n1 = nextTime();
        const minute = (p1 && /^h|^\[h/i.test(p1)) || (n1 && /^s/i.test(n1));
        out += minute ? String(t.mi).padStart(low.length, "0") : String(t.m).padStart(low.length, "0");
      }
      else if (low === "dddd") out += DAYS[t.wd];
      else if (low === "ddd") out += DAYS[t.wd].slice(0, 3);
      else if (low === "dd" || low === "d") out += String(t.d).padStart(low.length, "0");
      else if (low === "hh" || low === "h") { const hh = hasAmPm ? ((t.hh + 11) % 12) + 1 : t.hh; out += String(hh).padStart(low.length, "0"); }
      else if (low === "ss" || low === "s") out += String(t.ss).padStart(low.length, "0");
      else if (low === "am/pm") out += t.hh < 12 ? "AM" : "PM";
      else if (low === "a/p") out += t.hh < 12 ? "A" : "P";
      else if (low.startsWith(".0")) out += (t.frac * 86400 % 1).toFixed(low.length - 1).slice(1);
      else out += tok;
    });
  }
  return out;
}
function generalNumber(n) {
  if (!isFinite(n)) return "#NUM!";
  const a = Math.abs(n);
  if (a !== 0 && (a >= 1e11 || a < 1e-9)) {
    const s = n.toExponential(5).replace(/\.?0+e/, "e");
    const [m, x] = s.split("e");
    return `${m}E${x[0] === "-" ? "-" : "+"}${x.replace(/^[+-]/, "").padStart(2, "0")}`;
  }
  return String(Number(n.toPrecision(11)));
}
// value + format code -> { text, color }
function format(value, code) {
  if (isErr(value)) return { text: value.code };
  if (isBlank(value)) return { text: "" };
  if (typeof value === "boolean") return { text: value ? "TRUE" : "FALSE" };
  if (typeof value === "string") {
    const text = value.startsWith("'") ? value.slice(1) : value;
    if (code) { const secs = getFormat(code); const t = secs.find((s) => s.isText) || (secs.length === 4 ? secs[3] : null); if (t) return { text: literalPieces(t.src).map((p) => (p.lit != null ? p.lit : p.code.replace(/@/g, text))).join(""), color: t.color }; }
    return { text };
  }
  if (!code || /^general$/i.test(code)) return { text: generalNumber(value) };
  if (code === "@") return { text: generalNumber(value) };
  const secs = getFormat(code);
  let sec, useAbs = false;
  const conds = secs.filter((s) => s.cond);
  if (conds.length) {
    const hit = secs.find((s) => s.cond && ({ "<": value < s.cond.n, ">": value > s.cond.n, "=": value === s.cond.n, "<=": value <= s.cond.n, ">=": value >= s.cond.n, "<>": value !== s.cond.n })[s.cond.op]);
    sec = hit || secs.find((s) => !s.cond && !s.isText) || secs[0]; useAbs = !!hit && value < 0 && hit !== secs[0];
  } else if (value < 0 && secs.length >= 2 && !secs[1].isText) { sec = secs[1]; useAbs = true; }
  else if (value === 0 && secs.length >= 3 && !secs[2].isText) sec = secs[2];
  else sec = secs[0];
  if (sec.isDate) return { text: formatDateSection(value, sec), color: sec.color };
  if (!sec.hasDigits) return { text: literalPieces(sec.src).map((p) => (p.lit != null ? p.lit : p.code.replace(/@/g, generalNumber(value)))).join("") || generalNumber(value), color: sec.color };
  return { text: formatNumberSection(value, sec, useAbs), color: sec.color };
}
const isDateFormat = (code) => !!code && getFormat(code).some((s) => s.isDate);

// ---------------------------------------------------------------- criteria for COUNTIF, SUMIF and friends
function wildcard(pattern) {
  let re = "";
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === "~" && i + 1 < pattern.length) { re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); continue; }
    re += ch === "*" ? ".*" : ch === "?" ? "." : ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(`^${re}$`, "is");
}
function criteria(c) {
  if (isErr(c)) throw c;
  if (typeof c === "number") return (v) => typeof v === "number" ? v === c : typeof v === "string" && v.trim() !== "" && !isNaN(Number(v)) && Number(v) === c;
  if (typeof c === "boolean") return (v) => v === c;
  if (isBlank(c)) return (v) => isBlank(v) || v === "";
  const s = String(c);
  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(s);
  const op = m[1] || "=", rhsText = m[2];
  const parsed = parseInput(rhsText);
  const rhs = rhsText === "" ? null : parsed.v;
  if (op === "=" || op === "<>") {
    let test;
    if (rhs === null) test = (v) => isBlank(v) || v === "";
    else if (typeof rhs === "number") test = (v) => typeof v === "number" && v === rhs;
    else if (typeof rhs === "boolean") test = (v) => v === rhs;
    else { const re = wildcard(rhsText); test = (v) => typeof v === "string" && re.test(v); }
    return op === "=" ? test : (v) => !test(v);
  }
  if (typeof rhs === "number") return (v) => typeof v === "number" && ({ "<": v < rhs, ">": v > rhs, "<=": v <= rhs, ">=": v >= rhs })[op];
  const low = rhsText.toLowerCase();
  return (v) => typeof v === "string" && ({ "<": v.toLowerCase() < low, ">": v.toLowerCase() > low, "<=": v.toLowerCase() <= low, ">=": v.toLowerCase() >= low })[op];
}

// ---------------------------------------------------------------- the workbook
const K = (si, r, c) => (si * MAX_ROWS + r) * MAX_COLS + c; // one number per cell, quicker than "A1" text keys
class Book {
  constructor(doc) { this.doc = doc; this.memo = new Map(); this.spill = new Map(); this.readSpill = this.spill; this.anchors = new Map(); this.parsed = new Map(); this.now = new Date(); this.stack = new Set(); }
  sheetIndex(name) {
    if (name == null) return null;
    const low = name.toLowerCase();
    const i = this.doc.sheets.findIndex((s) => s.name.toLowerCase() === low);
    return i < 0 ? -1 : i;
  }
  bounds(si) {
    const sh = this.doc.sheets[si];
    if (!sh._bounds) {
      let mr = -1, mc = -1;
      for (const k of Object.keys(sh.cells)) { const p = parseAddr(k); if (p && !isBlank(sh.cells[k].v) && sh.cells[k].v !== "") { if (p.r > mr) mr = p.r; if (p.c > mc) mc = p.c; } }
      sh._bounds = { r: mr, c: mc };
    }
    const b = sh._bounds;
    let r = b.r, c = b.c;
    for (const a of this.anchors.values()) if (a.si === si) { r = Math.max(r, a.r + a.h - 1); c = Math.max(c, a.c + a.w - 1); }
    return { r, c };
  }
  touched(si) { if (this.doc.sheets[si]) this.doc.sheets[si]._bounds = null; }
  ast(formula) {
    let a = this.parsed.get(formula);
    if (!a) { try { a = parse(formula.slice(1)); } catch (e) { a = { k: "lit", v: isErr(e) ? e : E.parse(String(e.message || e)) }; } this.parsed.set(formula, a); if (this.parsed.size > 20000) this.parsed.clear(); }
    return a;
  }
  // Recalculate everything. Spilled results can feed other formulas, so it settles over a few passes.
  recalc() {
    this.now = new Date();
    this.index();
    let prevSpill = this.spill;
    for (let pass = 0; pass < 4; pass++) {
      this.memo = new Map(); this.readSpill = prevSpill; this.spill = new Map(); this.anchors = new Map(); this.stack = new Set();
      this.formulas.forEach((list, si) => { for (const [r, c] of list) this.cell(si, r, c); });
      if (sameSpill(prevSpill, this.spill)) break;
      prevSpill = this.spill;
    }
    this.readSpill = this.spill;
  }
  // Every sheet's cells as rows of plain values, and its formulas in reading order, rebuilt before each recalc.
  index() {
    this.grid = []; this.formulas = [];
    this.doc.sheets.forEach((sh, si) => {
      const rows = [], list = [];
      let mr = -1, mc = -1;
      for (const k in sh.cells) {
        const p = parseAddr(k), v = sh.cells[k].v;
        if (!p || isBlank(v) || v === "") continue;
        (rows[p.r] || (rows[p.r] = []))[p.c] = v;
        if (p.r > mr) mr = p.r; if (p.c > mc) mc = p.c;
        if (typeof v === "string" && v.length > 1 && v[0] === "=") list.push([p.r, p.c]);
      }
      list.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
      sh._bounds = { r: mr, c: mc };
      this.grid.push(rows); this.formulas.push(list);
    });
  }
  raw(si, r, c) {
    if (this.grid && this.grid[si]) { const row = this.grid[si][r]; const v = row ? row[c] : undefined; return v === undefined ? null : v; }
    const sh = this.doc.sheets[si]; const cell = sh && sh.cells[addr(r, c)]; return cell ? cell.v : null;
  }
  // A cell's value: its own, a formula's result, or a value spilled into it by a formula nearby.
  cell(si, r, c) {
    const v = this.raw(si, r, c);
    if (typeof v === "string" && v.startsWith("=") && v.length > 1) {
      const key = K(si, r, c);
      const known = this.memo.get(key);
      if (known !== undefined) return known;
      if (this.stack.has(key)) return E.ref("Circular dependency: this formula refers back to itself");
      this.stack.add(key);
      let out;
      try { out = this.evaluate(this.ast(v), { si, r, c, scope: null }); out = this.settle(out, si, r, c); }
      catch (e) { out = isErr(e) ? e : e instanceof RangeError ? E.ref("The formulas refer to each other too deeply") : E.calc(String(e && e.message || e)); }
      this.stack.delete(key);
      this.memo.set(key, out);
      return out;
    }
    if (isBlank(v) || v === "") { if (!this.readSpill.size) return null; const s = this.readSpill.get(K(si, r, c)); return s === undefined ? null : s; }
    if (typeof v === "string" && v.startsWith("'")) return v.slice(1);
    return v;
  }
  // A formula's result in its cell. A block of results spills right and down into empty cells, or shows #SPILL!.
  settle(out, si, r, c) {
    if (out instanceof Ref) out = out.single ? this.cell(out.sheet, out.r1, out.c1) : this.values(out);
    if (!(out instanceof Arr)) return isBlank(out) ? 0 : out;
    if (out.h === 0 || out.w === 0) return E.calc("The result is empty");
    if (out.h === 1 && out.w === 1) { const v = out.at(0, 0); return isBlank(v) ? 0 : v; }
    if (out.h * out.w > 200000) return E.spill("The result is too big");
    for (let i = 0; i < out.h; i++) for (let j = 0; j < out.w; j++) {
      if (!i && !j) continue;
      const v = this.raw(si, r + i, c + j);
      if (!(isBlank(v) || v === "") || this.spill.has(K(si, r + i, c + j))) return E.spill(`${addr(r + i, c + j)} isn't empty, so the results can't spill`);
    }
    for (let i = 0; i < out.h; i++) for (let j = 0; j < out.w; j++) if (i || j) { const v = out.at(i, j); this.spill.set(K(si, r + i, c + j), isBlank(v) ? null : v); }
    this.anchors.set(K(si, r, c), { si, r, c, h: out.h, w: out.w });
    const first = out.at(0, 0);
    return isBlank(first) ? 0 : first;
  }
  spillAt(si, r, c) { return this.anchors.get(K(si, r, c)) || null; }
  // The formula cell whose results spilled into this cell, if any.
  spillOwner(si, r, c) { for (const a of this.anchors.values()) if (a.si === si && r >= a.r && r < a.r + a.h && c >= a.c && c < a.c + a.w) return a; return null; }
  // The value shown in a cell after recalc().
  value(si, r, c) { return this.cell(si, r, c); }
  // ---- references
  resolve(p, ctx) {
    let si = ctx.si;
    if (p.sheet != null) { si = this.sheetIndex(p.sheet); if (si < 0) throw E.ref(`There's no sheet called “${p.sheet}”`); }
    if (p.kind === "cell") return new Ref(si, p.a.r, p.a.c, p.a.r, p.a.c);
    if (p.kind === "range") return new Ref(si, p.a.r, p.a.c, p.b.r, p.b.c);
    const b = this.bounds(si);
    if (p.kind === "cols") return new Ref(si, 0, p.a.c, Math.max(0, b.r), p.b.c);
    return new Ref(si, p.a.r, 0, p.b.r, Math.max(0, b.c));
  }
  // Calls fn(value) for each cell of the range that has something in it, without building a block first.
  each(ref, fn) {
    const b = this.bounds(ref.sheet), r2 = Math.min(ref.r2, b.r), c2 = Math.min(ref.c2, b.c), g = this.grid && this.grid[ref.sheet];
    const spilled = this.readSpill.size > 0;
    for (let r = ref.r1; r <= r2; r++) {
      const row = g ? g[r] : null;
      if (!row && !spilled && g) continue;
      for (let c = ref.c1; c <= c2; c++) {
        const raw = row ? row[c] : undefined;
        let v;
        if (raw === undefined) { if (!spilled) continue; v = this.cell(ref.sheet, r, c); }
        else if (typeof raw !== "string" || !g) v = g ? raw : this.cell(ref.sheet, r, c);
        else if (raw[0] === "=") { const m = this.memo.get(K(ref.sheet, r, c)); v = m !== undefined ? m : this.cell(ref.sheet, r, c); }
        else v = raw[0] === "'" ? raw.slice(1) : raw;
        if (v !== null && v !== undefined) fn(v);
      }
    }
  }
  values(ref) {
    const rows = [];
    for (let r = ref.r1; r <= ref.r2; r++) { const row = []; for (let c = ref.c1; c <= ref.c2; c++) row.push(this.cell(ref.sheet, r, c)); rows.push(row); }
    return new Arr(rows);
  }
  // ---- evaluation. Results can be values, Arr blocks, or Ref references (for functions that need them).
  evaluate(n, ctx) {
    switch (n.k) {
      case "lit": return n.v;
      case "empty": return null;
      case "ref": return this.resolve(n.p, ctx);
      case "name": {
        for (let s = ctx.scope; s; s = s.parent) if (s.vars.has(n.v)) return s.vars.get(n.v);
        return E.name(`Unknown name “${n.v}”. Text needs "quotes"; a function needs brackets.`);
      }
      case "neg": return lift1(this.deref(this.evaluate(n.x, ctx)), (v) => -toNum(v));
      case "pct": return lift1(this.deref(this.evaluate(n.x, ctx)), (v) => toNum(v) / 100);
      case "join": {
        const a = this.evaluate(n.l, ctx), b = this.evaluate(n.r, ctx);
        if (!(a instanceof Ref) || !(b instanceof Ref) || a.sheet !== b.sheet) return E.value("Both ends of a range need to be cells on one sheet");
        return new Ref(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
      }
      case "bin": return binop(n.op, this.deref(this.evaluate(n.l, ctx)), this.deref(this.evaluate(n.r, ctx)));
      case "call": {
        const f = FN[n.name];
        if (!f) return E.name(`There's no function called ${n.name}`);
        if (f.lazy) return f.fn.call(this, n.args, ctx);
        const args = n.args.map((a) => (a.k === "empty" ? undefined : this.evaluate(a, ctx)));
        if (f.min != null && args.length < f.min) return E.value(`${n.name} needs at least ${f.min} argument${f.min === 1 ? "" : "s"}`);
        if (f.max != null && args.length > f.max) return E.value(`${n.name} takes at most ${f.max} argument${f.max === 1 ? "" : "s"}`);
        if (f.lift) return liftCall(this, f, args, ctx);
        try { return f.fn.call(this, args, ctx); } catch (e) { if (isErr(e)) return e; throw e; }
      }
      default: return E.parse();
    }
  }
  // A single-cell reference becomes its value; bigger ones become blocks of values.
  deref(v) { return v instanceof Ref ? (v.single ? this.cell(v.sheet, v.r1, v.c1) : this.values(v)) : v; }
}
function sameSpill(a, b) {
  if (a.size !== b.size) return false;
  for (const [k, v] of a) { const w = b.get(k); if (w === undefined && !b.has(k)) return false; if (!(w === v || (isErr(w) && isErr(v) && w.code === v.code))) return false; }
  return true;
}
function lift1(v, fn) {
  if (v instanceof Arr) return new Arr(v.rows.map((row) => row.map((x) => { try { return isErr(x) ? x : fn(x); } catch (e) { if (isErr(e)) return e; throw e; } })));
  if (isErr(v)) return v;
  try { return fn(v); } catch (e) { if (isErr(e)) return e; throw e; }
}
// Two blocks combined cell by cell. A single row or column stretches to match, like both apps' array maths.
function broadcast(a, b, fn) {
  const A = a instanceof Arr ? a : null, B = b instanceof Arr ? b : null;
  const h = Math.max(A ? A.h : 1, B ? B.h : 1), w = Math.max(A ? A.w : 1, B ? B.w : 1);
  const pick = (X, x, i, j) => {
    if (!X) return x;
    const ii = X.h === 1 ? 0 : i, jj = X.w === 1 ? 0 : j;
    return ii < X.h && jj < X.w ? X.at(ii, jj) : E.na("The ranges are different sizes");
  };
  const rows = [];
  for (let i = 0; i < h; i++) { const row = []; for (let j = 0; j < w; j++) row.push(fn(pick(A, a, i, j), pick(B, b, i, j))); rows.push(row); }
  return new Arr(rows);
}
function binop(op, a, b) {
  const one = (x, y) => {
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    try {
      switch (op) {
        case "+": return toNum(x) + toNum(y);
        case "-": return toNum(x) - toNum(y);
        case "*": return toNum(x) * toNum(y);
        case "/": { const d = toNum(y); if (d === 0) return E.div(); return toNum(x) / d; }
        case "^": { const r = Math.pow(toNum(x), toNum(y)); if (!isFinite(r) || isNaN(r)) return E.num(); return r; }
        case "&": return toStr(x) + toStr(y);
        case "=": return compare(x, y) === 0;
        case "<>": return compare(x, y) !== 0;
        case "<": return compare(x, y) < 0;
        case ">": return compare(x, y) > 0;
        case "<=": return compare(x, y) <= 0;
        case ">=": return compare(x, y) >= 0;
      }
    } catch (e) { if (isErr(e)) return e; throw e; }
    return E.value();
  };
  if (a instanceof Arr || b instanceof Arr) return broadcast(a, b, one);
  const out = one(a, b);
  return typeof out === "number" && (isNaN(out) || !isFinite(out)) ? E.num() : out;
}
function liftCall(book, f, args, ctx) {
  const vals = args.map((a, i) => (f.keep && f.keep.includes(i) ? a : book.deref(a)));
  const arrays = vals.filter((v, i) => v instanceof Arr && !(f.keep && f.keep.includes(i)) && !(f.arrayArgs && f.arrayArgs.includes(i)));
  const call = (xs) => {
    for (let i = 0; i < xs.length; i++) if (isErr(xs[i]) && !(f.errOk)) return xs[i];
    try { const out = f.fn.apply(book, xs); return typeof out === "number" && (isNaN(out) || !isFinite(out)) ? E.num() : out; }
    catch (e) { if (isErr(e)) return e; throw e; }
  };
  if (!arrays.length) return call(vals);
  const h = Math.max(...arrays.map((a) => a.h)), w = Math.max(...arrays.map((a) => a.w));
  const rows = [];
  for (let i = 0; i < h; i++) {
    const row = [];
    for (let j = 0; j < w; j++) row.push(call(vals.map((v, k) => (v instanceof Arr && !(f.arrayArgs && f.arrayArgs.includes(k)) ? (v.h === 1 && v.w === 1 ? v.at(0, 0) : v.at(v.h === 1 ? 0 : i, v.w === 1 ? 0 : j)) : v))));
    rows.push(row);
  }
  return new Arr(rows);
}

// ---------------------------------------------------------------- function helpers
// Everything in the arguments as a flat list. In ranges, text and TRUE/FALSE are skipped by number functions (like both
// apps); typed straight into the formula they count.
function flat(book, args, keepAll) {
  const out = [];
  for (const a of args) {
    if (a === undefined) continue;
    const v = a instanceof Ref ? book.values(a) : a;
    if (v instanceof Arr) { for (const row of v.rows) for (const x of row) out.push({ v: x, inRange: true }); }
    else out.push({ v, inRange: false });
  }
  return keepAll ? out : out.filter((x) => !isBlank(x.v));
}
function nums(book, args, opts = {}) {
  const out = [];
  const inRangeValue = (v) => {
    if (isErr(v)) throw v;
    if (typeof v === "number") out.push(v);
    else if (opts.textAsZero && (typeof v === "string" || typeof v === "boolean")) out.push(typeof v === "boolean" ? (v ? 1 : 0) : 0);
  };
  const rest = [];
  for (const a of args) {
    if (a instanceof Ref) book.each(a, inRangeValue);
    else if (a instanceof Arr) { for (const row of a.rows) for (const v of row) if (!isBlank(v)) inRangeValue(v); }
    else rest.push(a);
  }
  for (const { v, inRange } of flat(book, rest)) {
    if (isErr(v)) throw v;
    if (typeof v === "number") out.push(v);
    else if (inRange) { if (opts.textAsZero && (typeof v === "string" || typeof v === "boolean")) out.push(typeof v === "boolean" ? (v ? 1 : 0) : 0); }
    else out.push(toNum(v));
  }
  return out;
}
function block(book, v) {
  if (v instanceof Ref) return book.values(v);
  if (v instanceof Arr) return v;
  return new Arr([[v === undefined ? null : v]]);
}
const cells = (A) => { const out = []; for (const row of A.rows) for (const x of row) out.push(x); return out; };
const sum = (xs) => xs.reduce((a, b) => a + b, 0);
const mean = (xs) => { if (!xs.length) throw E.div("No numbers to average"); return sum(xs) / xs.length; };
function variance(xs, sample) { if (xs.length < (sample ? 2 : 1)) throw E.div(); const m = mean(xs); return sum(xs.map((x) => (x - m) ** 2)) / (xs.length - (sample ? 1 : 0)); }
function percentile(xs, p, exclusive) {
  if (!xs.length) throw E.num(); const s = [...xs].sort((a, b) => a - b), n = s.length;
  if (p < 0 || p > 1) throw E.num();
  let rank;
  if (exclusive) { rank = p * (n + 1) - 1; if (rank < 0 || rank > n - 1) throw E.num(); } else rank = p * (n - 1);
  const lo = Math.floor(rank), hi = Math.ceil(rank);
  return s[lo] + (s[hi] - s[lo]) * (rank - lo);
}
function roundTo(x, d) { const f = Math.pow(10, d); const y = Math.abs(x) * f; const r = Math.round(y + 1e-9 * Math.max(1, y)) / f; return x < 0 ? -r : r; }
// pairs of (range, criteria) as one test per cell position, for the *IFS functions
function ifsMask(book, pairs, shape) {
  let mask = null;
  for (let i = 0; i < pairs.length; i += 2) {
    const A = block(book, pairs[i]);
    const test = criteria(book.deref(pairs[i + 1]) instanceof Arr ? book.deref(pairs[i + 1]).at(0, 0) : book.deref(pairs[i + 1]));
    if (shape && (A.h !== shape.h || A.w !== shape.w)) throw E.value("Every range needs to be the same size");
    shape = shape || A;
    const m = cells(A).map(test);
    mask = mask ? mask.map((x, k) => x && m[k]) : m;
  }
  return mask;
}
function lookupScan(list, value, mode) {
  // mode 0 exact (wildcards for text), 1 next smaller, -1 next larger; list is plain values
  if (mode === 0) {
    const re = typeof value === "string" && /[*?~]/.test(value) ? wildcard(value) : null;
    for (let i = 0; i < list.length; i++) { const x = list[i]; if (re ? typeof x === "string" && re.test(x) : !isBlank(x) && typeTie(x, value) && compare(x, value) === 0) return i; }
    return -1;
  }
  let best = -1;
  for (let i = 0; i < list.length; i++) {
    const x = list[i];
    if (isBlank(x) || isErr(x) || !typeTie(x, value)) continue;
    const c = compare(x, value);
    if (c === 0) return i;
    if (mode === 1 && c < 0 && (best < 0 || compare(x, list[best]) > 0)) best = i;
    if (mode === -1 && c > 0 && (best < 0 || compare(x, list[best]) < 0)) best = i;
  }
  return best;
}
const typeTie = (x, v) => typeRank(x) === typeRank(v);
function sortedLookup(list, value) {
  // like MATCH type 1 / VLOOKUP TRUE: the last position whose value is <= value, assuming sorted
  let lo = 0, hi = list.length - 1, ans = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1, x = list[mid];
    if (isBlank(x) || !typeTie(x, value)) { // skip over odd types like both apps do, by linear fallback
      return lookupScan(list, value, 1);
    }
    if (compare(x, value) <= 0) { ans = mid; lo = mid + 1; } else hi = mid - 1;
  }
  return ans;
}
function vec(A) { return A.h === 1 ? A.rows[0] : A.rows.map((r) => r[0]); }

// ---------------------------------------------------------------- the functions
// def(name, args, category, description, fn, options). With lift, each argument is a plain value and a block of values
// runs the function once per cell (so =LEN(A1:A5) spills five answers).
const FN = {}, CATALOG = [];
function def(name, args, cat, desc, fn, opts = {}) {
  FN[name] = { fn, ...opts };
  CATALOG.push({ name, args, cat, desc });
}
const L = { lift: true };
const D = (o = {}) => ({ lift: true, ...o });

// ---- Math
def("SUM", "value1, [value2, …]", "Math", "Adds numbers and ranges.", function (a) { return sum(nums(this, a)); }, { min: 1 });
def("PRODUCT", "value1, [value2, …]", "Math", "Multiplies numbers together.", function (a) { const x = nums(this, a); return x.length ? x.reduce((p, v) => p * v, 1) : 0; }, { min: 1 });
def("SUMSQ", "value1, [value2, …]", "Math", "The sum of the squares.", function (a) { return sum(nums(this, a).map((x) => x * x)); });
def("SUMIF", "range, criterion, [sum_range]", "Math", "Adds the cells that match a condition, like \">10\" or \"apples\".", function ([range, crit, sr]) {
  const A = block(this, range), test = criteria(this.deref(crit));
  const S = sr === undefined ? A : (sr instanceof Ref ? this.values(new Ref(sr.sheet, sr.r1, sr.c1, sr.r1 + A.h - 1, sr.c1 + A.w - 1)) : block(this, sr));
  let t = 0;
  for (let i = 0; i < A.h; i++) for (let j = 0; j < A.w; j++) if (test(A.at(i, j))) { const v = S.at(i, j); if (isErr(v)) throw v; if (typeof v === "number") t += v; }
  return t;
}, { min: 2, max: 3 });
def("SUMIFS", "sum_range, criteria_range1, criterion1, …", "Math", "Adds the cells that match every condition.", function ([sr, ...pairs]) {
  const S = block(this, sr); const mask = ifsMask(this, pairs, S);
  return sum(cells(S).filter((v, k) => mask[k] && typeof v === "number"));
}, { min: 3 });
def("SUMPRODUCT", "array1, [array2, …]", "Math", "Multiplies matching cells of the ranges and adds the results.", function (a) {
  const B = a.map((x) => block(this, this.deref(x) instanceof Arr ? this.deref(x) : x));
  if (B.some((x) => x.h !== B[0].h || x.w !== B[0].w)) throw E.value("The ranges need to be the same size");
  let t = 0;
  for (let i = 0; i < B[0].h; i++) for (let j = 0; j < B[0].w; j++) {
    let p = 1;
    for (const X of B) { const v = X.at(i, j); if (isErr(v)) throw v; p *= typeof v === "number" ? v : typeof v === "boolean" ? (v ? 1 : 0) : 0; }
    t += p;
  }
  return t;
}, { min: 1 });
def("ABS", "value", "Math", "The number without its sign.", (x) => Math.abs(toNum(x)), L);
def("ROUND", "value, [places]", "Math", "Rounds to a number of decimal places (halves away from zero).", (x, d) => roundTo(toNum(x), toInt(d ?? 0)), L);
def("ROUNDUP", "value, [places]", "Math", "Rounds away from zero.", (x, d) => { const f = Math.pow(10, toInt(d ?? 0)), v = toNum(x); return Math.sign(v) * Math.ceil(Number((Math.abs(v) * f).toPrecision(15))) / f; }, L);
def("ROUNDDOWN", "value, [places]", "Math", "Rounds toward zero.", (x, d) => { const f = Math.pow(10, toInt(d ?? 0)), v = toNum(x); return Math.sign(v) * Math.floor(Number((Math.abs(v) * f).toPrecision(15))) / f; }, L);
def("MROUND", "value, factor", "Math", "Rounds to the nearest multiple.", (x, m) => { const v = toNum(x), f = toNum(m); if (f === 0) return 0; if (Math.sign(v) * Math.sign(f) < 0) throw E.num(); return roundTo(v / f, 0) * f; }, L);
def("INT", "value", "Math", "Rounds down to a whole number.", (x) => Math.floor(toNum(x)), L);
def("TRUNC", "value, [places]", "Math", "Cuts off the decimals.", (x, d) => { const f = Math.pow(10, toInt(d ?? 0)); return Math.trunc(toNum(x) * f) / f; }, L);
def("MOD", "dividend, divisor", "Math", "The remainder after dividing (same sign as the divisor).", (a, b) => { const d = toNum(b); if (d === 0) throw E.div(); const n = toNum(a); return n - d * Math.floor(n / d); }, L);
def("QUOTIENT", "dividend, divisor", "Math", "The whole-number part of a division.", (a, b) => { const d = toNum(b); if (d === 0) throw E.div(); return Math.trunc(toNum(a) / d); }, L);
def("POWER", "base, exponent", "Math", "A number raised to a power.", (a, b) => Math.pow(toNum(a), toNum(b)), L);
def("SQRT", "value", "Math", "The square root.", (x) => { const v = toNum(x); if (v < 0) throw E.num("Can't take the square root of a negative number"); return Math.sqrt(v); }, L);
def("SQRTPI", "value", "Math", "The square root of the number times pi.", (x) => { const v = toNum(x); if (v < 0) throw E.num(); return Math.sqrt(v * Math.PI); }, L);
def("EXP", "value", "Math", "e raised to the power.", (x) => Math.exp(toNum(x)), L);
def("LN", "value", "Math", "The natural logarithm.", (x) => { const v = toNum(x); if (v <= 0) throw E.num(); return Math.log(v); }, L);
def("LOG", "value, [base]", "Math", "The logarithm, base 10 unless given.", (x, b) => { const v = toNum(x), base = b === undefined ? 10 : toNum(b); if (v <= 0 || base <= 0 || base === 1) throw E.num(); return Math.log(v) / Math.log(base); }, L);
def("LOG10", "value", "Math", "The base-10 logarithm.", (x) => { const v = toNum(x); if (v <= 0) throw E.num(); return Math.log10(v); }, L);
def("PI", "", "Math", "The number pi.", () => Math.PI, { max: 0 });
def("SIGN", "value", "Math", "1, 0 or -1, by the sign of the number.", (x) => Math.sign(toNum(x)), L);
def("CEILING", "value, [factor]", "Math", "Rounds up to a multiple.", (x, f) => { const v = toNum(x), m = f === undefined ? 1 : toNum(f); if (m === 0) return 0; return Math.ceil(Number((v / m).toPrecision(15))) * m; }, L);
def("CEILING.MATH", "value, [factor], [mode]", "Math", "Rounds up to a multiple (negatives toward zero unless mode is set).", (x, f, mode) => { const v = toNum(x), m = Math.abs(f === undefined ? 1 : toNum(f)); if (m === 0) return 0; return v < 0 && mode && toNum(mode) ? -Math.ceil(-v / m) * m : Math.ceil(Number((v / m).toPrecision(15))) * m; }, L);
def("FLOOR", "value, [factor]", "Math", "Rounds down to a multiple.", (x, f) => { const v = toNum(x), m = f === undefined ? 1 : toNum(f); if (m === 0) throw E.div(); return Math.floor(Number((v / m).toPrecision(15))) * m; }, L);
def("FLOOR.MATH", "value, [factor], [mode]", "Math", "Rounds down to a multiple.", (x, f, mode) => { const v = toNum(x), m = Math.abs(f === undefined ? 1 : toNum(f)); if (m === 0) return 0; return v < 0 && mode && toNum(mode) ? -Math.floor(-v / m) * m : Math.floor(Number((v / m).toPrecision(15))) * m; }, L);
def("EVEN", "value", "Math", "Rounds away from zero to an even number.", (x) => { const v = toNum(x), r = Math.ceil(Math.abs(v) / 2) * 2; return v < 0 ? -r : r; }, L);
def("ODD", "value", "Math", "Rounds away from zero to an odd number.", (x) => { const v = toNum(x); let r = Math.ceil(Math.abs(v)); if (r % 2 === 0) r += 1; return v < 0 ? -r : r; }, L);
def("FACT", "value", "Math", "The factorial (5! = 120).", (x) => { const n = Math.floor(toNum(x)); if (n < 0) throw E.num(); let r = 1; for (let i = 2; i <= n; i++) r *= i; return r; }, L);
def("COMBIN", "n, k", "Math", "How many ways to choose k from n.", (a, b) => { const n = Math.floor(toNum(a)), k = Math.floor(toNum(b)); if (n < 0 || k < 0 || k > n) throw E.num(); let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return Math.round(r); }, L);
def("PERMUT", "n, k", "Math", "How many ordered ways to pick k from n.", (a, b) => { const n = Math.floor(toNum(a)), k = Math.floor(toNum(b)); if (n < 0 || k < 0 || k > n) throw E.num(); let r = 1; for (let i = 0; i < k; i++) r *= n - i; return r; }, L);
def("GCD", "value1, [value2, …]", "Math", "The greatest common divisor.", function (a) { const g = (x, y) => (y ? g(y, x % y) : x); return nums(this, a).map((x) => Math.floor(x)).reduce((p, v) => { if (v < 0) throw E.num(); return g(p, v); }, 0); });
def("LCM", "value1, [value2, …]", "Math", "The least common multiple.", function (a) { const g = (x, y) => (y ? g(y, x % y) : x); return nums(this, a).map((x) => Math.floor(x)).reduce((p, v) => { if (v < 0) throw E.num(); return v === 0 || p === 0 ? 0 : (p * v) / g(p, v); }, 1); });
def("RAND", "", "Math", "A random number from 0 up to 1. Changes on every edit.", () => Math.random(), { max: 0 });
def("RANDBETWEEN", "low, high", "Math", "A random whole number between two numbers.", (a, b) => { const lo = Math.ceil(toNum(a)), hi = Math.floor(toNum(b)); if (lo > hi) throw E.num(); return lo + Math.floor(Math.random() * (hi - lo + 1)); }, L);
def("RANDARRAY", "[rows], [columns], [min], [max], [whole_number]", "Math", "A block of random numbers.", (r, c, lo, hi, whole) => {
  const h = r === undefined ? 1 : toInt(r), w = c === undefined ? 1 : toInt(c), a = lo === undefined ? 0 : toNum(lo), b = hi === undefined ? 1 : toNum(hi), int = whole !== undefined && toBool(whole);
  if (h < 1 || w < 1) throw E.value();
  return new Arr(Array.from({ length: h }, () => Array.from({ length: w }, () => (int ? Math.ceil(a) + Math.floor(Math.random() * (Math.floor(b) - Math.ceil(a) + 1)) : a + Math.random() * (b - a)))));
}, L);
def("SEQUENCE", "rows, [columns], [start], [step]", "Math", "A block of numbers counting up, like 1, 2, 3…", (r, c, s, st) => {
  const h = toInt(r), w = c === undefined ? 1 : toInt(c), start = s === undefined ? 1 : toNum(s), step = st === undefined ? 1 : toNum(st);
  if (h < 1 || w < 1) throw E.calc("SEQUENCE needs at least 1 row and 1 column");
  return new Arr(Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => start + (i * w + j) * step)));
}, L);
def("SUBTOTAL", "function_code, range1, [range2, …]", "Math", "A total by code: 1 average, 2 count, 3 counta, 4 max, 5 min, 6 product, 9 sum (101-111 the same).", function ([code, ...rest]) {
  const k = toInt(this.deref(code)) % 100;
  const map = { 1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN", 6: "PRODUCT", 7: "STDEV", 8: "STDEVP", 9: "SUM", 10: "VAR", 11: "VARP" };
  if (!map[k]) throw E.value("Unknown SUBTOTAL code");
  return FN[map[k]].fn.call(this, rest);
}, { min: 2 });
def("DEGREES", "radians", "Math", "Radians to degrees.", (x) => (toNum(x) * 180) / Math.PI, L);
def("RADIANS", "degrees", "Math", "Degrees to radians.", (x) => (toNum(x) * Math.PI) / 180, L);
for (const [n, f, d] of [["SIN", Math.sin, "The sine of an angle in radians."], ["COS", Math.cos, "The cosine."], ["TAN", Math.tan, "The tangent."],
  ["ASIN", Math.asin, "The arcsine, in radians."], ["ACOS", Math.acos, "The arccosine, in radians."], ["ATAN", Math.atan, "The arctangent, in radians."],
  ["SINH", Math.sinh, "The hyperbolic sine."], ["COSH", Math.cosh, "The hyperbolic cosine."], ["TANH", Math.tanh, "The hyperbolic tangent."]]) def(n, "value", "Math", d, (x) => f(toNum(x)), L);
def("ATAN2", "x, y", "Math", "The angle of the point (x, y), in radians.", (x, y) => { const a = toNum(x), b = toNum(y); if (a === 0 && b === 0) throw E.div(); return Math.atan2(b, a); }, L);
def("ROMAN", "value", "Math", "A number as Roman numerals.", (x) => { let n = toInt(x); if (n < 0 || n > 3999) throw E.value(); const t = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]]; let s = ""; for (const [v, r] of t) while (n >= v) { s += r; n -= v; } return s; }, L);
def("BASE", "value, radix, [min_length]", "Math", "A number written in another base, like 2 for binary.", (x, b, m) => { const n = toInt(x), r = toInt(b); if (n < 0 || r < 2 || r > 36) throw E.num(); return n.toString(r).toUpperCase().padStart(m === undefined ? 0 : toInt(m), "0"); }, L);
def("DECIMAL", "text, radix", "Math", "Text in another base as a number.", (t, b) => { const r = toInt(b), v = parseInt(toStr(t), r); if (isNaN(v)) throw E.num(); return v; }, L);

// ---- Statistics
def("AVERAGE", "value1, [value2, …]", "Statistics", "The average (mean) of the numbers.", function (a) { return mean(nums(this, a)); }, { min: 1 });
def("AVERAGEA", "value1, [value2, …]", "Statistics", "The average, counting text as 0 and TRUE as 1.", function (a) { return mean(nums(this, a, { textAsZero: true })); }, { min: 1 });
def("AVERAGEIF", "range, criterion, [average_range]", "Statistics", "The average of the cells that match a condition.", function ([range, crit, ar]) {
  const A = block(this, range), test = criteria(this.deref(crit)), S = ar === undefined ? A : block(this, ar);
  const xs = []; for (let i = 0; i < A.h; i++) for (let j = 0; j < A.w; j++) if (test(A.at(i, j))) { const v = S.at(i, j); if (typeof v === "number") xs.push(v); }
  return mean(xs);
}, { min: 2, max: 3 });
def("AVERAGEIFS", "average_range, criteria_range1, criterion1, …", "Statistics", "The average of the cells that match every condition.", function ([sr, ...pairs]) {
  const S = block(this, sr), mask = ifsMask(this, pairs, S); return mean(cells(S).filter((v, k) => mask[k] && typeof v === "number"));
}, { min: 3 });
def("COUNT", "value1, [value2, …]", "Statistics", "How many numbers there are.", function (a) {
  let n = 0; for (const { v, inRange } of flat(this, a)) { if (typeof v === "number") n++; else if (!inRange && !isErr(v)) { try { toNum(v); n++; } catch (e) { /* not a number */ } } } return n;
});
def("COUNTA", "value1, [value2, …]", "Statistics", "How many cells aren't empty.", function (a) { return flat(this, a).filter((x) => !(x.inRange && x.v === "")).length; });
def("COUNTBLANK", "range", "Statistics", "How many cells are empty.", function ([r]) { return cells(block(this, r)).filter((v) => isBlank(v) || v === "").length; }, { min: 1, max: 1 });
def("COUNTIF", "range, criterion", "Statistics", "How many cells match a condition, like \">10\" or \"*apple*\".", function ([r, c]) { const t = criteria(this.deref(c) instanceof Arr ? this.deref(c).at(0, 0) : this.deref(c)); return cells(block(this, r)).filter(t).length; }, { min: 2, max: 2 });
def("COUNTIFS", "criteria_range1, criterion1, …", "Statistics", "How many rows match every condition.", function (pairs) { return ifsMask(this, pairs).filter(Boolean).length; }, { min: 2 });
def("COUNTUNIQUE", "value1, [value2, …]", "Statistics", "How many different values there are.", function (a) { return new Set(flat(this, a).map((x) => (typeof x.v === "string" ? "s" + x.v.toLowerCase() : typeof x.v + x.v))).size; });
def("MAX", "value1, [value2, …]", "Statistics", "The largest number.", function (a) { const x = nums(this, a); return x.length ? Math.max(...x) : 0; }, { min: 1 });
def("MIN", "value1, [value2, …]", "Statistics", "The smallest number.", function (a) { const x = nums(this, a); return x.length ? Math.min(...x) : 0; }, { min: 1 });
def("MAXA", "value1, [value2, …]", "Statistics", "The largest value, counting text as 0.", function (a) { const x = nums(this, a, { textAsZero: true }); return x.length ? Math.max(...x) : 0; });
def("MINA", "value1, [value2, …]", "Statistics", "The smallest value, counting text as 0.", function (a) { const x = nums(this, a, { textAsZero: true }); return x.length ? Math.min(...x) : 0; });
def("MAXIFS", "max_range, criteria_range1, criterion1, …", "Statistics", "The largest number where every condition matches.", function ([sr, ...pairs]) { const S = block(this, sr), m = ifsMask(this, pairs, S); const x = cells(S).filter((v, k) => m[k] && typeof v === "number"); return x.length ? Math.max(...x) : 0; }, { min: 3 });
def("MINIFS", "min_range, criteria_range1, criterion1, …", "Statistics", "The smallest number where every condition matches.", function ([sr, ...pairs]) { const S = block(this, sr), m = ifsMask(this, pairs, S); const x = cells(S).filter((v, k) => m[k] && typeof v === "number"); return x.length ? Math.min(...x) : 0; }, { min: 3 });
def("MEDIAN", "value1, [value2, …]", "Statistics", "The middle number.", function (a) { const s = nums(this, a).sort((x, y) => x - y); if (!s.length) throw E.num(); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; }, { min: 1 });
def("MODE", "value1, [value2, …]", "Statistics", "The most common number.", function (a) { const x = nums(this, a), c = new Map(); let best = null, bc = 1; for (const v of x) { const n = (c.get(v) || 0) + 1; c.set(v, n); if (n > bc) { bc = n; best = v; } } if (best === null) throw E.na("No number repeats"); return best; });
FN["MODE.SNGL"] = FN.MODE;
def("STDEV", "value1, [value2, …]", "Statistics", "Standard deviation of a sample.", function (a) { return Math.sqrt(variance(nums(this, a), true)); });
def("STDEVP", "value1, [value2, …]", "Statistics", "Standard deviation of a whole population.", function (a) { return Math.sqrt(variance(nums(this, a), false)); });
def("VAR", "value1, [value2, …]", "Statistics", "Variance of a sample.", function (a) { return variance(nums(this, a), true); });
def("VARP", "value1, [value2, …]", "Statistics", "Variance of a whole population.", function (a) { return variance(nums(this, a), false); });
FN["STDEV.S"] = FN.STDEV; FN["STDEV.P"] = FN.STDEVP; FN["VAR.S"] = FN.VAR; FN["VAR.P"] = FN.VARP;
CATALOG.push({ name: "STDEV.S", args: "value1, [value2, …]", cat: "Statistics", desc: "Same as STDEV." }, { name: "STDEV.P", args: "value1, [value2, …]", cat: "Statistics", desc: "Same as STDEVP." },
  { name: "VAR.S", args: "value1, [value2, …]", cat: "Statistics", desc: "Same as VAR." }, { name: "VAR.P", args: "value1, [value2, …]", cat: "Statistics", desc: "Same as VARP." });
def("LARGE", "range, k", "Statistics", "The k-th largest number.", function ([r, k]) { const s = nums(this, [r]).sort((x, y) => y - x), i = toInt(this.deref(k)); if (i < 1 || i > s.length) throw E.num(); return s[i - 1]; }, { min: 2, max: 2 });
def("SMALL", "range, k", "Statistics", "The k-th smallest number.", function ([r, k]) { const s = nums(this, [r]).sort((x, y) => x - y), i = toInt(this.deref(k)); if (i < 1 || i > s.length) throw E.num(); return s[i - 1]; }, { min: 2, max: 2 });
def("RANK", "value, range, [ascending]", "Statistics", "The position of a number in a list (1 = largest unless ascending).", function ([v, r, o]) {
  const x = toNum(this.deref(v)), s = nums(this, [r]), asc = o !== undefined && toBool(this.deref(o));
  if (!s.includes(x)) throw E.na("The number isn't in the list");
  return 1 + s.filter((y) => (asc ? y < x : y > x)).length;
}, { min: 2, max: 3 });
FN["RANK.EQ"] = FN.RANK;
def("RANK.AVG", "value, range, [ascending]", "Statistics", "The position, averaging ties.", function ([v, r, o]) {
  const x = toNum(this.deref(v)), s = nums(this, [r]), asc = o !== undefined && toBool(this.deref(o));
  const ties = s.filter((y) => y === x).length; if (!ties) throw E.na();
  return 1 + s.filter((y) => (asc ? y < x : y > x)).length + (ties - 1) / 2;
});
def("PERCENTILE", "range, percentile", "Statistics", "The value at a percentile (0 to 1).", function ([r, p]) { return percentile(nums(this, [r]), toNum(this.deref(p))); }, { min: 2, max: 2 });
FN["PERCENTILE.INC"] = FN.PERCENTILE;
def("PERCENTILE.EXC", "range, percentile", "Statistics", "The value at a percentile, excluding the ends.", function ([r, p]) { return percentile(nums(this, [r]), toNum(this.deref(p)), true); });
def("QUARTILE", "range, quart", "Statistics", "A quartile: 0 min, 1, 2 median, 3, 4 max.", function ([r, q]) { const k = toInt(this.deref(q)); if (k < 0 || k > 4) throw E.num(); return percentile(nums(this, [r]), k / 4); });
FN["QUARTILE.INC"] = FN.QUARTILE;
def("PERCENTRANK", "range, value, [significance]", "Statistics", "Where a value sits in a list, from 0 to 1.", function ([r, v, d]) {
  const s = nums(this, [r]).sort((a, b) => a - b), x = toNum(this.deref(v)), places = d === undefined ? 3 : toInt(this.deref(d));
  if (!s.length || x < s[0] || x > s[s.length - 1]) throw E.na();
  let i = s.findIndex((y) => y >= x); if (s[i] !== x) i = i - 1 + (x - s[i - 1]) / (s[i] - s[i - 1]);
  return Math.floor((i / (s.length - 1 || 1)) * Math.pow(10, places)) / Math.pow(10, places);
});
def("CORREL", "range1, range2", "Statistics", "How closely two lists move together (-1 to 1).", function ([a, b]) {
  const [x, y] = pairNums(this, a, b); const mx = mean(x), my = mean(y);
  const num = sum(x.map((v, i) => (v - mx) * (y[i] - my))), den = Math.sqrt(sum(x.map((v) => (v - mx) ** 2)) * sum(y.map((v) => (v - my) ** 2)));
  if (!den) throw E.div(); return num / den;
});
FN.PEARSON = FN.CORREL;
def("COVARIANCE.P", "range1, range2", "Statistics", "Population covariance.", function ([a, b]) { const [x, y] = pairNums(this, a, b), mx = mean(x), my = mean(y); return sum(x.map((v, i) => (v - mx) * (y[i] - my))) / x.length; });
def("COVARIANCE.S", "range1, range2", "Statistics", "Sample covariance.", function ([a, b]) { const [x, y] = pairNums(this, a, b), mx = mean(x), my = mean(y); if (x.length < 2) throw E.div(); return sum(x.map((v, i) => (v - mx) * (y[i] - my))) / (x.length - 1); });
def("SLOPE", "known_y, known_x", "Statistics", "The slope of the best straight line.", function ([ys, xs]) { const [y, x] = pairNums(this, ys, xs), mx = mean(x), my = mean(y), d = sum(x.map((v) => (v - mx) ** 2)); if (!d) throw E.div(); return sum(x.map((v, i) => (v - mx) * (y[i] - my))) / d; });
def("INTERCEPT", "known_y, known_x", "Statistics", "Where the best straight line crosses zero.", function ([ys, xs]) { const s = FN.SLOPE.fn.call(this, [ys, xs]); const [y, x] = pairNums(this, ys, xs); return mean(y) - s * mean(x); });
def("FORECAST", "x, known_y, known_x", "Statistics", "Predicts a value on the best straight line.", function ([x, ys, xs]) { const s = FN.SLOPE.fn.call(this, [ys, xs]), i = FN.INTERCEPT.fn.call(this, [ys, xs]); return i + s * toNum(this.deref(x)); });
FN["FORECAST.LINEAR"] = FN.FORECAST;
def("RSQ", "known_y, known_x", "Statistics", "R squared of the best straight line.", function ([ys, xs]) { const r = FN.CORREL.fn.call(this, [ys, xs]); return r * r; });
def("GEOMEAN", "value1, [value2, …]", "Statistics", "The geometric mean.", function (a) { const x = nums(this, a); if (!x.length || x.some((v) => v <= 0)) throw E.num(); return Math.exp(sum(x.map(Math.log)) / x.length); });
def("HARMEAN", "value1, [value2, …]", "Statistics", "The harmonic mean.", function (a) { const x = nums(this, a); if (!x.length || x.some((v) => v <= 0)) throw E.num(); return x.length / sum(x.map((v) => 1 / v)); });
def("AVEDEV", "value1, [value2, …]", "Statistics", "The average distance from the mean.", function (a) { const x = nums(this, a), m = mean(x); return mean(x.map((v) => Math.abs(v - m))); });
def("STANDARDIZE", "x, mean, standard_dev", "Statistics", "How many standard deviations x is from the mean.", (x, m, s) => { const d = toNum(s); if (d <= 0) throw E.num(); return (toNum(x) - toNum(m)) / d; }, L);
def("NORM.DIST", "x, mean, standard_dev, cumulative", "Statistics", "The normal distribution.", (x, m, s, c) => { const z = (toNum(x) - toNum(m)) / toNum(s); if (toNum(s) <= 0) throw E.num(); return toBool(c) ? 0.5 * (1 + erf(z / Math.SQRT2)) : Math.exp(-z * z / 2) / (toNum(s) * Math.sqrt(2 * Math.PI)); }, L);
def("NORM.S.DIST", "z, cumulative", "Statistics", "The standard normal distribution.", (z, c) => { const v = toNum(z); return toBool(c) ? 0.5 * (1 + erf(v / Math.SQRT2)) : Math.exp(-v * v / 2) / Math.sqrt(2 * Math.PI); }, L);
def("NORM.INV", "probability, mean, standard_dev", "Statistics", "The inverse of the normal distribution.", (p, m, s) => { const q = toNum(p); if (q <= 0 || q >= 1 || toNum(s) <= 0) throw E.num(); return toNum(m) + toNum(s) * normInv(q); }, L);
def("NORM.S.INV", "probability", "Statistics", "The inverse of the standard normal distribution.", (p) => { const q = toNum(p); if (q <= 0 || q >= 1) throw E.num(); return normInv(q); }, L);
FN.NORMDIST = FN["NORM.DIST"]; FN.NORMINV = FN["NORM.INV"]; FN.NORMSDIST = { ...FN["NORM.S.DIST"], fn: (z) => 0.5 * (1 + erf(toNum(z) / Math.SQRT2)) }; FN.NORMSINV = FN["NORM.S.INV"];
function erf(x) { const t = 1 / (1 + 0.3275911 * Math.abs(x)); const y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-x * x); return x >= 0 ? y : -y; }
function normInv(p) {
  // Acklam's approximation
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  const pl = 0.02425;
  if (p < pl) { const q = Math.sqrt(-2 * Math.log(p)); return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
  if (p > 1 - pl) { const q = Math.sqrt(-2 * Math.log(1 - p)); return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
  const q = p - 0.5, r = q * q;
  return (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
}
function pairNums(book, a, b) {
  const A = cells(block(book, a)), B = cells(block(book, b));
  if (A.length !== B.length) throw E.na("The ranges need to be the same size");
  const x = [], y = [];
  A.forEach((v, i) => { if (isErr(v)) throw v; if (isErr(B[i])) throw B[i]; if (typeof v === "number" && typeof B[i] === "number") { x.push(v); y.push(B[i]); } });
  return [x, y];
}

// ---- Logical (these look at only the arguments they need, so =IF(A1=0, "-", 1/A1) never shows #DIV/0!)
def("IF", "condition, value_if_true, [value_if_false]", "Logical", "One value when the condition is true, another when it's false.", function (args, ctx) {
  if (args.length < 2 || args.length > 3) return E.value("IF needs 2 or 3 arguments");
  const cond = this.deref(this.evaluate(args[0], ctx));
  const pick = (b) => (b ? (args[1].k === "empty" ? 0 : this.evaluate(args[1], ctx)) : args.length < 3 ? false : args[2].k === "empty" ? 0 : this.evaluate(args[2], ctx));
  if (cond instanceof Arr) {
    const t = args[1].k === "empty" ? 0 : this.deref(this.evaluate(args[1], ctx)), f = args.length < 3 ? false : args[2].k === "empty" ? 0 : this.deref(this.evaluate(args[2], ctx));
    return broadcast(broadcast(cond, t, (c, tv) => ({ c, tv })), f, (ct, fv) => { if (isErr(ct)) return ct; if (isErr(ct.c)) return ct.c; try { return toBool(ct.c) ? ct.tv : fv; } catch (e) { return e; } });
  }
  if (isErr(cond)) return cond;
  try { return pick(toBool(cond)); } catch (e) { if (isErr(e)) return e; throw e; }
}, { lazy: true });
def("IFS", "condition1, value1, [condition2, value2, …]", "Logical", "The value for the first condition that's true.", function (args, ctx) {
  if (args.length < 2 || args.length % 2) return E.value("IFS needs pairs of condition and value");
  for (let i = 0; i < args.length; i += 2) { const c = this.deref(this.evaluate(args[i], ctx)); if (isErr(c)) return c; try { if (toBool(c instanceof Arr ? c.at(0, 0) : c)) return this.evaluate(args[i + 1], ctx); } catch (e) { return e; } }
  return E.na("No condition was true");
}, { lazy: true });
def("IFERROR", "value, value_if_error", "Logical", "The value, or something else if it's an error.", function (args, ctx) {
  const v = this.deref(this.evaluate(args[0], ctx));
  const alt = () => (args.length > 1 && args[1].k !== "empty" ? this.deref(this.evaluate(args[1], ctx)) : "");
  if (v instanceof Arr) { const a = alt(); return new Arr(v.rows.map((r) => r.map((x) => (isErr(x) ? (a instanceof Arr ? a.at(0, 0) : a) : x)))); }
  return isErr(v) ? alt() : v;
}, { lazy: true });
def("IFNA", "value, value_if_na", "Logical", "The value, or something else if it's #N/A.", function (args, ctx) {
  const v = this.deref(this.evaluate(args[0], ctx));
  const alt = () => (args.length > 1 ? this.deref(this.evaluate(args[1], ctx)) : "");
  if (v instanceof Arr) { const a = alt(); return new Arr(v.rows.map((r) => r.map((x) => (isErr(x) && x.code === "#N/A" ? a : x)))); }
  return isErr(v) && v.code === "#N/A" ? alt() : v;
}, { lazy: true });
function logicals(book, args) {
  const out = [];
  for (const { v, inRange } of flat(book, args)) {
    if (isErr(v)) throw v;
    if (typeof v === "boolean" || typeof v === "number") out.push(toBool(v));
    else if (!inRange && typeof v === "string") out.push(toBool(v));
  }
  if (!out.length) throw E.value("No TRUE or FALSE values to check");
  return out;
}
def("AND", "logical1, [logical2, …]", "Logical", "TRUE when every value is true.", function (a) { return logicals(this, a).every(Boolean); }, { min: 1 });
def("OR", "logical1, [logical2, …]", "Logical", "TRUE when any value is true.", function (a) { return logicals(this, a).some(Boolean); }, { min: 1 });
def("XOR", "logical1, [logical2, …]", "Logical", "TRUE when an odd number of values are true.", function (a) { return logicals(this, a).filter(Boolean).length % 2 === 1; }, { min: 1 });
def("NOT", "logical", "Logical", "Flips TRUE and FALSE.", (x) => !toBool(x), L);
def("TRUE", "", "Logical", "The value TRUE.", () => true, { max: 0 });
def("FALSE", "", "Logical", "The value FALSE.", () => false, { max: 0 });
def("SWITCH", "expression, case1, value1, […], [default]", "Logical", "Compares a value against cases and returns the matching result.", function (args, ctx) {
  if (args.length < 3) return E.value("SWITCH needs a value, a case and a result");
  const v = this.deref(this.evaluate(args[0], ctx)); if (isErr(v)) return v;
  let i = 1;
  for (; i + 1 < args.length; i += 2) { const c = this.deref(this.evaluate(args[i], ctx)); if (!isErr(c) && compare(v, c) === 0 && typeRank(v) === typeRank(c)) return this.evaluate(args[i + 1], ctx); }
  return i < args.length ? this.evaluate(args[i], ctx) : E.na("No case matched");
}, { lazy: true });
def("CHOOSE", "index, choice1, [choice2, …]", "Logical", "Picks a choice by its number.", function (args, ctx) {
  const k = this.deref(this.evaluate(args[0], ctx)); if (isErr(k)) return k;
  let i; try { i = toInt(k); } catch (e) { return e; }
  if (i < 1 || i >= args.length) return E.value("CHOOSE's index is outside the choices");
  return this.evaluate(args[i], ctx);
}, { lazy: true });
def("LET", "name1, value1, [name2, value2, …], calculation", "Logical", "Names values so a formula can reuse them: =LET(x, A1*2, x+x).", function (args, ctx) {
  if (args.length < 3 || args.length % 2 === 0) return E.value("LET needs pairs of name and value, then the calculation");
  const scope = { vars: new Map(), parent: ctx.scope }, inner = { ...ctx, scope };
  for (let i = 0; i + 1 < args.length; i += 2) {
    if (args[i].k !== "name") return E.value("LET's names have to be plain words like total");
    scope.vars.set(args[i].v, this.evaluate(args[i + 1], inner));
  }
  return this.evaluate(args[args.length - 1], inner);
}, { lazy: true });

// ---- Lookup and reference
def("VLOOKUP", "search_key, range, index, [is_sorted]", "Lookup", "Finds a value in the first column and returns the value from another column in that row.", function ([key, range, idx, sorted]) {
  const k = this.deref(key), A = block(this, range), col = toInt(this.deref(idx));
  if (isErr(k)) return k;
  if (col < 1 || col > A.w) throw E.ref("The column number is outside the range");
  const firsts = A.rows.map((r) => r[0]);
  const approx = sorted === undefined ? true : toBool(this.deref(sorted));
  const i = approx ? sortedLookup(firsts, k) : lookupScan(firsts, k, 0);
  if (i < 0) throw E.na(`${toStr(k)} wasn't found`);
  return A.at(i, col - 1);
}, { min: 3, max: 4 });
def("HLOOKUP", "search_key, range, index, [is_sorted]", "Lookup", "Finds a value in the first row and returns the value from another row in that column.", function ([key, range, idx, sorted]) {
  const k = this.deref(key), A = block(this, range), row = toInt(this.deref(idx));
  if (isErr(k)) return k;
  if (row < 1 || row > A.h) throw E.ref("The row number is outside the range");
  const approx = sorted === undefined ? true : toBool(this.deref(sorted));
  const i = approx ? sortedLookup(A.rows[0], k) : lookupScan(A.rows[0], k, 0);
  if (i < 0) throw E.na(`${toStr(k)} wasn't found`);
  return A.at(row - 1, i);
}, { min: 3, max: 4 });
def("LOOKUP", "search_key, search_range, [result_range]", "Lookup", "Finds a value in a sorted row or column and returns the matching result.", function ([key, sr, rr]) {
  const k = this.deref(key), S = block(this, sr);
  const list = S.h >= S.w ? S.rows.map((r) => r[0]) : S.rows[0];
  const i = sortedLookup(list, k); if (i < 0) throw E.na();
  if (rr === undefined) return S.h >= S.w ? S.at(i, S.w - 1) : S.at(S.h - 1, i);
  const R = vec(block(this, rr)); return i < R.length ? R[i] : E.na();
}, { min: 2, max: 3 });
def("MATCH", "search_key, range, [search_type]", "Lookup", "The position of a value in a row or column. Type 0 is an exact match.", function ([key, range, type]) {
  const k = this.deref(key); if (isErr(k)) return k;
  const A = block(this, range); if (A.h > 1 && A.w > 1) throw E.na("MATCH needs a single row or column");
  const list = vec(A), t = type === undefined ? 1 : toInt(this.deref(type));
  let i;
  if (t === 0) i = lookupScan(list, k, 0);
  else if (t === 1) i = sortedLookup(list, k);
  else { i = -1; for (let j = 0; j < list.length; j++) { if (isBlank(list[j]) || !typeTie(list[j], k)) continue; if (compare(list[j], k) >= 0) i = j; else break; } }
  if (i < 0) throw E.na(`${toStr(k)} wasn't found`);
  return i + 1;
}, { min: 2, max: 3 });
def("XMATCH", "search_key, lookup_range, [match_mode], [search_mode]", "Lookup", "The position of a value. Match mode 0 exact, -1 next smaller, 1 next larger, 2 wildcards. Search mode -1 searches from the end.", function ([key, range, mm, sm]) {
  const k = this.deref(key); if (isErr(k)) return k;
  const list = vec(block(this, range)), mode = mm === undefined ? 0 : toInt(this.deref(mm)), dir = sm === undefined ? 1 : toInt(this.deref(sm));
  const order = dir < 0 ? list.map((_, i) => list.length - 1 - i) : list.map((_, i) => i);
  const i = xfind(list, order, k, mode); if (i < 0) throw E.na(`${toStr(k)} wasn't found`); return i + 1;
}, { min: 2, max: 4 });
function xfind(list, order, k, mode) {
  const sub = order.map((i) => list[i]);
  let j;
  if (mode === 2) { const re = wildcard(toStr(k)); j = sub.findIndex((x) => typeof x === "string" && re.test(x)); }
  else if (mode === 0) j = sub.findIndex((x) => !isBlank(x) && typeTie(x, k) && compare(x, k) === 0);
  else j = lookupScan(sub, k, mode === -1 ? 1 : -1);
  return j < 0 ? -1 : order[j];
}
def("XLOOKUP", "search_key, lookup_range, result_range, [if_not_found], [match_mode], [search_mode]", "Lookup", "Finds a value and returns the matching item from another range. Works left, right, up or down.", function ([key, lr, rr, nf, mm, sm]) {
  const k = this.deref(key); if (isErr(k)) return k;
  const LA = block(this, lr), RA = block(this, rr);
  const list = vec(LA), mode = mm === undefined ? 0 : toInt(this.deref(mm)), dir = sm === undefined ? 1 : toInt(this.deref(sm));
  const order = dir < 0 ? list.map((_, i) => list.length - 1 - i) : list.map((_, i) => i);
  const i = xfind(list, order, k, mode);
  if (i < 0) { if (nf !== undefined) return this.deref(nf); throw E.na(`${toStr(k)} wasn't found`); }
  if (LA.h === 1) return RA.h === 1 ? RA.at(0, i) : new Arr(RA.rows.map((r) => [r[i]]));
  return RA.w === 1 ? RA.at(i, 0) : new Arr([RA.rows[i]]);
}, { min: 3, max: 6 });
def("INDEX", "range, [row], [column]", "Lookup", "The cell at a row and column inside a range. 0 means the whole row or column.", function ([range, row, col]) {
  const r = row === undefined ? 0 : toInt(this.deref(row)), c = col === undefined ? 0 : toInt(this.deref(col));
  if (range instanceof Ref) {
    const h = range.h, w = range.w;
    let rr = r, cc = c;
    if (col === undefined && h === 1 && w > 1) { cc = r; rr = 1; }
    if (rr < 0 || cc < 0 || rr > h || cc > w) throw E.ref("INDEX's row or column is outside the range");
    if (rr === 0 && cc === 0) return range;
    if (rr === 0) return new Ref(range.sheet, range.r1, range.c1 + cc - 1, range.r2, range.c1 + cc - 1);
    if (cc === 0) return w === 1 ? new Ref(range.sheet, range.r1 + rr - 1, range.c1, range.r1 + rr - 1, range.c1) : new Ref(range.sheet, range.r1 + rr - 1, range.c1, range.r1 + rr - 1, range.c2);
    return new Ref(range.sheet, range.r1 + rr - 1, range.c1 + cc - 1, range.r1 + rr - 1, range.c1 + cc - 1);
  }
  const A = block(this, range);
  let rr = r, cc = c; if (col === undefined && A.h === 1) { cc = r; rr = 1; }
  if (rr < 0 || cc < 0 || rr > A.h || cc > A.w) throw E.ref();
  if (rr === 0) return new Arr(A.rows.map((x) => [x[(cc || 1) - 1]]));
  if (cc === 0) return A.w === 1 ? A.at(rr - 1, 0) : new Arr([A.rows[rr - 1]]);
  return A.at(rr - 1, cc - 1);
}, { min: 1, max: 3 });
def("ROW", "[cell]", "Lookup", "The row number of a cell (this cell if left out).", function ([ref], ctx) { if (ref === undefined) return ctx.r + 1; if (!(ref instanceof Ref)) throw E.value(); return ref.single ? ref.r1 + 1 : new Arr(Array.from({ length: ref.h }, (_, i) => [ref.r1 + i + 1])); }, { max: 1 });
def("COLUMN", "[cell]", "Lookup", "The column number of a cell (this cell if left out).", function ([ref], ctx) { if (ref === undefined) return ctx.c + 1; if (!(ref instanceof Ref)) throw E.value(); return ref.single ? ref.c1 + 1 : new Arr([Array.from({ length: ref.w }, (_, j) => ref.c1 + j + 1)]); }, { max: 1 });
def("ROWS", "range", "Lookup", "How many rows a range has.", function ([r]) { return r instanceof Ref ? r.h : block(this, r).h; }, { min: 1, max: 1 });
def("COLUMNS", "range", "Lookup", "How many columns a range has.", function ([r]) { return r instanceof Ref ? r.w : block(this, r).w; }, { min: 1, max: 1 });
def("ADDRESS", "row, column, [absolute], [a1], [sheet]", "Lookup", "A cell address as text, like $B$3.", (r, c, abs, a1, sh) => {
  const R = toInt(r), C = toInt(c), mode = abs === undefined ? 1 : toInt(abs);
  if (R < 1 || C < 1) throw E.value();
  const ar = mode === 1 || mode === 2, ac = mode === 1 || mode === 3;
  const text = `${ac ? "$" : ""}${colName(C - 1)}${ar ? "$" : ""}${R}`;
  return sh === undefined ? text : `${quoteSheet(toStr(sh))}!${text}`;
}, L);
def("INDIRECT", "reference_text", "Lookup", "The cells named by some text, like INDIRECT(\"B\"&A1).", function ([t], ctx) {
  const text = toStr(this.deref(t)).trim();
  const m = RE_REF.exec(text);
  if (!m || m[0].length !== text.length) throw E.ref(`“${text}” isn't a cell address`);
  return this.resolve(refParts({ m }), ctx);
}, { min: 1, max: 2 });
def("OFFSET", "cell, rows, columns, [height], [width]", "Lookup", "A range moved by some rows and columns from a starting cell.", function ([ref, r, c, h, w]) {
  if (!(ref instanceof Ref)) throw E.value("OFFSET needs a cell reference");
  const dr = toInt(this.deref(r)), dc = toInt(this.deref(c));
  const H = h === undefined ? ref.h : toInt(this.deref(h)), W = w === undefined ? ref.w : toInt(this.deref(w));
  const r1 = ref.r1 + dr, c1 = ref.c1 + dc;
  if (r1 < 0 || c1 < 0 || H < 1 || W < 1) throw E.ref("OFFSET goes off the sheet");
  return new Ref(ref.sheet, r1, c1, r1 + H - 1, c1 + W - 1);
}, { min: 3, max: 5 });
def("TRANSPOSE", "range", "Lookup", "Turns rows into columns and columns into rows.", function ([r]) { const A = block(this, r); return new Arr(Array.from({ length: A.w }, (_, j) => A.rows.map((row) => row[j]))); }, { min: 1, max: 1 });
def("FILTER", "range, condition1, [condition2 or if_empty, …]", "Lookup", "Only the rows (or columns) where the conditions are TRUE. Spills into the cells below. A last plain value is shown when nothing matches (like Excel); more ranges are more conditions (like Google Sheets).", function ([range, ...conds]) {
  const A = block(this, range);
  let ifEmpty;
  const last = conds[conds.length - 1];
  if (conds.length > 1 && !(last instanceof Ref) && !(this.deref(last) instanceof Arr)) { ifEmpty = this.deref(conds.pop()); }
  const C = conds.map((c) => block(this, this.deref(c) instanceof Arr || c instanceof Ref ? c : this.deref(c)));
  if (!C.length) throw E.value("FILTER needs a condition");
  const byRow = C[0].w === 1 && C[0].h === A.h;
  for (const X of C) { if (byRow ? !(X.w === 1 && X.h === A.h) : !(X.h === 1 && X.w === A.w)) throw E.value("FILTER's condition must be one column as tall as the range (or one row as wide)"); }
  const keep = (i) => C.every((X) => { const v = byRow ? X.at(i, 0) : X.at(0, i); if (isErr(v)) throw v; return toBool(v); });
  const rows = byRow ? A.rows.filter((_, i) => keep(i)) : A.rows.map((row) => row.filter((_, j) => keep(j)));
  if (!rows.length || !rows[0].length) { if (ifEmpty !== undefined) return ifEmpty; throw E.calc("FILTER found nothing"); }
  return new Arr(rows);
}, { min: 2 });
def("SORT", "range, [sort_column], [ascending], …", "Lookup", "The rows sorted by a column. Spills into the cells below.", function ([range, ...rest]) {
  const A = block(this, range), keys = [];
  if (!rest.length) keys.push({ i: 0, asc: true });
  for (let k = 0; k < rest.length; k += 2) {
    const idx = rest[k] === undefined ? 1 : toInt(this.deref(rest[k])), asc = rest[k + 1] === undefined ? true : (() => { const v = this.deref(rest[k + 1]); return typeof v === "number" ? v !== -1 : toBool(v); })();
    if (idx < 1 || idx > A.w) throw E.value("SORT's column is outside the range");
    keys.push({ i: idx - 1, asc });
  }
  const rows = [...A.rows].sort((x, y) => { for (const k of keys) { const c = compare(x[k.i], y[k.i]); if (c) return k.asc ? c : -c; } return 0; });
  return new Arr(rows);
}, { min: 1 });
def("SORTBY", "range, by_range1, [order1], …", "Lookup", "The rows sorted by other ranges.", function ([range, ...rest]) {
  const A = block(this, range), keys = [];
  for (let k = 0; k < rest.length; k += 2) { const B = vec(block(this, rest[k])); if (B.length !== A.h) throw E.value(); keys.push({ B, asc: rest[k + 1] === undefined ? true : toNum(this.deref(rest[k + 1])) !== -1 }); }
  const idx = A.rows.map((_, i) => i).sort((x, y) => { for (const k of keys) { const c = compare(k.B[x], k.B[y]); if (c) return k.asc ? c : -c; } return 0; });
  return new Arr(idx.map((i) => A.rows[i]));
}, { min: 2 });
def("UNIQUE", "range, [by_column], [exactly_once]", "Lookup", "The range with repeated rows taken out. Spills into the cells below.", function ([range, byCol, once]) {
  let A = block(this, range); const cols = byCol !== undefined && toBool(this.deref(byCol)), one = once !== undefined && toBool(this.deref(once));
  if (cols) A = new Arr(Array.from({ length: A.w }, (_, j) => A.rows.map((r) => r[j])));
  const key = (row) => JSON.stringify(row.map((v) => (typeof v === "string" ? v.toLowerCase() : isErr(v) ? v.code : v)));
  const count = new Map(); for (const r of A.rows) count.set(key(r), (count.get(key(r)) || 0) + 1);
  const seen = new Set(), out = [];
  for (const r of A.rows) { const k = key(r); if (seen.has(k)) continue; seen.add(k); if (!one || count.get(k) === 1) out.push(r); }
  if (!out.length) throw E.calc("Nothing is unique");
  return cols ? new Arr(Array.from({ length: out[0].length }, (_, i) => out.map((r) => r[i]))) : new Arr(out);
}, { min: 1, max: 3 });
def("VSTACK", "range1, [range2, …]", "Lookup", "Ranges stacked on top of each other.", function (a) { const B = a.map((x) => block(this, x)), w = Math.max(...B.map((x) => x.w)); return new Arr(B.flatMap((X) => X.rows.map((r) => [...r, ...Array(w - r.length).fill(E.na())]))); }, { min: 1 });
def("HSTACK", "range1, [range2, …]", "Lookup", "Ranges side by side.", function (a) { const B = a.map((x) => block(this, x)), h = Math.max(...B.map((x) => x.h)); return new Arr(Array.from({ length: h }, (_, i) => B.flatMap((X) => (i < X.h ? X.rows[i] : Array(X.w).fill(E.na()))))); }, { min: 1 });
def("TOCOL", "range, [ignore]", "Lookup", "The values of a range in one column. Ignore 1 skips blanks, 2 errors, 3 both.", function ([r, ig]) { const g = ig === undefined ? 0 : toInt(this.deref(ig)); const v = cells(block(this, r)).filter((x) => !((g & 1) && isBlank(x)) && !((g & 2) && isErr(x))); return new Arr(v.map((x) => [x])); }, { min: 1 });
def("TOROW", "range, [ignore]", "Lookup", "The values of a range in one row.", function ([r, ig]) { const g = ig === undefined ? 0 : toInt(this.deref(ig)); return new Arr([cells(block(this, r)).filter((x) => !((g & 1) && isBlank(x)) && !((g & 2) && isErr(x)))]); }, { min: 1 });
def("CHOOSECOLS", "range, column1, [column2, …]", "Lookup", "Only the columns you list (negative counts from the right).", function ([r, ...cs]) { const A = block(this, r); const idx = cs.map((c) => { const i = toInt(this.deref(c)); const j = i < 0 ? A.w + i : i - 1; if (j < 0 || j >= A.w) throw E.value(); return j; }); return new Arr(A.rows.map((row) => idx.map((j) => row[j]))); }, { min: 2 });
def("CHOOSEROWS", "range, row1, [row2, …]", "Lookup", "Only the rows you list (negative counts from the bottom).", function ([r, ...rs]) { const A = block(this, r); return new Arr(rs.map((x) => { const i = toInt(this.deref(x)); const j = i < 0 ? A.h + i : i - 1; if (j < 0 || j >= A.h) throw E.value(); return A.rows[j]; })); }, { min: 2 });
def("TAKE", "range, rows, [columns]", "Lookup", "The first rows and columns (negative takes from the end).", function ([r, rs, cs]) { let A = block(this, r).rows; const n = rs === undefined ? A.length : toInt(this.deref(rs)); A = n >= 0 ? A.slice(0, n) : A.slice(n); if (cs !== undefined) { const m = toInt(this.deref(cs)); A = A.map((row) => (m >= 0 ? row.slice(0, m) : row.slice(m))); } if (!A.length || !A[0].length) throw E.calc(); return new Arr(A); }, { min: 2 });
def("DROP", "range, rows, [columns]", "Lookup", "The range without its first rows and columns (negative drops from the end).", function ([r, rs, cs]) { let A = block(this, r).rows; const n = toInt(this.deref(rs)); A = n >= 0 ? A.slice(n) : A.slice(0, n); if (cs !== undefined) { const m = toInt(this.deref(cs)); A = A.map((row) => (m >= 0 ? row.slice(m) : row.slice(0, m))); } if (!A.length || !A[0].length) throw E.calc(); return new Arr(A); }, { min: 2 });
def("ARRAYFORMULA", "formula", "Lookup", "Shows a calculation over whole ranges. Ranges already work this way here, so it just passes the result on.", function ([x]) { return x; }, { min: 1, max: 1 });

// ---- Text
def("CONCATENATE", "text1, [text2, …]", "Text", "Joins texts together.", function (a) { return a.map((x) => { const v = this.deref(x); if (v instanceof Arr) return cells(v).map(toStr).join(""); return toStr(v); }).join(""); });
def("CONCAT", "text1, [text2, …]", "Text", "Joins texts and whole ranges together.", function (a) { return flat(this, a, true).map((x) => toStr(x.v)).join(""); });
def("TEXTJOIN", "delimiter, ignore_empty, text1, [text2, …]", "Text", "Joins texts with a separator, like a comma, optionally skipping blanks.", function ([d, ig, ...rest]) {
  const sep = toStr(this.deref(d)), skip = toBool(this.deref(ig));
  return flat(this, rest, true).map((x) => toStr(x.v)).filter((s) => !(skip && s === "")).join(sep);
}, { min: 3 });
def("JOIN", "delimiter, value_or_range1, …", "Text", "Joins values with a separator.", function ([d, ...rest]) { const sep = toStr(this.deref(d)); return flat(this, rest, true).filter((x) => !isBlank(x.v)).map((x) => toStr(x.v)).join(sep); }, { min: 2 });
def("SPLIT", "text, delimiter, [split_by_each], [remove_empty]", "Text", "Splits text into cells across the row.", (t, d, each, rm) => {
  const s = toStr(t), sep = toStr(d), byEach = each === undefined ? true : toBool(each), drop = rm === undefined ? true : toBool(rm);
  let parts = byEach && sep.length > 1 ? s.split(new RegExp(`[${sep.replace(/[\]\\^-]/g, "\\$&")}]`)) : s.split(sep);
  if (drop) parts = parts.filter((p) => p !== "");
  if (!parts.length) throw E.value();
  return new Arr([parts.map((p) => { const v = parseInput(p).v; return typeof v === "number" ? v : p; })]);
}, L);
def("TEXTSPLIT", "text, column_delimiter, [row_delimiter], [ignore_empty]", "Text", "Splits text into columns and rows.", (t, cd, rd, ig) => {
  const s = toStr(t), cs = toStr(cd), rs = rd === undefined ? null : toStr(rd), skip = ig !== undefined && toBool(ig);
  let rows = rs ? s.split(rs) : [s]; if (skip) rows = rows.filter((r) => r !== "");
  const grid = rows.map((r) => { let p = cs ? r.split(cs) : [r]; if (skip) p = p.filter((x) => x !== ""); return p; });
  const w = Math.max(...grid.map((r) => r.length));
  return new Arr(grid.map((r) => [...r, ...Array(w - r.length).fill(E.na())]));
}, L);
def("LEFT", "text, [characters]", "Text", "The first characters of a text.", (t, n) => { const k = n === undefined ? 1 : toInt(n); if (k < 0) throw E.value(); return toStr(t).slice(0, k); }, L);
def("RIGHT", "text, [characters]", "Text", "The last characters of a text.", (t, n) => { const k = n === undefined ? 1 : toInt(n); if (k < 0) throw E.value(); const s = toStr(t); return k ? s.slice(-k) : ""; }, L);
def("MID", "text, start, characters", "Text", "Characters from the middle of a text.", (t, s, n) => { const a = toInt(s), k = toInt(n); if (a < 1 || k < 0) throw E.value(); return toStr(t).substr(a - 1, k); }, L);
def("LEN", "text", "Text", "How many characters a text has.", (t) => toStr(t).length, L);
def("LOWER", "text", "Text", "The text in lowercase.", (t) => toStr(t).toLowerCase(), L);
def("UPPER", "text", "Text", "The text in uppercase.", (t) => toStr(t).toUpperCase(), L);
def("PROPER", "text", "Text", "Each Word Starting With A Capital.", (t) => toStr(t).toLowerCase().replace(/(^|[^a-zÀ-￿'])([a-zÀ-￿])/g, (m, a, b) => a + b.toUpperCase()), L);
def("TRIM", "text", "Text", "Removes extra spaces.", (t) => toStr(t).replace(/ +/g, " ").trim(), L);
def("CLEAN", "text", "Text", "Removes characters that can't be printed.", (t) => toStr(t).replace(/[\x00-\x1f]/g, ""), L);
def("SUBSTITUTE", "text, search_for, replace_with, [occurrence]", "Text", "Replaces some text with other text.", (t, a, b, n) => {
  const s = toStr(t), find = toStr(a), rep = toStr(b);
  if (!find) return s;
  if (n === undefined) return s.split(find).join(rep);
  const k = toInt(n); if (k < 1) throw E.value();
  let idx = -1; for (let i = 0; i < k; i++) { idx = s.indexOf(find, idx + 1); if (idx < 0) return s; }
  return s.slice(0, idx) + rep + s.slice(idx + find.length);
}, L);
def("REPLACE", "text, position, length, new_text", "Text", "Replaces characters at a position.", (t, p, n, r) => { const s = toStr(t), a = toInt(p), k = toInt(n); if (a < 1 || k < 0) throw E.value(); return s.slice(0, a - 1) + toStr(r) + s.slice(a - 1 + k); }, L);
def("FIND", "search_for, text, [start]", "Text", "Where some text starts inside another (matches case).", (f, t, s) => { const st = s === undefined ? 1 : toInt(s); const i = toStr(t).indexOf(toStr(f), st - 1); if (st < 1 || i < 0) throw E.value(`“${toStr(f)}” wasn't found`); return i + 1; }, L);
def("SEARCH", "search_for, text, [start]", "Text", "Where some text starts inside another (any case, wildcards allowed).", (f, t, s) => {
  const st = s === undefined ? 1 : toInt(s), hay = toStr(t), needle = toStr(f);
  if (st < 1 || st > hay.length + 1) throw E.value();
  const re = new RegExp(wildcard(needle).source.slice(1, -1), "is");
  const m = re.exec(hay.slice(st - 1)); if (!m) throw E.value(`“${needle}” wasn't found`);
  return st + m.index;
}, L);
def("REPT", "text, times", "Text", "The text repeated.", (t, n) => { const k = toInt(n); if (k < 0) throw E.value(); return toStr(t).repeat(k); }, L);
def("EXACT", "text1, text2", "Text", "TRUE when two texts are exactly the same, case included.", (a, b) => toStr(a) === toStr(b), L);
def("CHAR", "number", "Text", "The character for a code number.", (n) => { const k = toInt(n); if (k < 1 || k > 0x10ffff) throw E.value(); return String.fromCodePoint(k); }, L);
FN.UNICHAR = FN.CHAR;
def("CODE", "text", "Text", "The code number of the first character.", (t) => { const s = toStr(t); if (!s) throw E.value(); return s.codePointAt(0); }, L);
FN.UNICODE = FN.CODE;
def("TEXT", "value, format", "Text", "A number shown with a format, like TEXT(A1, \"$#,##0.00\") or TEXT(A1, \"mmm d\").", (v, f) => format(typeof v === "string" ? (typeof parseInput(v).v === "number" ? parseInput(v).v : v) : v, toStr(f)).text, L);
def("VALUE", "text", "Text", "Text that looks like a number, as a number.", (t) => (typeof t === "number" ? t : toNum(toStr(t))), L);
def("NUMBERVALUE", "text, [decimal_separator], [group_separator]", "Text", "Text as a number, with your own separators.", (t, d, g) => { let s = toStr(t).trim(); const dec = d === undefined ? "." : toStr(d), grp = g === undefined ? "," : toStr(g); s = s.split(grp).join("").split(dec).join("."); const pct = (s.match(/%/g) || []).length; s = s.replace(/%/g, ""); if (s === "") return 0; const n = Number(s); if (isNaN(n)) throw E.value(); return n / Math.pow(100, pct); }, L);
def("T", "value", "Text", "The value if it's text, otherwise empty.", (v) => (typeof v === "string" ? v : ""), L);
def("N", "value", "Text", "The value as a number (text becomes 0).", (v) => (typeof v === "number" ? v : typeof v === "boolean" ? (v ? 1 : 0) : 0), L);
def("FIXED", "number, [decimals], [no_commas]", "Text", "A number as text with set decimals.", (n, d, nc) => format(roundTo(toNum(n), d === undefined ? 2 : toInt(d)), `${nc !== undefined && toBool(nc) ? "0" : "#,##0"}${(d === undefined ? 2 : toInt(d)) > 0 ? "." + "0".repeat(toInt(d ?? 2)) : ""}`).text, L);
def("DOLLAR", "number, [decimals]", "Text", "A number as money text.", (n, d) => { const k = d === undefined ? 2 : toInt(d); return format(roundTo(toNum(n), k), `$#,##0${k > 0 ? "." + "0".repeat(k) : ""};($#,##0${k > 0 ? "." + "0".repeat(k) : ""})`).text; }, L);
def("TEXTBEFORE", "text, delimiter, [instance]", "Text", "The text before a delimiter.", (t, d, n) => { const s = toStr(t), sep = toStr(d), k = n === undefined ? 1 : toInt(n); const parts = s.split(sep); if (parts.length - 1 < Math.abs(k)) throw E.na(); return k > 0 ? parts.slice(0, k).join(sep) : parts.slice(0, parts.length + k).join(sep); }, L);
def("TEXTAFTER", "text, delimiter, [instance]", "Text", "The text after a delimiter.", (t, d, n) => { const s = toStr(t), sep = toStr(d), k = n === undefined ? 1 : toInt(n); const parts = s.split(sep); if (parts.length - 1 < Math.abs(k)) throw E.na(); return k > 0 ? parts.slice(k).join(sep) : parts.slice(parts.length + k).join(sep); }, L);
def("REGEXMATCH", "text, regular_expression", "Text", "TRUE when the text matches a regular expression.", (t, r) => { try { return new RegExp(toStr(r)).test(toStr(t)); } catch (e) { throw E.value("That regular expression isn't valid"); } }, L);
def("REGEXEXTRACT", "text, regular_expression", "Text", "The part of the text that matches (or its first group).", (t, r) => { let m; try { m = new RegExp(toStr(r)).exec(toStr(t)); } catch (e) { throw E.value("That regular expression isn't valid"); } if (!m) throw E.na("No match"); return m[1] !== undefined ? m[1] : m[0]; }, L);
def("REGEXREPLACE", "text, regular_expression, replacement", "Text", "Replaces every match of a regular expression.", (t, r, w) => { try { return toStr(t).replace(new RegExp(toStr(r), "g"), toStr(w)); } catch (e) { throw E.value("That regular expression isn't valid"); } }, L);

// ---- Date and time
def("TODAY", "", "Date", "Today's date. Updates on every edit.", function () { return todaySerial(this.now); }, { max: 0 });
def("NOW", "", "Date", "The date and time now.", function () { return nowSerial(this.now); }, { max: 0 });
def("DATE", "year, month, day", "Date", "A date from its year, month and day. Months and days past the end roll over.", (y, m, d) => {
  let Y = toInt(y); const M = toInt(m), Dd = toInt(d); if (Y < 1900) Y += 1900;
  const dt = new Date(Date.UTC(Y, M - 1, 1)); const s = Math.round((dt.getTime() - EPOCH) / DAY) + Dd - 1;
  if (s < 0) throw E.num(); return s;
}, L);
def("TIME", "hour, minute, second", "Date", "A time of day as a fraction of a day.", (h, m, s) => { const v = (toInt(h) * 3600 + toInt(m) * 60 + toInt(s)) / 86400; if (v < 0) throw E.num(); return v - Math.floor(v); }, L);
const dpart = (f) => (v) => { const n = toNum(v); if (n < 0) throw E.num(); return f(fromSerial(n)); };
def("YEAR", "date", "Date", "The year of a date.", dpart((t) => t.y), L);
def("MONTH", "date", "Date", "The month of a date (1-12).", dpart((t) => t.m), L);
def("DAY", "date", "Date", "The day of the month.", dpart((t) => t.d), L);
def("HOUR", "time", "Date", "The hour (0-23).", dpart((t) => t.hh), L);
def("MINUTE", "time", "Date", "The minute.", dpart((t) => t.mi), L);
def("SECOND", "time", "Date", "The second.", dpart((t) => t.ss), L);
def("WEEKDAY", "date, [type]", "Date", "The day of the week: 1 Sunday to 7 Saturday (type 2: 1 Monday, type 3: 0 Monday).", (d, t) => {
  const wd = fromSerial(toNum(d)).wd, k = t === undefined ? 1 : toInt(t);
  if (k === 1 || k === 17) return wd + 1; if (k === 2 || k === 11) return ((wd + 6) % 7) + 1; if (k === 3) return (wd + 6) % 7;
  if (k >= 12 && k <= 16) return ((wd - (k - 10) + 7) % 7) + 1; throw E.num();
}, L);
def("WEEKNUM", "date, [type]", "Date", "The week of the year (weeks start on Sunday, or Monday with type 2).", (d, t) => {
  const s = toNum(d), x = fromSerial(s), k = t === undefined ? 1 : toInt(t), start = k === 2 || k === 11 ? 1 : 0;
  if (k === 21) return FN.ISOWEEKNUM.fn(s);
  const jan1 = dateSerial(x.y, 1, 1), wd = fromSerial(jan1).wd; const off = (wd - start + 7) % 7;
  return Math.floor((s - jan1 + off) / 7) + 1;
}, L);
def("ISOWEEKNUM", "date", "Date", "The ISO week number.", (d) => { const s = Math.floor(toNum(d)); const wd = (fromSerial(s).wd + 6) % 7; const th = s - wd + 3; const y = fromSerial(th).y; return Math.floor((th - dateSerial(y, 1, 1)) / 7) + 1; }, L);
def("EDATE", "start_date, months", "Date", "The same day some months later (or earlier).", (d, m) => { const t = fromSerial(toNum(d)); const mm = t.m - 1 + toInt(m); const y = t.y + Math.floor(mm / 12), mo = ((mm % 12) + 12) % 12; const last = new Date(Date.UTC(y, mo + 1, 0)).getUTCDate(); return dateSerial(y, mo + 1, Math.min(t.d, last)); }, L);
def("EOMONTH", "start_date, months", "Date", "The last day of the month, some months later.", (d, m) => { const t = fromSerial(toNum(d)); const mm = t.m - 1 + toInt(m); const y = t.y + Math.floor(mm / 12), mo = ((mm % 12) + 12) % 12; return dateSerial(y, mo + 1, new Date(Date.UTC(y, mo + 1, 0)).getUTCDate()); }, L);
def("DAYS", "end_date, start_date", "Date", "Days between two dates.", (e, s) => Math.floor(toNum(dateLike(e))) - Math.floor(toNum(dateLike(s))), L);
def("DAYS360", "start_date, end_date, [european]", "Date", "Days between dates on a 360-day year.", (s, e, eu) => { const a = fromSerial(toNum(s)), b = fromSerial(toNum(e)); let d1 = a.d, d2 = b.d; if (eu !== undefined && toBool(eu)) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; } else { if (d1 === 31) d1 = 30; if (d2 === 31 && d1 >= 30) d2 = 30; } return (b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1); }, L);
def("DATEDIF", "start_date, end_date, unit", "Date", "The time between dates in \"Y\" years, \"M\" months, \"D\" days, \"MD\", \"YM\" or \"YD\".", (s, e, u) => {
  const a = toNum(dateLike(s)), b = toNum(dateLike(e)); if (b < a) throw E.num("The end is before the start");
  const x = fromSerial(a), y = fromSerial(b), unit = toStr(u).toUpperCase();
  let months = (y.y - x.y) * 12 + (y.m - x.m); if (y.d < x.d) months--;
  switch (unit) {
    case "D": return Math.floor(b) - Math.floor(a);
    case "M": return months;
    case "Y": return Math.floor(months / 12);
    case "YM": return months % 12;
    case "MD": { let d = y.d - x.d; if (d < 0) d += new Date(Date.UTC(y.y, y.m - 1, 0)).getUTCDate(); return d; }
    case "YD": { let start = dateSerial(y.y, x.m, x.d); if (start > b) start = dateSerial(y.y - 1, x.m, x.d); return Math.floor(b) - start; }
    default: throw E.num("Unit is one of Y, M, D, MD, YM, YD");
  }
}, L);
function dateLike(v) { if (typeof v === "string") { const p = parseInput(v); if (typeof p.v === "number") return p.v; throw E.value(`“${v}” isn't a date`); } return v; }
def("DATEVALUE", "date_text", "Date", "Text that looks like a date, as a date.", (t) => { const p = parseInput(toStr(t)); if (typeof p.v !== "number" || !isDateFormat(p.fmt || "")) throw E.value(`“${toStr(t)}” isn't a date`); return Math.floor(p.v); }, L);
def("TIMEVALUE", "time_text", "Date", "Text that looks like a time, as a time.", (t) => { const p = parseInput(toStr(t)); if (typeof p.v !== "number" || !p.fmt) throw E.value(`“${toStr(t)}” isn't a time`); return p.v - Math.floor(p.v); }, L);
function holidaysOf(book, h) { return h === undefined ? new Set() : new Set(cells(block(book, h)).filter((v) => typeof v === "number").map(Math.floor)); }
const weekend = (s) => { const wd = fromSerial(s).wd; return wd === 0 || wd === 6; };
def("NETWORKDAYS", "start_date, end_date, [holidays]", "Date", "Working days (Monday to Friday) between two dates, both counted.", function ([s, e, h]) {
  let a = Math.floor(toNum(dateLike(this.deref(s)))), b = Math.floor(toNum(dateLike(this.deref(e)))), sign = 1; const hol = holidaysOf(this, h);
  if (b < a) { [a, b] = [b, a]; sign = -1; }
  let n = 0; for (let d = a; d <= b; d++) if (!weekend(d) && !hol.has(d)) n++;
  return sign * n;
}, { min: 2, max: 3 });
def("WORKDAY", "start_date, days, [holidays]", "Date", "The date some working days later.", function ([s, n, h]) {
  let d = Math.floor(toNum(dateLike(this.deref(s)))); let k = toInt(this.deref(n)); const hol = holidaysOf(this, h), step = k < 0 ? -1 : 1;
  while (k !== 0) { d += step; if (!weekend(d) && !hol.has(d)) k -= step; }
  return d;
}, { min: 2, max: 3 });
def("YEARFRAC", "start_date, end_date, [basis]", "Date", "The fraction of a year between dates.", (s, e, b) => { const a = toNum(s), z = toNum(e), basis = b === undefined ? 0 : toInt(b); if (basis === 0) return FN.DAYS360.fn(Math.min(a, z), Math.max(a, z)) / 360; return Math.abs(z - a) / (basis === 2 ? 360 : 365); }, L);

// ---- Information
def("ISBLANK", "value", "Info", "TRUE when the cell is empty.", function ([v]) { const x = this.deref(v); return x instanceof Arr ? new Arr(x.rows.map((r) => r.map(isBlank))) : isBlank(x); }, { min: 1, max: 1 });
const infoLift = (f) => function ([v]) { const x = this.deref(v); return x instanceof Arr ? new Arr(x.rows.map((r) => r.map(f))) : f(x); };
def("ISNUMBER", "value", "Info", "TRUE for a number.", infoLift((x) => typeof x === "number"), { min: 1, max: 1 });
def("ISTEXT", "value", "Info", "TRUE for text.", infoLift((x) => typeof x === "string"), { min: 1, max: 1 });
def("ISNONTEXT", "value", "Info", "TRUE for anything that isn't text.", infoLift((x) => typeof x !== "string"), { min: 1, max: 1 });
def("ISLOGICAL", "value", "Info", "TRUE for TRUE or FALSE.", infoLift((x) => typeof x === "boolean"), { min: 1, max: 1 });
def("ISERROR", "value", "Info", "TRUE for any error.", infoLift(isErr), { min: 1, max: 1 });
def("ISERR", "value", "Info", "TRUE for any error except #N/A.", infoLift((x) => isErr(x) && x.code !== "#N/A"), { min: 1, max: 1 });
def("ISNA", "value", "Info", "TRUE for #N/A.", infoLift((x) => isErr(x) && x.code === "#N/A"), { min: 1, max: 1 });
def("ISEVEN", "value", "Info", "TRUE for an even number.", (x) => Math.floor(Math.abs(toNum(x))) % 2 === 0, L);
def("ISODD", "value", "Info", "TRUE for an odd number.", (x) => Math.floor(Math.abs(toNum(x))) % 2 === 1, L);
def("ISREF", "value", "Info", "TRUE for a cell reference.", function ([v]) { return v instanceof Ref; }, { min: 1, max: 1 });
def("ISFORMULA", "cell", "Info", "TRUE when the cell holds a formula.", function ([v]) { if (!(v instanceof Ref)) throw E.value(); const raw = this.raw(v.sheet, v.r1, v.c1); return typeof raw === "string" && raw.startsWith("="); }, { min: 1, max: 1 });
def("FORMULATEXT", "cell", "Info", "The formula in a cell, as text.", function ([v]) { if (!(v instanceof Ref)) throw E.value(); const raw = this.raw(v.sheet, v.r1, v.c1); if (!(typeof raw === "string" && raw.startsWith("="))) throw E.na("That cell has no formula"); return raw; }, { min: 1, max: 1 });
def("NA", "", "Info", "The #N/A error, for “no value here”.", () => E.na("Marked as not available"), { max: 0 });
def("ERROR.TYPE", "value", "Info", "A number for the kind of error.", function ([v]) { const x = this.deref(v); if (!isErr(x)) throw E.na(); return ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"].indexOf(x.code) + 1 || 8; }, { min: 1, max: 1 });
def("TYPE", "value", "Info", "1 number, 2 text, 4 TRUE/FALSE, 16 error, 64 a range.", function ([v]) { const x = this.deref(v); return x instanceof Arr ? 64 : typeof x === "number" || isBlank(x) ? 1 : typeof x === "string" ? 2 : typeof x === "boolean" ? 4 : 16; }, { min: 1, max: 1 });
def("SHEET", "[value]", "Info", "The number of a sheet.", function ([v], ctx) { if (v === undefined) return ctx.si + 1; if (v instanceof Ref) return v.sheet + 1; const i = this.sheetIndex(toStr(v)); if (i < 0) throw E.na(); return i + 1; }, { max: 1 });
def("SHEETS", "", "Info", "How many sheets the workbook has.", function () { return this.doc.sheets.length; }, { max: 1 });

// ---- Finance (payments are negative when money goes out, like both apps)
const fvOf = (r, n, p, pv, t) => (r === 0 ? -(pv + p * n) : -(pv * Math.pow(1 + r, n) + p * (1 + r * t) * (Math.pow(1 + r, n) - 1) / r));
def("PMT", "rate, periods, present_value, [future_value], [end_or_beginning]", "Finance", "The payment each period for a loan. For a monthly loan use the yearly rate / 12.", (r, n, pv, fv, t) => {
  const R = toNum(r), N = toNum(n), P = toNum(pv), F = fv === undefined ? 0 : toNum(fv), T = t === undefined ? 0 : toNum(t) ? 1 : 0;
  if (N === 0) throw E.num(); if (R === 0) return -(P + F) / N;
  return -(R * (P * Math.pow(1 + R, N) + F)) / ((1 + R * T) * (Math.pow(1 + R, N) - 1));
}, L);
def("IPMT", "rate, period, periods, present_value, [future_value], [type]", "Finance", "The interest part of one payment.", (r, per, n, pv, fv, t) => {
  const R = toNum(r), K = toNum(per), N = toNum(n), P = toNum(pv), F = fv === undefined ? 0 : toNum(fv), T = t === undefined ? 0 : toNum(t) ? 1 : 0;
  if (K < 1 || K > N) throw E.num("The period is outside the loan");
  const pmt = FN.PMT.fn(R, N, P, F, T); if (K === 1 && T === 1) return 0;
  const ipmt = fvOf(R, K - 1, pmt, P, T) * R;
  return T === 1 ? ipmt / (1 + R) : ipmt;
}, L);
def("PPMT", "rate, period, periods, present_value, [future_value], [type]", "Finance", "The principal part of one payment.", (r, per, n, pv, fv, t) => FN.PMT.fn(r, n, pv, fv, t) - FN.IPMT.fn(r, per, n, pv, fv, t), L);
def("FV", "rate, periods, payment, [present_value], [type]", "Finance", "What savings or a loan will be worth after some periods.", (r, n, p, pv, t) => fvOf(toNum(r), toNum(n), toNum(p), pv === undefined ? 0 : toNum(pv), t === undefined ? 0 : toNum(t) ? 1 : 0), L);
def("PV", "rate, periods, payment, [future_value], [type]", "Finance", "What a series of payments is worth today.", (r, n, p, fv, t) => {
  const R = toNum(r), N = toNum(n), P = toNum(p), F = fv === undefined ? 0 : toNum(fv), T = t === undefined ? 0 : toNum(t) ? 1 : 0;
  if (R === 0) return -(F + P * N);
  return -(F + P * (1 + R * T) * (Math.pow(1 + R, N) - 1) / R) / Math.pow(1 + R, N);
}, L);
def("NPER", "rate, payment, present_value, [future_value], [type]", "Finance", "How many periods to pay off a loan or reach a goal.", (r, p, pv, fv, t) => {
  const R = toNum(r), P = toNum(p), V = toNum(pv), F = fv === undefined ? 0 : toNum(fv), T = t === undefined ? 0 : toNum(t) ? 1 : 0;
  if (R === 0) { if (P === 0) throw E.num(); return -(V + F) / P; }
  const num = P * (1 + R * T) - F * R, den = V * R + P * (1 + R * T);
  if (num / den <= 0) throw E.num(); return Math.log(num / den) / Math.log(1 + R);
}, L);
def("RATE", "periods, payment, present_value, [future_value], [type], [guess]", "Finance", "The interest rate per period.", (n, p, pv, fv, t, g) => {
  const N = toNum(n), P = toNum(p), V = toNum(pv), F = fv === undefined ? 0 : toNum(fv), T = t === undefined ? 0 : toNum(t) ? 1 : 0;
  let r = g === undefined ? 0.1 : toNum(g);
  const f = (x) => (Math.abs(x) < 1e-10 ? V + P * N + F : V * Math.pow(1 + x, N) + P * (1 + x * T) * (Math.pow(1 + x, N) - 1) / x + F);
  for (let i = 0; i < 100; i++) { const y = f(r), d = (f(r + 1e-6) - y) / 1e-6; if (!d) break; const nr = r - y / d; if (Math.abs(nr - r) < 1e-10) return nr; r = nr; }
  throw E.num("RATE didn't settle; try a guess");
}, L);
def("NPV", "rate, value1, [value2, …]", "Finance", "Net present value of future cash flows.", function ([r, ...rest]) { const R = toNum(this.deref(r)); return nums(this, rest).reduce((t, v, i) => t + v / Math.pow(1 + R, i + 1), 0); }, { min: 2 });
def("IRR", "cash_flows, [guess]", "Finance", "The rate that makes the cash flows' value zero.", function ([cf, g]) {
  const v = nums(this, [cf]); let r = g === undefined ? 0.1 : toNum(this.deref(g));
  if (!v.some((x) => x > 0) || !v.some((x) => x < 0)) throw E.num("Cash flows need money going out and coming in");
  const f = (x) => v.reduce((t, c, i) => t + c / Math.pow(1 + x, i), 0);
  for (let i = 0; i < 200; i++) { const y = f(r), d = (f(r + 1e-7) - y) / 1e-7; if (!d) break; const nr = r - y / d; if (Math.abs(nr - r) < 1e-12) return nr; r = nr; }
  throw E.num("IRR didn't settle; try a guess");
}, { min: 1, max: 2 });
def("SLN", "cost, salvage, life", "Finance", "Straight-line depreciation per period.", (c, s, l) => { const L2 = toNum(l); if (!L2) throw E.div(); return (toNum(c) - toNum(s)) / L2; }, L);
def("EFFECT", "nominal_rate, periods_per_year", "Finance", "The effective yearly rate.", (r, n) => { const N = toInt(n); if (N < 1) throw E.num(); return Math.pow(1 + toNum(r) / N, N) - 1; }, L);
def("NOMINAL", "effective_rate, periods_per_year", "Finance", "The nominal yearly rate.", (r, n) => { const N = toInt(n); if (N < 1) throw E.num(); return N * (Math.pow(1 + toNum(r), 1 / N) - 1); }, L);
def("RRI", "periods, present_value, future_value", "Finance", "The growth rate per period to get from one value to another.", (n, p, f) => { const N = toNum(n), P = toNum(p); if (N <= 0 || P === 0) throw E.num(); return Math.pow(toNum(f) / P, 1 / N) - 1; }, L);

// ---------------------------------------------------------------- export
const api = { FErr, Arr, Ref, Book, isErr, ERRS, E, colName, colIndex, addr, parseAddr, tokenize, parse, refParts, refText, shiftFormula, structFormula, renameSheetIn,
  dropSheetIn, refsIn, parseInput, format, isDateFormat, toNum, toStr, compare, criteria, dateSerial, fromSerial, numText, quoteSheet, FUNCTIONS: CATALOG, FN, MAX_ROWS, MAX_COLS };
if (typeof module !== "undefined" && module.exports) module.exports = api; else root.Formula = api;
})(typeof window !== "undefined" ? window : globalThis);
