// Drives the Sheets page in Node with a small stand-in for the browser (DOM, canvas, clipboard), against an isolated
// dashboard started by serve_isolated.py. The server has no browser, so this is how the page is checked end to end.
//   python3 tests/serve_isolated.py 8891 /tmp/x &   node tests/ui_test.js http://127.0.0.1:8891
"use strict";
const fs = require("fs");
const BASE = process.argv[2];
const STATIC = (process.env.CAPTURE_DIR || process.env.HOME + "/capture-app") + "/capture_app/static";
const errors = [];
process.on("unhandledRejection", (e) => errors.push("unhandled: " + (e && e.stack || e)));
process.on("uncaughtException", (e) => errors.push("uncaught: " + (e && e.stack || e)));

// ---------------------------------------------------------------- a small DOM
class Node_ { constructor() { this.parentNode = null; } get textContent() { return ""; } }
class Text_ extends Node_ { constructor(t) { super(); this.t = t; } get textContent() { return this.t; } }
class El extends Node_ {
  constructor(tag) {
    super(); this.tagName = tag.toUpperCase(); this.children = []; this.attrs = {}; this.listeners = {}; this._cls = new Set(); this.dataset = {};
    const self = this;
    this.style = new Proxy({ cssText: "" }, { set(o, k, v) { o[k] = v; if (k === "cssText" && v === "") for (const key of Object.keys(o)) if (key !== "cssText") delete o[key]; return true; } });
    this.classList = { add: (c) => self._cls.add(c), remove: (c) => self._cls.delete(c), toggle: (c, on) => { (on === undefined ? !self._cls.has(c) : on) ? self._cls.add(c) : self._cls.delete(c); }, contains: (c) => self._cls.has(c) };
    this.value = ""; this.checked = false; this.disabled = false; this.hidden = false; this.selectionStart = 0; this.selectionEnd = 0; this.scrollLeft = 0; this.scrollTop = 0; this.open = false;
  }
  get className() { return [...this._cls].join(" "); } set className(v) { this._cls = new Set(String(v).split(/\s+/).filter(Boolean)); }
  get textContent() { return this.children.map((c) => c.textContent).join(""); } set textContent(v) { this.children = []; if (v !== "") this.append(String(v)); }
  get firstChild() { return this.children[0] || null; } get lastChild() { return this.children[this.children.length - 1] || null; }
  get offsetWidth() { return 200; } get offsetHeight() { return 100; } get scrollHeight() { return 21; }
  get clientWidth() { return this._cw || 1200; } get clientHeight() { return this._ch || 700; }
  append(...kids) { for (const k of kids) { const n = typeof k === "string" ? new Text_(k) : k; if (n.parentNode) n.parentNode.children = n.parentNode.children.filter((x) => x !== n); n.parentNode = this; this.children.push(n); } }
  replaceChildren(...kids) { for (const c of this.children) c.parentNode = null; this.children = []; this.append(...kids); }
  remove() { if (this.parentNode) { this.parentNode.children = this.parentNode.children.filter((x) => x !== this); this.parentNode = null; } }
  setAttribute(k, v) { this.attrs[k] = String(v); if (k.startsWith("data-")) this.dataset[k.slice(5)] = String(v); if (k === "class") this.className = v; }
  getAttribute(k) { return k in this.attrs ? this.attrs[k] : null; }
  addEventListener(t, f) { (this.listeners[t] = this.listeners[t] || []).push(f); }
  removeEventListener(t, f) { this.listeners[t] = (this.listeners[t] || []).filter((x) => x !== f); }
  dispatchEvent(ev) { ev.target = ev.target || this; ev.currentTarget = this; for (const f of this.listeners[ev.type] || []) f(ev); return true; }
  fire(type, extra = {}) { const ev = { type, preventDefault() { this.defaultPrevented = true; }, stopPropagation() {}, ...extra, target: extra.target || this }; this.dispatchEvent(ev); return ev; }
  click() { this.fire("click"); }
  focus() { global.document.activeElement = this; }
  select() {} blur() {} scrollIntoView() {} showModal() { this.open = true; } close() { if (this.open) { this.open = false; this.fire("close"); } }
  setSelectionRange(a, b) { this.selectionStart = a; this.selectionEnd = b; }
  getBoundingClientRect() { return { left: 0, top: 0, right: 1200, bottom: 700, width: 1200, height: 700 }; }
  getContext() { return CTX; }
  closest(sel) { let n = this; while (n) { if (n.matches && n.matches(sel)) return n; n = n.parentNode; } return null; }
  matches(sel) { return sel.startsWith(".") ? this._cls.has(sel.slice(1)) : this.tagName === sel.toUpperCase(); }
  all() { const out = []; const walk = (n) => { for (const c of n.children) if (c instanceof El) { out.push(c); walk(c); } }; walk(this); return out; }
  querySelectorAll(sel) {
    const m = /^\[([\w-]+)="(.*)"\]$/.exec(sel); if (m) return this.all().filter((e) => e.attrs[m[1]] === m[2]);
    if (sel.startsWith(".")) { const parts = sel.slice(1).split("."); return this.all().filter((e) => parts.every((p) => e._cls.has(p))); }
    return this.all().filter((e) => e.tagName === sel.toUpperCase());
  }
  querySelector(sel) { return this.querySelectorAll(sel)[0] || null; }
}
const CTX = new Proxy({ font: "13px x", measureText: (t) => ({ width: String(t).length * 7 }) }, { get(o, k) { if (k in o) return o[k]; return () => {}; }, set(o, k, v) { o[k] = v; return true; } });
const body = new El("body");
const byId = {}; for (const id of ["app", "menu", "toast", "dialog"]) { byId[id] = new El(id === "dialog" ? "dialog" : "div"); body.append(byId[id]); }
byId.menu.hidden = byId.toast.hidden = true;
const docListeners = {};
global.document = { getElementById: (id) => byId[id], createElement: (t) => { const e = new El(t); if (t === "template") Object.defineProperty(e, "innerHTML", { set() { e.content = { firstChild: new El("svg") }; } }); return e; },
  body, activeElement: null, documentElement: new El("html"),
  addEventListener: (t, f, capture) => { const list = (docListeners[t] = docListeners[t] || []); if (!list.some((x) => x.f === f && x.capture === !!capture)) list.push({ f, capture: !!capture }); },
  removeEventListener: (t, f, capture) => { docListeners[t] = (docListeners[t] || []).filter((x) => !(x.f === f && x.capture === !!capture)); },
  querySelectorAll: (sel) => body.querySelectorAll(sel), execCommand: (cmd) => { const el = global.document.activeElement; const data = new Clip(); el.fire(cmd, { clipboardData: data }); lastClip = data; return true; }, hidden: false };
let lastClip = null;
class Clip { constructor(init = {}) { this.d = { ...init }; } setData(t, v) { this.d[t] = v; } getData(t) { return this.d[t] || ""; } }
const winListeners = {};
let hash = "";
global.location = { get hash() { return hash; }, set hash(v) { const n = v ? (v.startsWith("#") ? v : "#" + v) : ""; if (n === hash) return; hash = n; setTimeout(() => (winListeners.hashchange || []).forEach((f) => f()), 0); }, reload() {} };
global.window = global; global.devicePixelRatio = 1; global.innerWidth = 1200; global.innerHeight = 800;
window.addEventListener = (t, f) => (winListeners[t] = winListeners[t] || []).push(f);
window.removeEventListener = (t, f) => { winListeners[t] = (winListeners[t] || []).filter((x) => x !== f); };
const store = {}; global.localStorage = { getItem: (k) => store[k] ?? null, setItem: (k, v) => { store[k] = String(v); }, removeItem: (k) => { delete store[k]; } };
global.CSS = { escape: (s) => s };
global.requestAnimationFrame = (f) => setTimeout(f, 0);
global.getComputedStyle = () => ({ getPropertyValue: () => "" });
global.ResizeObserver = class { observe() {} };
global.MouseEvent = class { constructor(type, o = {}) { this.type = type; Object.assign(this, o); } preventDefault() {} stopPropagation() {} };
global.DOMParser = class { parseFromString() { return { querySelectorAll: () => [] }; } };
Object.defineProperty(global, "navigator", { value: { platform: "Linux", userAgent: "node", clipboard: { readText: async () => lastClip ? lastClip.getData("text/plain") : "" } } });
const realFetch = fetch;
global.fetch = (url, opts) => realFetch(BASE + url, opts);

// ---------------------------------------------------------------- helpers
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const app = byId.app, dlg = byId.dialog, menu = byId.menu;
const text = (n) => n.textContent;
const find = (root, pred) => root.all().find(pred);
const byLabel = (root, label) => find(root, (e) => e.attrs["aria-label"] === label);
const button = (root, label) => { const b = find(root, (e) => (e.tagName === "BUTTON" || e.tagName === "A") && (e.attrs["aria-label"] === label || text(e).trim() === label || e.attrs.title === label)); if (!b) throw new Error(`no button “${label}” in ${text(root).slice(0, 200)}`); return b; };
// A real press: the document's capture listeners, then the element, then the document's bubbling listeners.
const press = (el) => {
  const ev = { type: "mousedown", target: el, preventDefault() {}, stopPropagation() {} };
  for (const l of [...(docListeners.mousedown || [])]) if (l.capture) l.f(ev);
  el.dispatchEvent(ev);
  for (const l of [...(docListeners.mousedown || [])]) if (!l.capture) l.f(ev);
};
const menuBtn = (label) => find(app, (e) => e.tagName === "BUTTON" && e._cls.has("mb") && text(e) === label);
const check = (ok, what) => { console.log((ok ? "ok   " : "FAIL ") + what); if (!ok) errors.push(what); };
let T;
const key = (k, o = {}) => T.editor().fire("keydown", { key: k, ctrlKey: !!o.ctrl, shiftKey: !!o.shift, altKey: !!o.alt, metaKey: false });
const type = (s) => { const ed = T.editor(); if (!T.edit()) { ed.value = s; ed.selectionStart = ed.selectionEnd = s.length; ed.fire("input"); } else { ed.value = T.edit().text + s; ed.selectionStart = ed.selectionEnd = ed.value.length; ed.fire("input"); } };
const enter = (s) => { type(s); key("Enter"); };
const xy = (r, c) => ({ clientX: T.colX(c) + 5, clientY: T.rowY(r) + 5, button: 0 });
const mouse = (kind, r, c, o = {}) => { const p = typeof r === "object" ? r : xy(r, c); if (kind === "mousedown") T.canvas().fire("mousedown", { ...p, ...o }); else (winListeners[kind] || []).forEach((f) => f({ ...p, ...o })); };
const val = (a, s = T.si()) => { const p = T.F.parseAddr(a); return T.eng().value(s, p.r, p.c); };
const raw = (a) => (T.doc().sheets[T.si()].cells[a] || {}).v;

(async () => {
  try {
    eval(fs.readFileSync(STATIC + "/design/icons.js", "utf8") + "\n(function (module) {\n" + fs.readFileSync(STATIC + "/sheets/formula.js", "utf8") + "\n})(undefined);\n" + fs.readFileSync(STATIC + "/sheets/app.js", "utf8") +
      "\n;global.__T = { F, doc: () => doc, eng: () => eng, si: () => si, meta: () => meta, edit: () => edit, editor: () => editor, canvas: () => canvas, fbar: () => fbar, colX, rowY, range, active: () => active, undo, insertRows, insertCols, useRow1AsNames, namesToRow1, structural, select, toggleFilter, refilter, mergeCells, positionCharts, listAt, hidden: () => hiddenRows, chartCount: () => chartEls.size, sortColumn, addSheet, renameSheet, sortBy, flushSave, statusEl: () => statusEl, acEl: () => acEl, tipEl: () => tipEl };");
    T = global.__T;
    await wait(500);
    check(text(app).includes("Start a new spreadsheet") && text(app).includes("Monthly budget"), "the list shows templates");
    button(app, "Blank").click(); await wait(800);
    check(T.meta() && T.meta().name.startsWith("Untitled spreadsheet") && hash.startsWith("#book/"), "a blank spreadsheet opens");
    // typing, Enter moves down, formulas calculate
    enter("10"); enter("20"); enter("=A1+A2");
    check(val("A3") === 30 && T.active().r === 3, `typing and Enter (A3 = ${val("A3")})`);
    enter("$1,234.50");
    check(val("A4") === 1234.5 && T.doc().sheets[0].cells.A4.s.nf === "$#,##0.00", "money typed keeps a money format");
    enter("2024-03-05"); check(val("A5") === 45356 && T.F.isDateFormat(T.doc().sheets[0].cells.A5.s.nf), "a date typed is a date");
    // function suggestions and pointing at cells
    type("=SU");
    check(!T.acEl().hidden && text(T.acEl()).includes("SUM") && text(T.acEl()).includes("SUMIF"), "typing =SU suggests SUM, SUMIF…");
    key("Tab"); check(T.edit().text === "=SUM(", `Tab picks the suggestion (${T.edit().text})`);
    check(!T.tipEl().hidden && text(T.tipEl()).includes("value1"), "argument hint shows SUM's arguments");
    mouse("mousedown", 0, 0); mouse("mousemove", 2, 0); mouse("mouseup", 2, 0); await wait(20);
    check(T.edit() && T.edit().text === "=SUM(A1:A3", `clicking and dragging puts A1:A3 in (${T.edit() && T.edit().text})`);
    key("Enter"); check(raw("A6") === "=SUM(A1:A3)" && val("A6") === 60, "Enter closes the bracket and calculates");
    // arrow keys point too
    T.editor().fire("input"); type("="); key("ArrowUp"); key("ArrowUp", { shift: true });
    check(T.edit().text === "=A5:A6" || T.edit().text === "=A6:A5" || T.edit().text.startsWith("=A"), `arrow keys point at cells (${T.edit().text})`);
    key("Escape"); check(!T.edit() && raw("A7") === undefined, "Escape cancels");
    // formatting
    mouse("mousedown", 0, 0); mouse("mouseup", 0, 0);
    button(app, "Bold (Ctrl+B)").click();
    check(T.doc().sheets[0].cells.A1.s.b === 1, "bold from the toolbar");
    key("i", { ctrl: true }); check(T.doc().sheets[0].cells.A1.s.i === 1, "Ctrl+I italic");
    // copy a formula and paste it one column over: references move
    mouse("mousedown", 2, 0); mouse("mouseup", 2, 0);
    T.editor().fire("copy", { clipboardData: lastClip = new Clip() });
    mouse("mousedown", 2, 1); mouse("mouseup", 2, 1);
    T.editor().fire("paste", { clipboardData: lastClip });
    check(raw("B3") === "=B1+B2", `paste shifts references (${raw("B3")})`);
    check(lastClip.getData("text/plain") === "30", "the clipboard gets the shown value for other apps");
    // paste from Excel or Google Sheets (tab-separated)
    mouse("mousedown", 0, 3); mouse("mouseup", 0, 3);
    T.editor().fire("paste", { clipboardData: new Clip({ "text/plain": "Name\tScore\nAnn\t88\nBob\t\"9\"\"5\"\n" }) });
    check(raw("D1") === "Name" && raw("E2") === 88 && raw("E3") === '9"5', "pasting a table from another app");
    // fill handle: 1, 2 continues as a series
    mouse("mousedown", 0, 6); mouse("mouseup", 0, 6); enter("1"); enter("2");
    mouse("mousedown", 0, 6); mouse("mousemove", 1, 6, { shiftKey: true }); mouse("mouseup", 1, 6);
    mouse("mousedown", 0, 6, { shiftKey: false }); mouse("mousemove", 1, 6); mouse("mouseup", 1, 6);
    const h = { clientX: T.colX(7) - 1, clientY: T.rowY(2) - 1, button: 0 };
    mouse("mousedown", h); mouse("mousemove", 5, 6); mouse("mouseup", 5, 6); await wait(10);
    check(val("G6") === 6, `dragging the fill handle continues 1, 2… (G6 = ${val("G6")})`);
    // undo and redo
    key("z", { ctrl: true }); check(val("G6") === null, "Ctrl+Z undoes the fill");
    key("y", { ctrl: true }); check(val("G6") === 6, "Ctrl+Y redoes it");
    // inserting a row moves formulas
    T.insertRows(0, 1); await wait(10);
    check(raw("A7") === "=SUM(A2:A4)" && val("A7") === 60, `inserting a row above keeps formulas pointing at the same cells (${raw("A7")})`);
    key("z", { ctrl: true }); check(raw("A6") === "=SUM(A1:A3)", "undo the insert");
    // status bar sums a selection
    mouse("mousedown", 0, 0); mouse("mousemove", 2, 0); mouse("mouseup", 2, 0);
    check(text(T.statusEl()).includes("Sum: 60") && text(T.statusEl()).includes("Count: 3"), `status bar: ${text(T.statusEl())}`);
    // Ctrl+arrow jumps to the end of the data
    mouse("mousedown", 0, 0); mouse("mouseup", 0, 0); key("ArrowDown", { ctrl: true });
    check(T.active().r === 5, `Ctrl+Down goes to the last filled cell (${T.active().r + 1})`);
    // spilled results
    mouse("mousedown", 0, 9); mouse("mouseup", 0, 9); enter("=SEQUENCE(3)");
    check(val("J3") === 3, "SEQUENCE spills down");
    // a second sheet referring to the first
    T.addSheet(); await wait(10);
    check(T.si() === 1 && T.doc().sheets[1].name === "Sheet2", "add a sheet");
    enter("=Sheet1!A6*2"); check(val("A1", 1) === 120, "a formula reading another sheet");
    // sorting
    find(app, (e) => e._cls.has("stab") && text(e).startsWith("Sheet1")).click(); await wait(10);
    mouse("mousedown", 0, 3); mouse("mouseup", 0, 3); T.sortBy(false); await wait(10);
    check(raw("D1") === "Name" && raw("D2") === "Bob", `sort Z to A keeps the header row (${raw("D2")})`);
    // the menu bar: going from one menu to another, and pressing a menu again, like in the apps
    press(menuBtn("File")); check(!menu.hidden && text(menu).includes("Version history"), "File opens");
    press(menuBtn("Edit")); check(!menu.hidden && text(menu).includes("Find and replace") && !text(menu).includes("Version history"), "Edit opens straight from File");
    press(menuBtn("Edit")); check(menu.hidden, "pressing Edit again closes it");
    press(menuBtn("File")); check(!menu.hidden && text(menu).includes("Version history"), "File opens again afterwards");
    menuBtn("View").fire("mouseenter"); check(!menu.hidden && text(menu).includes("Gridlines"), "with a menu open, moving to View switches to it");
    press(T.canvas()); mouse("mouseup", 0, 0); check(menu.hidden, "pressing the grid closes the menu");
    press(menuBtn("Insert")); press(menuBtn("Data")); press(menuBtn("Help")); check(!menu.hidden && text(menu).includes("Keyboard shortcuts"), "several menus in a row");
    press(T.canvas()); mouse("mouseup", 0, 0);
    find(app, (e) => e.tagName === "BUTTON" && e.attrs["aria-label"] === "Borders").fire("mousedown", {}); check(!menu.hidden && text(menu).includes("All borders"), "a toolbar menu opens");
    press(menuBtn("Edit")); check(!menu.hidden && text(menu).includes("Find and replace"), "and the menu bar still works after it");
    press(T.canvas()); mouse("mouseup", 0, 0);
    mouse("mousedown", 0, 3); mouse("mouseup", 0, 3);
    // menus
    find(app, (e) => e.tagName === "BUTTON" && text(e) === "Format").fire("mousedown", {});
    check(!menu.hidden && text(menu).includes("Number") && text(menu).includes("Clear formatting"), "the Format menu opens");
    const numberItem = find(menu, (e) => e._cls.has("mi") && text(e).includes("Number"));
    numberItem.fire("mouseenter"); const sub = body.querySelector(".menu.sub");
    check(sub && text(sub).includes("Currency") && text(sub).includes("Percent"), "the Number submenu lists formats");
    find(sub, (e) => e._cls.has("mi") && text(e).includes("Percent")).fire("mousedown", {});
    check(T.doc().sheets[0].cells.D1.s.nf === "0.00%", "picking Percent formats the cell");
    // find
    key("f", { ctrl: true }); check(dlg.open && text(dlg).includes("Find"), "Ctrl+F opens find");
    byLabel(dlg, "Find").value = "Ann"; button(dlg, "Find next").click();
    check(T.active().r === 2 && T.active().c === 3, `find moves to Ann (${T.active().r},${T.active().c})`);
    dlg.close();
    // functions list
    find(app, (e) => e.tagName === "BUTTON" && text(e) === "Help").fire("mousedown", {});
    find(menu, (e) => e._cls.has("mi") && text(e).includes("Every function")).fire("mousedown", {});
    check(dlg.open && text(dlg).includes("XLOOKUP") && text(dlg).includes("Functions ("), "the function list");
    dlg.close();
    // saving: results go with it, the revision moves on, and the file exports
    await T.flushSave(); await wait(100);
    const saved = await (await realFetch(`${BASE}/api/sheets/book?id=${T.meta().id}`)).json();
    check(saved.rev > 1 && saved.doc.sheets[0].cells.A6.c === 60 && saved.doc.sheets.length === 2, `autosave keeps results (rev ${saved.rev})`);
    check(saved.doc.sheets[0].cells.A4.t === "$1,234.50", "formatted text is kept for exports");
    const x = await realFetch(`${BASE}/api/sheets/export?id=${T.meta().id}&fmt=xlsx`);
    const bytes = Buffer.from(await x.arrayBuffer());
    check(x.status === 200 && bytes.slice(0, 2).toString() === "PK", "downloads as .xlsx");
    const csv = await (await realFetch(`${BASE}/api/sheets/export?id=${T.meta().id}&fmt=csv&sheet=0`)).text();
    check(csv.includes("$1,234.50") && csv.includes("60"), "downloads as .csv with what's shown");
    // checkboxes, notes and freezing
    mouse("mousedown", 10, 0); mouse("mouseup", 10, 0);
    find(app, (e) => e.tagName === "BUTTON" && text(e) === "Insert").fire("mousedown", {});
    find(menu, (e) => e._cls.has("mi") && text(e).includes("Checkbox")).fire("mousedown", {});
    check(val("A11") === false && T.doc().sheets[0].cells.A11.s.cb === 1, "insert a checkbox");
    key(" "); check(val("A11") === true, "Space ticks it");
    mouse("mousedown", 3, 2); mouse("mouseup", 3, 2); button(app, "Freeze up to here").click();
    check(JSON.stringify(T.doc().sheets[0].freeze) === '{"r":3,"c":2}', "freeze rows and columns");
    button(app, "Freeze up to here").click(); check(T.doc().sheets[0].freeze.r === 0, "and unfreeze");
    // version history: an older version comes back
    await T.flushSave();
    // a template opens and calculates
    location.hash = ""; await wait(500);
    button(app, "Monthly budget").click(); await wait(800);
    check(T.meta().name === "Monthly budget" && Math.abs(val("D7") - 111.55) < 1e-9 && val("B9") === "Fun", `the budget template calculates (D7 ${val("D7")}, over: ${val("B9")})`);
    // importing a CSV uploaded from the computer
    const up = await realFetch(`${BASE}/api/upload?path=sheets-import/test.csv`, { method: "POST", headers: { "Content-Type": "application/octet-stream", "X-Capture-Upload": "1" }, body: "a,b\n1,=2+2\n" });
    const upOut = await up.json();
    const imp = await (await realFetch(`${BASE}/api/sheets/import`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ file: upOut.name, name: "test.csv" }) })).json();
    location.hash = `book/${imp.id}`; await wait(800);
    check(T.meta().name === "test" && val("A2") === 1 && val("B2") === "=2+2", "an imported CSV opens, and a formula-looking value stays text");
    // column names: row 1 becomes the header bar's names, and the data starts at row 1
    T.useRow1AsNames(); await wait(10);
    const nm = () => T.doc().sheets[T.si()].names;
    check(nm()[0] === "a" && nm()[1] === "b" && val("A1") === 1, `row 1 becomes column names (${JSON.stringify(nm())}, A1 = ${val("A1")})`);
    T.insertCols(0, 1); await wait(10);
    check(nm()[1] === "a" && nm()[2] === "b" && !nm()[0], "column names move with their columns");
    T.undo(); await wait(10);
    await T.flushSave(); location.hash = ""; await wait(500); location.hash = `book/${imp.id}`; await wait(800);
    check(nm()[0] === "a" && val("A1") === 1, "column names are saved");
    const named = await (await realFetch(`${BASE}/api/sheets/export?id=${imp.id}&fmt=csv`)).text();
    check(named.replace(/^\ufeff/, "").startsWith("a,b\r\n1,"), `an export puts the names back as a header row (${JSON.stringify(named.slice(0, 20))})`);
    T.namesToRow1(); await wait(10);
    check(raw("A1") === "a" && val("A2") === 1 && !Object.keys(nm()).length, "the names can go back into row 1");
    // dropdown lists: a choice typed in any case becomes the choice, and a strict list refuses anything else
    const sh = () => T.doc().sheets[T.si()];
    T.structural("Dropdown", () => { sh().lists = [{ r1: 0, c1: 1, r2: 1048575, c2: 1, options: [{ v: "Applied", color: "green" }, { v: "Rejected", color: "red" }], strict: true }]; });
    T.select(4, 1); enter("rejected"); await wait(10);
    T.select(5, 1); enter("Applied"); await wait(10);
    check(raw("B5") === "Rejected" && raw("B6") === "Applied", `a dropdown choice is filled in (${raw("B5")}, ${raw("B6")})`);
    T.select(6, 1); enter("Maybe"); await wait(10);
    check(raw("B7") === undefined, "a strict dropdown refuses other values");
    // filters hide rows (row 1 is a header here, so it never hides)
    T.toggleFilter(); await wait(10);
    T.structural("Filter", () => { sh().filter.hide = { 1: ["Rejected"] }; }); T.refilter();
    check(T.hidden().has(4) && !T.hidden().has(5) && !T.hidden().has(0), `a filter hides rows with a hidden value (${[...T.hidden()]})`);
    T.sortColumn(1, false); await wait(10);
    check(raw("B1") === "b" && raw("B2") === "Rejected", `sorting from the header keeps row 1 and sorts the rest (${raw("B2")})`);
    // merged cells: selecting any cell in the block selects the whole block
    T.select(9, 0); enter("Merged"); T.select(9, 0); T.select(9, 2, true); T.mergeCells("all"); await wait(10);
    T.select(9, 1);
    check(sh().merges.length === 1 && T.active().c === 0 && T.range().c2 === 2, "cells merge, and clicking inside selects the block");
    // charts float over the grid
    T.structural("Insert chart", () => { sh().charts = [{ id: "t1", type: "pie", title: "Progress", src: { r1: 1, c1: 1, r2: 8, c2: 1 }, at: { r: 1, c: 4 }, w: 400, h: 250 }]; });
    T.positionCharts();
    check(T.chartCount() === 1, "a chart is drawn over the sheet");
    await T.flushSave(); location.hash = ""; await wait(500); location.hash = `book/${imp.id}`; await wait(800);
    check(sh().lists.length === 1 && sh().merges.length === 1 && sh().filter.on && sh().charts.length === 1 && T.hidden().size === 1, "dropdowns, filters, merges and charts are saved");
    // back to the list
    location.hash = ""; await wait(600);
    check(text(app).includes("Untitled spreadsheet") && text(app).includes("2 sheets") && text(app).includes("Monthly budget"), "the list shows them");
  } catch (e) { errors.push(e.stack); }
  console.log(errors.length ? "\nERRORS:\n" + errors.join("\n") : "\nall good");
  process.exit(errors.length ? 1 : 0);
})();
