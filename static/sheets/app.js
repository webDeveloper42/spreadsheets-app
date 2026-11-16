"use strict";
// Sheets: a spreadsheet for the dashboard, built from how Excel and Google Sheets work (no code or artwork from either).
// The grid is drawn on a canvas, like Google Sheets, so big sheets stay quick. Formulas come from formula.js.
// Everything you type reaches the page through input values and textContent, never innerHTML; only icons() is parsed.
const F = window.Formula;
const app = document.getElementById("app");
const menuEl = document.getElementById("menu");
const dialog = document.getElementById("dialog");
const toastEl = document.getElementById("toast");
Object.assign(ICON_PATHS, {
  undo: "M9 14L4 9l5-5M4 9h10a6 6 0 0 1 0 12h-3", redo: "M15 14l5-5-5-5M20 9H10a6 6 0 0 0 0 12h3", fx: "M14 4h-2a3 3 0 0 0-3 3v13M6 10h7M15 13l5 6M20 13l-5 6",
  alignL: "M4 6h16M4 10h10M4 14h16M4 18h10", alignC: "M4 6h16M7 10h10M4 14h16M7 18h10", alignR: "M4 6h16M10 10h10M4 14h16M10 18h10",
  wrap: "M4 6h16M4 12h13a3 3 0 0 1 0 6h-4M15 16l-2 2 2 2M4 18h5", border: "M4 4h16v16H4zM4 12h16M12 4v16", fill: "M5 13l7-7 7 7-7 7zM19 15s2 2.5 2 4a2 2 0 0 1-4 0c0-1.5 2-4 2-4",
  textColor: "M6 18L12 4l6 14M8.5 13h7", sortUp: "M7 18V6M4 9l3-3 3 3M13 7h7M13 12h5M13 17h3", sortDown: "M7 6v12M4 15l3 3 3-3M13 7h3M13 12h5M13 17h7",
  freeze: "M4 4h16v16H4zM4 9h16M9 4v16", sum: "M17 5H7l6 7-6 7h10", plusMinus: "M12 5v6M9 8h6M9 17h6", decMore: "M3 17h.01M7 12a3 3 0 1 0 6 0 3 3 0 1 0-6 0M15 12a3 3 0 1 0 6 0 3 3 0 1 0-6 0M14 20h7M18 17l3 3-3 3",
  decLess: "M3 17h.01M7 12a3 3 0 1 0 6 0 3 3 0 1 0-6 0M14 20h7M17 17l-3 3 3 3", strike: "M5 12h14M16 6c-1-1-2.5-2-4.5-2C9 4 7.5 5.5 7.5 7.5M8 18c1 1 2.5 2 4.5 2 2.5 0 4-1.5 4-3.5",
  under: "M7 4v7a5 5 0 0 0 10 0V4M5 20h14", ital: "M10 4h8M6 20h8M14 4l-4 16", bold: "M7 4h6a4 4 0 0 1 0 8H7zM7 12h7a4 4 0 0 1 0 8H7z", history: "M3 12a9 9 0 1 0 3-6.7L3 8M3 3v5h5M12 7v5l3 2",
  dollar: "M12 3v18M17 7c-1-1.5-2.8-2-5-2-2.8 0-4.5 1.3-4.5 3.3 0 4.7 9.5 2.7 9.5 7.4 0 2-1.8 3.3-4.6 3.3-2.3 0-4-.7-5.2-2.3", pct: "M19 5L5 19M7 7h.01M17 17h.01",
  check2: "M4 4h16v16H4zM8 12l3 3 5-6", note3: "M5 3h14v12l-6 6H5zM13 21v-6h6", menuDots: "M12 5h.01M12 12h.01M12 19h.01", vmid: "M4 12h16M12 4v4M12 16v4",
});

// ---------------------------------------------------------------- state
let books = [], templates = [];         // the workbook list
let meta = null, doc = null, eng = null; // the open workbook: {id, name, rev}, its JSON, its formula engine
let si = 0;                              // active sheet
let active = { r: 0, c: 0 }, cursor = { r: 0, c: 0 }, selKind = "cells"; // selection: active cell and the far corner
let edit = null;                         // { r, c, text, point, viaBar, original }
let undoStack = [], redoStack = [], clip = null, copyBox = null, showFormulas = false, gridlines = true, zoom = 1;
let saveTimer = null, saving = false, dirty = false, conflict = false, busy = false, toastTimer = null;
let canvas, ctx, scroller, spacer, editor, nameBox, fbar, statusEl, tabsEl, acEl, tipEl, titleEl, savedEl, wrapEl;
let viewW = 0, viewH = 0, scrollX = 0, scrollY = 0, raf = 0, drag = null, hoverNote = null;
try { zoom = Number(localStorage.getItem("sheets-zoom")) || 1; gridlines = localStorage.getItem("sheets-grid") !== "0"; } catch (e) { /* defaults */ }

const DEF_W = 100, DEF_H = 21, HW = 46, HH = 24, FONT = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif';
const sheet = () => doc.sheets[si];
const A = (r, c) => F.addr(r, c);
const cellAt = (r, c, s = sheet()) => s.cells[A(r, c)];
const styleAt = (r, c) => (cellAt(r, c) || {}).s || {};
const range = () => ({ r1: Math.min(active.r, cursor.r), c1: Math.min(active.c, cursor.c), r2: Math.max(active.r, cursor.r), c2: Math.max(active.c, cursor.c) });
const isMac = /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent || "");
const mod = (e) => (isMac ? e.metaKey : e.ctrlKey);
const KEY = isMac ? "⌘" : "Ctrl+";

function h(tag, attrs = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v == null || v === false) continue;
    if (k === "class") el.className = v;
    else if (k === "style") el.style.cssText = v;
    else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
    else if (k === "value" || k === "checked" || k === "disabled") el[k] = v;
    else el.setAttribute(k, v === true ? "" : v);
  }
  for (const kid of kids.flat(Infinity)) if (kid != null && kid !== false) el.append(kid);
  return el;
}
function ic(name, size = 18) { const t = document.createElement("template"); t.innerHTML = icon(name, size); return t.content.firstChild; }
async function get(path) { const r = await fetch(path); const out = await r.json().catch(() => ({})); if (!r.ok) throw new Error(out.error || "Something went wrong"); return out; }
async function act(action, data = {}) {
  const r = await fetch("/api/sheets/do", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ action, ...data }) });
  const out = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(out.error || "Something went wrong");
  return out;
}
function toast(text, bad) {
  if (!text) return;
  toastEl.textContent = text; toastEl.className = bad ? "bad" : ""; toastEl.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { toastEl.hidden = true; }, bad ? 7000 : 2600);
}
const fail = (e) => toast(e.message || "Can't reach the server. Is the tunnel still open?", true);

// ---------------------------------------------------------------- routing: #  (the list) or #book/ID
window.addEventListener("hashchange", route);
async function route() {
  const m = /^#book\/(\d+)$/.exec(location.hash);
  if (meta && (!m || Number(m[1]) !== meta.id)) { await flushSave(); closeBook(); }
  if (m) { if (!meta) await openBook(Number(m[1])); }
  else await showHome();
}

// ---------------------------------------------------------------- the workbook list
async function showHome() {
  document.title = "Sheets";
  try { const s = await get("/api/sheets/state"); books = s.books; templates = s.templates; } catch (e) { fail(e); }
  const fileIn = h("input", { type: "file", accept: ".xlsx,.xlsm,.csv,.tsv,.txt", hidden: true, onchange: (e) => importFile(e.target.files[0]) });
  const when = (t) => { if (!t) return ""; const d = new Date(t.replace(" ", "T")); const today = new Date(); return d.toDateString() === today.toDateString() ? d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }) : d.toLocaleDateString([], { month: "short", day: "numeric", year: d.getFullYear() === today.getFullYear() ? undefined : "numeric" }); };
  const tpl = (key, name, art) => h("button", { class: "tpl", onclick: () => newBook(key) }, h("div", { class: `tpl-art ${key || "blank"}` }, art), h("span", {}, name));
  app.replaceChildren(h("div", { class: "home" },
    h("header", { class: "home-head" }, h("div", { class: "logo" }, ic("grid", 26)), h("div", { class: "grow" }, h("h1", {}, "Sheets"),
      h("p", {}, "Spreadsheets with the formulas and shortcuts you know from Excel and Google Sheets. Saved on this server as you type."))),
    h("section", {}, h("h2", {}, "Start a new spreadsheet"),
      h("div", { class: "tpls" }, tpl(null, "Blank", ic("plus", 34)),
        templates.map((t) => tpl(t.key, t.name, h("div", { class: "mini-grid" }, Array.from({ length: 12 }, (_, i) => h("i", { class: i < 3 ? "hd" : "" }))))),
        h("button", { class: "tpl", onclick: () => fileIn.click() }, h("div", { class: "tpl-art import" }, ic("download", 30)), h("span", {}, "Import .xlsx or .csv")), fileIn)),
    h("section", {}, h("h2", {}, "Your spreadsheets"),
      books.length ? h("div", { class: "books" }, books.map((b) => h("div", { class: "brow" },
        h("a", { class: "bname", href: `#book/${b.id}` }, ic("grid", 18), h("span", {}, h("b", {}, b.name), h("small", {}, `${b.sheets.length} sheet${b.sheets.length === 1 ? "" : "s"} · ${b.sheets.slice(0, 4).join(", ")}${b.sheets.length > 4 ? "…" : ""}`))),
        h("small", { class: "when" }, when(b.opened || b.updated)),
        h("button", { class: "ghost", title: "More", "aria-label": `More for ${b.name}`, onclick: (e) => openMenu(e.currentTarget, bookMenu(b)) }, ic("menuDots", 16)))))
        : h("div", { class: "empty" }, "Nothing yet. Start a blank spreadsheet or a template above, or import one from Excel or Google Sheets (download it there as .xlsx)."))));
  const drop = app.querySelector(".home");
  drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("over"));
  drop.addEventListener("drop", (e) => { e.preventDefault(); drop.classList.remove("over"); if (e.dataTransfer.files[0]) importFile(e.dataTransfer.files[0]); });
}
function bookMenu(b) {
  return [
    { label: "Open", action: () => { location.hash = `book/${b.id}`; } },
    { label: "Rename", action: () => promptDialog("Rename", b.name, (name) => act("rename", { id: b.id, name }).then(showHome).catch(fail)) },
    { label: "Make a copy", action: () => act("duplicate", { id: b.id }).then(() => { toast("Copy made"); showHome(); }).catch(fail) },
    { sep: true },
    { label: "Download as Excel (.xlsx)", action: () => download(b.id, "xlsx") },
    { label: "Download first sheet as CSV", action: () => download(b.id, "csv", 0) },
    { sep: true },
    { label: "Delete", danger: true, action: () => confirmDialog(`Delete “${b.name}”? Its version history goes too.`, () => act("delete", { id: b.id }).then(() => { toast("Deleted"); showHome(); }).catch(fail)) },
  ];
}
async function newBook(template) {
  try { const out = await act("new", { template }); location.hash = `book/${out.id}`; } catch (e) { fail(e); }
}
function download(id, fmt, sheetIndex = 0) {
  const a = h("a", { href: `/api/sheets/export?id=${id}&fmt=${fmt}&sheet=${sheetIndex}`, download: "" });
  document.body.append(a); a.click(); a.remove();
}
async function importFile(file) {
  if (!file) return;
  if (file.size > 50 * 1024 * 1024) return toast("That file is over 50 MB", true);
  toast(`Importing ${file.name}…`);
  try {
    const path = `sheets-import/${Date.now()}-${file.name.replace(/[\\/]/g, "_")}`;
    const r = await fetch(`/api/upload?path=${encodeURIComponent(path)}`, { method: "POST", headers: { "Content-Type": "application/octet-stream", "X-Capture-Upload": "1" }, body: file });
    const up = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(up.error || "The upload didn't work");
    const res = await fetch("/api/sheets/import", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ file: up.name, name: file.name }) });
    const out = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(out.error || "That file couldn't be read");
    location.hash = `book/${out.id}`;
  } catch (e) { fail(e); }
}

// ---------------------------------------------------------------- opening a workbook
async function openBook(id) {
  let b;
  try { b = await get(`/api/sheets/book?id=${id}`); } catch (e) { fail(e); location.hash = ""; return; }
  meta = { id: b.id, name: b.name, rev: b.rev }; doc = b.doc;
  for (const s of doc.sheets) { s.colW = s.colW || {}; s.rowH = s.rowH || {}; s.freeze = s.freeze || { r: 0, c: 0 }; s.rows = s.rows || 1000; s.cols = s.cols || 26; s.names = s.names || {}; }
  si = Math.min(doc.active || 0, doc.sheets.length - 1);
  if (doc.sheets[si].hidden) si = Math.max(0, doc.sheets.findIndex((s) => !s.hidden));
  eng = new F.Book(doc); eng.recalc();
  undoStack = []; redoStack = []; clip = null; copyBox = null; edit = null; conflict = false; dirty = false;
  active = { r: 0, c: 0 }; cursor = { r: 0, c: 0 }; selKind = "cells";
  clearCharts(); hiddenRows = new Set();
  buildBookUI();
  refilter();
  document.title = `${meta.name} · Sheets`;
  // a workbook from a template or an import has never been calculated here; keep its results for exports
  if (doc.sheets.some((s) => Object.values(s.cells).some((c) => typeof c.v === "string" && c.v.startsWith("=") && c.c === undefined))) markDirty();
}
function closeBook() { clearCharts(); hiddenRows = new Set(); meta = null; doc = null; eng = null; window.removeEventListener("resize", onResize); }

// ---------------------------------------------------------------- the workbook page
function buildBookUI() {
  titleEl = h("input", { class: "title", value: meta.name, "aria-label": "Workbook name", spellcheck: "false",
    onchange: (e) => { const name = e.target.value.trim(); if (!name) { e.target.value = meta.name; return; } act("rename", { id: meta.id, name }).then(() => { meta.name = name; document.title = `${name} · Sheets`; }).catch(fail); },
    onkeydown: (e) => { if (e.key === "Enter" || e.key === "Escape") { if (e.key === "Escape") e.target.value = meta.name; editor.focus(); } } });
  savedEl = h("span", { class: "saved" }, "All changes saved");
  const menus = [["File", fileMenu], ["Edit", editMenu], ["View", viewMenu], ["Insert", insertMenu], ["Format", formatMenu], ["Data", dataMenu], ["Help", helpMenu]];
  nameBox = h("input", { class: "namebox", "aria-label": "Name box", spellcheck: "false", onkeydown: (e) => {
    if (e.key === "Enter") { e.preventDefault(); gotoRef(nameBox.value); editor.focus(); }
    if (e.key === "Escape") { updateBars(); editor.focus(); }
  }, onfocus: (e) => e.target.select() });
  fbar = h("textarea", { class: "fbar", rows: 1, "aria-label": "Formula bar", spellcheck: "false",
    onfocus: () => { if (!edit) startEdit(cellText(active.r, active.c), { viaBar: true, keep: true }); },
    oninput: () => { if (!edit) startEdit(fbar.value, { viaBar: true }); else { edit.text = fbar.value; editor.value = fbar.value; edit.point = null; sizeEditor(); suggest(fbar); draw(); } },
    onkeydown: (e) => editorKeys(e, fbar) });
  canvas = h("canvas", { class: "grid", "aria-label": "Spreadsheet grid" });
  spacer = h("div", { class: "spacer" });
  scroller = h("div", { class: "scroller" }, canvas, spacer);
  editor = h("textarea", { class: "editor idle", "aria-label": "Cell editor", spellcheck: "false", autocomplete: "off", autocapitalize: "off", wrap: "soft" });
  acEl = h("div", { class: "ac", hidden: true, role: "listbox" });
  tipEl = h("div", { class: "tip", hidden: true });
  wrapEl = h("div", { class: "gridwrap" }, scroller, editor, acEl, tipEl);
  statusEl = h("div", { class: "status" });
  tabsEl = h("div", { class: "tabs" });
  app.replaceChildren(h("div", { class: "book" },
    h("header", { class: "bar1" }, h("a", { class: "logo small", href: "#", title: "All spreadsheets", "aria-label": "All spreadsheets" }, ic("grid", 22)),
      h("div", { class: "grow" }, h("div", { class: "trow" }, titleEl, savedEl),
        h("nav", { class: "menubar" }, menus.map(([label, make]) => h("button", { class: "mb", onmousedown: (e) => { e.preventDefault(); openMenu(e.currentTarget, make(), true); },
          // with a menu already open, moving across the bar switches menus, like both apps
          onmouseenter: (e) => { if (!menuEl.hidden && menuFromBar && menuAnchor !== e.currentTarget) openMenu(e.currentTarget, make(), true); } }, label))))),
    toolbar(),
    h("div", { class: "bar3" }, nameBox, h("span", { class: "fx" }, ic("fx", 16)), fbar),
    wrapEl,
    h("footer", { class: "bottom" }, h("button", { class: "ghost addsheet", title: "Add a sheet", "aria-label": "Add a sheet", onclick: () => addSheet() }, ic("plus", 16)),
      h("button", { class: "ghost", title: "All sheets", "aria-label": "All sheets", onclick: (e) => openMenu(e.currentTarget, doc.sheets.map((s, i) => ({ label: s.name + (s.hidden ? " (hidden)" : ""), check: i === si, action: () => { if (s.hidden) { s.hidden = false; markDirty(); } switchSheet(i); } }))) }, ic("list", 16)),
      tabsEl, statusEl)));
  ctx = canvas.getContext("2d");
  scroller.addEventListener("scroll", () => { scrollX = scroller.scrollLeft; scrollY = scroller.scrollTop; growIfNeeded(); positionEditor(); draw(); });
  canvas.addEventListener("mousedown", onMouseDown);
  canvas.addEventListener("dblclick", onDblClick);
  canvas.addEventListener("mousemove", onHover);
  canvas.addEventListener("contextmenu", onContext);
  canvas.addEventListener("mouseleave", () => { if (hoverNote) { hoverNote = null; tipEl.hidden = true; } });
  editor.addEventListener("keydown", (e) => (edit ? editorKeys(e, editor) : gridKeys(e)));
  editor.addEventListener("input", onEditorInput);
  editor.addEventListener("copy", (e) => { if (!edit) { e.preventDefault(); copy(e.clipboardData, false); } });
  editor.addEventListener("cut", (e) => { if (!edit) { e.preventDefault(); copy(e.clipboardData, true); } });
  editor.addEventListener("paste", (e) => { if (!edit) { e.preventDefault(); paste(e.clipboardData, pasteValuesNext); pasteValuesNext = false; } });
  editor.addEventListener("compositionstart", () => { if (!edit) startEdit("", { keepIdle: true }); });
  window.addEventListener("resize", onResize);
  new ResizeObserver(onResize).observe(wrapEl);
  onResize();
  renderTabs();
  updateBars();
  editor.focus();
}
let pasteValuesNext = false;

// ---------------------------------------------------------------- toolbar
let tb = {};
function toolbar() {
  const btn = (name, title, action, key) => { const b = h("button", { class: "tb", title, "aria-label": title, onmousedown: (e) => e.preventDefault(), onclick: () => { action(); editor.focus(); } }, ic(name, 17)); if (key) tb[key] = b; return b; };
  const sep = () => h("span", { class: "tsep" });
  const nf = h("button", { class: "tb wide", title: "Number format", "aria-label": "Number format", onmousedown: (e) => { e.preventDefault(); openMenu(e.currentTarget, numberMenu()); } }, "123 ▾");
  tb.zoom = h("select", { class: "tbsel", title: "Zoom", "aria-label": "Zoom", onchange: (e) => { zoom = Number(e.target.value); try { localStorage.setItem("sheets-zoom", zoom); } catch (x) { /* fine */ } onResize(); editor.focus(); } },
    [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2].map((z) => h("option", { value: z, selected: z === zoom }, `${Math.round(z * 100)}%`)));
  tb.size = h("select", { class: "tbsel", title: "Font size", "aria-label": "Font size", onchange: (e) => { setStyle({ fs: Number(e.target.value) === 10 ? undefined : Number(e.target.value) }); editor.focus(); } },
    [6, 7, 8, 9, 10, 11, 12, 14, 16, 18, 24, 30, 36].map((n) => h("option", { value: n }, String(n))));
  const colorBtn = (name, title, key) => { const b = h("button", { class: "tb", title, "aria-label": title, onmousedown: (e) => { e.preventDefault(); openMenu(e.currentTarget, colorMenu(key)); } }, ic(name, 17), h("i", { class: "swatch" })); tb[key] = b; return b; };
  return h("div", { class: "bar2", role: "toolbar", "aria-label": "Formatting" },
    btn("undo", `Undo (${KEY}Z)`, undo), btn("redo", `Redo (${KEY}${isMac ? "⇧Z" : "Y"})`, redo), sep(), tb.zoom, sep(),
    btn("dollar", "Format as currency", () => setStyle({ nf: "$#,##0.00" })), btn("pct", "Format as percent", () => setStyle({ nf: "0%" })),
    btn("decLess", "Fewer decimal places", () => decimals(-1)), btn("decMore", "More decimal places", () => decimals(1)), nf, sep(), tb.size, sep(),
    btn("bold", `Bold (${KEY}B)`, () => toggle("b"), "b"), btn("ital", `Italic (${KEY}I)`, () => toggle("i"), "i"), btn("strike", "Strikethrough", () => toggle("st"), "st"), btn("under", `Underline (${KEY}U)`, () => toggle("u"), "u"),
    colorBtn("textColor", "Text colour", "fc"), sep(), colorBtn("fill", "Fill colour", "bg"),
    h("button", { class: "tb", title: "Borders", "aria-label": "Borders", onmousedown: (e) => { e.preventDefault(); openMenu(e.currentTarget, borderMenu()); } }, ic("border", 17)), sep(),
    btn("alignL", "Align left", () => setStyle({ ha: "left" }), "left"), btn("alignC", "Align centre", () => setStyle({ ha: "center" }), "center"), btn("alignR", "Align right", () => setStyle({ ha: "right" }), "right"),
    btn("vmid", "Middle vertically", () => setStyle({ va: styleAt(active.r, active.c).va === "middle" ? undefined : "middle" }), "middle"),
    btn("wrap", "Wrap text", () => { toggle("wr"); fitRows(range()); }, "wr"), sep(),
    btn("sortUp", "Sort A to Z by this column", () => sortBy(true)), btn("sortDown", "Sort Z to A by this column", () => sortBy(false)),
    btn("freeze", "Freeze up to here", () => toggleFreeze()),
    h("button", { class: "tb", title: "Functions", "aria-label": "Functions", onmousedown: (e) => { e.preventDefault(); openMenu(e.currentTarget, sumMenu()); } }, ic("sum", 17)));
}
function updateToolbar() {
  const s = styleAt(active.r, active.c);
  for (const k of ["b", "i", "u", "st", "wr"]) if (tb[k]) tb[k].classList.toggle("on", !!s[k]);
  for (const k of ["left", "center", "right"]) tb[k].classList.toggle("on", s.ha === k);
  tb.middle.classList.toggle("on", s.va === "middle");
  tb.size.value = String(s.fs || 10);
  tb.fc.lastChild.style.background = s.fc || "#202124";
  tb.bg.lastChild.style.background = s.bg || "transparent";
}

// ---------------------------------------------------------------- geometry
// Column widths and row heights are summed once and searched by halving, so jumping to row 100,000 is instant.
let geo = null;
function geometry() {
  const s = sheet();
  if (geo && geo.s === s && geo.cols === s.cols && geo.rows === s.rows && geo.zoom === zoom && geo.ver === geoVer) return geo;
  const cx = new Float64Array(s.cols + 1), ry = new Float64Array(s.rows + 1);
  for (let c = 0; c < s.cols; c++) cx[c + 1] = cx[c] + (s.colW[c] ?? DEF_W) * zoom;
  for (let r = 0; r < s.rows; r++) ry[r + 1] = ry[r] + (hiddenRows.has(r) ? 0 : (s.rowH[r] ?? DEF_H) * zoom);   // a filter hides rows
  geo = { s, cols: s.cols, rows: s.rows, zoom, ver: geoVer, cx, ry };
  return geo;
}
let geoVer = 0;
const relayout = () => { geoVer++; };
function search(arr, v) { let lo = 0, hi = arr.length - 2; while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (arr[mid] <= v) lo = mid; else hi = mid - 1; } return lo; }
const hw = () => HW * Math.min(zoom, 1.25), hh = () => HH * Math.min(zoom, 1.25);
function frozen() { const g = geometry(), f = sheet().freeze || { r: 0, c: 0 }; const fr = Math.min(f.r, g.rows), fc = Math.min(f.c, g.cols); return { fr, fc, fw: g.cx[fc], fh: g.ry[fr] }; }
function colX(c) { const g = geometry(), z = frozen(); return hw() + (c < z.fc ? g.cx[c] : z.fw + g.cx[c] - g.cx[z.fc] - scrollX); }
function rowY(r) { const g = geometry(), z = frozen(); return hh() + (r < z.fr ? g.ry[r] : z.fh + g.ry[r] - g.ry[z.fr] - scrollY); }
const colW = (c) => geometry().cx[c + 1] - geometry().cx[c];
const rowH = (r) => geometry().ry[r + 1] - geometry().ry[r];
function colFromX(x) { const g = geometry(), z = frozen(); const lx = x - hw(); if (lx < z.fw) return search(g.cx, Math.max(0, lx)); return Math.min(g.cols - 1, search(g.cx, lx - z.fw + g.cx[z.fc] + scrollX)); }
function rowFromY(y) { const g = geometry(), z = frozen(); const ly = y - hh(); if (ly < z.fh) return search(g.ry, Math.max(0, ly)); return Math.min(g.rows - 1, search(g.ry, ly - z.fh + g.ry[z.fr] + scrollY)); }
function onResize() {
  if (!canvas || !doc) return;
  const dpr = window.devicePixelRatio || 1;
  viewW = scroller.clientWidth; viewH = scroller.clientHeight;
  canvas.width = Math.round(viewW * dpr); canvas.height = Math.round(viewH * dpr);
  canvas.style.width = viewW + "px"; canvas.style.height = viewH + "px";
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  sizeSpacer(); positionEditor(); draw();
}
function sizeSpacer() {
  const g = geometry();
  spacer.style.width = hw() + g.cx[g.cols] + "px";
  spacer.style.height = hh() + g.ry[g.rows] + "px";
  spacer.style.marginTop = -viewH + "px";
}
// Like Google Sheets, more rows appear when you scroll near the bottom.
function growIfNeeded() {
  const g = geometry(), s = sheet();
  if (scrollY + viewH > g.ry[g.rows] - 200 && s.rows < F.MAX_ROWS) { s.rows = Math.min(F.MAX_ROWS, s.rows + 1000); sizeSpacer(); }
  if (scrollX + viewW > g.cx[g.cols] - 200 && s.cols < Math.min(F.MAX_COLS, 702)) { s.cols = Math.min(702, s.cols + 26); sizeSpacer(); }
}
// Scrolls so a cell is in view (frozen cells always are).
function reveal(r, c) {
  const g = geometry(), z = frozen();
  if (c >= z.fc) {
    const left = g.cx[c] - g.cx[z.fc], right = g.cx[c + 1] - g.cx[z.fc], room = viewW - hw() - z.fw - 14;
    if (left < scrollX) scroller.scrollLeft = left; else if (right > scrollX + room) scroller.scrollLeft = right - room;
  }
  if (r >= z.fr) {
    const top = g.ry[r] - g.ry[z.fr], bottom = g.ry[r + 1] - g.ry[z.fr], room = viewH - hh() - z.fh - 14;
    if (top < scrollY) scroller.scrollTop = top; else if (bottom > scrollY + room) scroller.scrollTop = bottom - room;
  }
  scrollX = scroller.scrollLeft; scrollY = scroller.scrollTop;
}

// ---------------------------------------------------------------- what a cell shows
function cellText(r, c) {
  const cell = cellAt(r, c);
  if (!cell || cell.v === undefined || cell.v === null) return "";
  const v = cell.v, nf = (cell.s || {}).nf;
  if (typeof v === "number") {
    if (nf && F.isDateFormat(nf)) return F.format(v, Math.floor(v) === v ? "m/d/yyyy" : v < 1 ? "h:mm:ss AM/PM" : "m/d/yyyy h:mm:ss").text;
    if (nf && /%/.test(nf)) return F.numText(v * 100) + "%";
    return F.numText(v);
  }
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  return String(v);
}
function display(r, c) {
  const cell = cellAt(r, c), s = (cell && cell.s) || {};
  const v = eng.value(si, r, c);
  if (showFormulas && cell && typeof cell.v === "string" && cell.v.startsWith("=")) return { text: cell.v, align: "left", value: v, s };
  if (s.cb && (typeof v === "boolean" || v === null)) return { text: "", check: v === true, value: v, s };
  const out = F.format(v, s.nf);
  const align = s.ha || (F.isErr(v) || typeof v === "boolean" ? "center" : typeof v === "number" ? "right" : "left");
  return { text: out.text, color: out.color, align, value: v, s, err: F.isErr(v) };
}

// ---------------------------------------------------------------- drawing
const COLORS = { red: "#d93025", blue: "#1a73e8", green: "#188038", yellow: "#f9ab00", magenta: "#c5221f", cyan: "#129eaf", black: "#202124", white: "#ffffff" };
let theme = {};
function readTheme() {
  const cs = getComputedStyle(document.documentElement);
  const v = (n) => cs.getPropertyValue(n).trim();
  theme = { bg: v("--cell") || "#fff", ink: v("--ink") || "#202124", grid: v("--gridline") || "#e2e3e3", head: v("--head") || "#f8f9fa", headInk: v("--mid") || "#5f6368",
    headOn: v("--head-on") || "#d3e3fd", headOnInk: v("--head-on-ink") || "#0b57d0", sel: v("--sel") || "rgba(14,101,235,.1)", accent: v("--accent") || "#1a73e8", line: v("--line") || "#c4c7c5" };
}
function draw() { if (!raf) raf = requestAnimationFrame(paint); }
function paint() {
  raf = 0;
  if (!ctx || !doc) return;
  readTheme();
  const g = geometry(), z = frozen(), W = viewW, H = viewH, HWp = hw(), HHp = hh();
  ctx.save();
  ctx.fillStyle = theme.bg; ctx.fillRect(0, 0, W, H);
  // visible scrolling ranges
  const c0 = z.fc + 0, cStart = Math.max(z.fc, search(g.cx, g.cx[z.fc] + scrollX)), cEnd = Math.min(g.cols - 1, search(g.cx, g.cx[z.fc] + scrollX + W - HWp - z.fw));
  const rStart = Math.max(z.fr, search(g.ry, g.ry[z.fr] + scrollY)), rEnd = Math.min(g.rows - 1, search(g.ry, g.ry[z.fr] + scrollY + H - HHp - z.fh));
  void c0;
  const regions = [
    [rStart, rEnd, cStart, cEnd, HWp + z.fw, HHp + z.fh, W, H],
    [0, z.fr - 1, cStart, cEnd, HWp + z.fw, HHp, W, HHp + z.fh],
    [rStart, rEnd, 0, z.fc - 1, HWp, HHp + z.fh, HWp + z.fw, H],
    [0, z.fr - 1, 0, z.fc - 1, HWp, HHp, HWp + z.fw, HHp + z.fh],
  ];
  for (const [ra, rb, ca, cb, x1, y1, x2, y2] of regions) if (rb >= ra && cb >= ca) drawRegion(ra, rb, ca, cb, x1, y1, x2, y2);
  drawSelection(z);
  drawHeaders(g, z, cStart, cEnd, rStart, rEnd);
  positionCharts();
  if (z.fr) { ctx.fillStyle = "#bdc1c6"; ctx.fillRect(0, HHp + z.fh - 1, W, 3); }
  if (z.fc) { ctx.fillStyle = "#bdc1c6"; ctx.fillRect(HWp + z.fw - 1, 0, 3, H); }
  ctx.restore();
}
function font(s) { const pt = s.fs || 10; return `${s.i ? "italic " : ""}${s.b ? "bold " : ""}${Math.round(pt * 1.33 * zoom)}px ${FONT}`; }
function drawRegion(ra, rb, ca, cb, x1, y1, x2, y2) {
  ctx.save();
  ctx.beginPath(); ctx.rect(x1, y1, x2 - x1, y2 - y1); ctx.clip();
  const s = sheet();
  // fills
  for (let r = ra; r <= rb; r++) for (let c = ca; c <= cb; c++) {
    const cell = s.cells[A(r, c)];
    if (cell && cell.s && cell.s.bg) { ctx.fillStyle = cell.s.bg; ctx.fillRect(colX(c), rowY(r), colW(c), rowH(r)); }
  }
  // grid lines
  if (gridlines) {
    ctx.strokeStyle = theme.grid; ctx.lineWidth = 1; ctx.beginPath();
    for (let c = ca; c <= cb + 1; c++) { const x = Math.round(colX(c)) - 0.5; ctx.moveTo(x, y1); ctx.lineTo(x, y2); }
    for (let r = ra; r <= rb + 1; r++) { const y = Math.round(rowY(r)) - 0.5; ctx.moveTo(x1, y); ctx.lineTo(x2, y); }
    ctx.stroke();
  }
  ctx.textBaseline = "alphabetic";
  drawMerges(ra, rb, ca, cb);
  // text, with long text running into empty cells to the right like both apps
  for (let r = ra; r <= rb; r++) {
    const y = rowY(r), hgt = rowH(r);
    if (hgt < 0.5) continue;
    for (let c = ca; c <= cb; c++) {
      if (s.merges && s.merges.length && mergeAt(r, c)) continue;
      const d = display(r, c);
      const x = colX(c), w = colW(c);
      if (d.check !== undefined) { drawCheckbox(x, y, w, hgt, d.check); continue; }
      const rule = s.lists && s.lists.length ? listAt(r, c) : null;
      if (rule && !showFormulas) { drawChip(rule, d, x, y, w, hgt, r === active.r && c === active.c); continue; }
      if (!d.text) { if (cellAt(r, c) && cellAt(r, c).n) drawNoteMark(x, y, w); continue; }
      const st = d.s;
      ctx.font = font(st);
      ctx.fillStyle = d.color ? (COLORS[d.color] || d.color) : st.fc || (d.err ? "#c5221f" : theme.ink);
      const pad = 3 * zoom;
      let text = d.text;
      let tw = ctx.measureText(text).width;
      let clipW = w;
      if (typeof d.value === "number" && !showFormulas && tw > w - 2 * pad) {
        text = fitNumber(d, w - 2 * pad); tw = ctx.measureText(text).width;
      }
      let size = Math.round((st.fs || 10) * 1.33 * zoom);
      if (st.fit && !showFormulas && tw > w - 2 * pad) {   // shrink to fit: smaller text instead of running over
        const k = Math.max(0.45, (w - 2 * pad) / tw); size = Math.round(size * k);
        ctx.font = `${st.i ? "italic " : ""}${st.b ? "bold " : ""}${size}px ${FONT}`; tw = ctx.measureText(text).width;
      }
      if (st.wr && !showFormulas) { drawWrapped(text, x, y, w, hgt, d.align, st, size); }
      else {
        if (d.align === "left" && tw > w - pad && typeof d.value === "string") {
          let cc = c + 1; while (cc < Math.min(s.cols, cb + 30) && !display(r, cc).text && !(cellAt(r, cc) || {}).v) { clipW += colW(cc); if (clipW > tw + pad * 2) break; cc++; }
        }
        const tx = d.align === "right" ? x + w - pad - tw : d.align === "center" ? x + (w - tw) / 2 : x + pad;
        const ty = st.va === "top" ? y + size + 2 * zoom : st.va === "middle" ? y + hgt / 2 + size * 0.36 : y + hgt - 5 * zoom;
        ctx.save(); ctx.beginPath(); ctx.rect(d.align === "left" ? x : x, y, d.align === "left" ? clipW : w, hgt); ctx.clip();
        ctx.fillText(text, tx, ty);
        if (st.u || st.st) {
          ctx.fillRect(tx, st.u ? ty + 2 : ty - size * 0.3, tw, Math.max(1, zoom));
          if (st.u && st.st) ctx.fillRect(tx, ty - size * 0.3, tw, Math.max(1, zoom));
        }
        ctx.restore();
      }
      if (cellAt(r, c) && cellAt(r, c).n) drawNoteMark(x, y, w);
    }
  }
  // borders set on cells
  ctx.strokeStyle = theme.ink; ctx.lineWidth = 1; ctx.beginPath();
  for (let r = ra; r <= rb; r++) for (let c = ca; c <= cb; c++) {
    const bd = (styleAt(r, c).bd); if (!bd) continue;
    const x = Math.round(colX(c)) - 0.5, y = Math.round(rowY(r)) - 0.5, x2b = Math.round(colX(c) + colW(c)) - 0.5, y2b = Math.round(rowY(r) + rowH(r)) - 0.5;
    if (bd.t) { ctx.moveTo(x, y); ctx.lineTo(x2b, y); } if (bd.b) { ctx.moveTo(x, y2b); ctx.lineTo(x2b, y2b); }
    if (bd.l) { ctx.moveTo(x, y); ctx.lineTo(x, y2b); } if (bd.r) { ctx.moveTo(x2b, y); ctx.lineTo(x2b, y2b); }
  }
  ctx.stroke();
  ctx.restore();
}
function fitNumber(d, room) {
  // General numbers lose decimals to fit; anything else shows #### like both apps
  if (!d.s.nf) {
    for (let p = 10; p >= 1; p--) { const t = F.format(Number(d.value.toPrecision(p)), undefined).text; if (ctx.measureText(t).width <= room) return t; }
  }
  let t = "#"; while (ctx.measureText(t + "#").width <= room) t += "#"; return t;
}
function drawWrapped(text, x, y, w, hgt, align, st, size) {
  const pad = 3 * zoom, lines = [], words = text.split(/(\s+)/);
  for (const para of text.split("\n")) {
    let line = "";
    for (const word of para.split(/(?<=\s)/)) {
      if (ctx.measureText(line + word).width > w - 2 * pad && line) { lines.push(line); line = word.trimStart(); }
      else line += word;
    }
    lines.push(line);
  }
  void words;
  const lh = size * 1.25;
  ctx.save(); ctx.beginPath(); ctx.rect(x, y, w, hgt); ctx.clip();
  const total = lines.length * lh, start = st.va === "top" ? y + size + 2 : st.va === "middle" ? y + (hgt - total) / 2 + size : y + hgt - total + size - 3;
  lines.forEach((ln, i) => { const tw = ctx.measureText(ln).width; const tx = align === "right" ? x + w - pad - tw : align === "center" ? x + (w - tw) / 2 : x + pad; ctx.fillText(ln, tx, start + i * lh); });
  ctx.restore();
}
function drawCheckbox(x, y, w, hgt, on) {
  const s = Math.round(14 * zoom), bx = Math.round(x + (w - s) / 2) + 0.5, by = Math.round(y + (hgt - s) / 2) + 0.5;
  ctx.save(); ctx.lineWidth = 1.5;
  if (on) { ctx.fillStyle = theme.accent; ctx.beginPath(); ctx.roundRect(bx, by, s, s, 2); ctx.fill(); ctx.strokeStyle = "#fff"; ctx.beginPath(); ctx.moveTo(bx + s * 0.22, by + s * 0.52); ctx.lineTo(bx + s * 0.42, by + s * 0.72); ctx.lineTo(bx + s * 0.78, by + s * 0.3); ctx.stroke(); }
  else { ctx.strokeStyle = theme.headInk; ctx.beginPath(); ctx.roundRect(bx, by, s, s, 2); ctx.stroke(); }
  ctx.restore();
}
function drawNoteMark(x, y, w) { ctx.fillStyle = "#f9ab00"; ctx.beginPath(); ctx.moveTo(x + w - 7 * zoom, y); ctx.lineTo(x + w, y); ctx.lineTo(x + w, y + 7 * zoom); ctx.fill(); }
function boxRect(r1, c1, r2, c2) {
  const x = colX(c1), y = rowY(r1);
  return { x, y, w: colX(c2) + colW(c2) - x, h: rowY(r2) + rowH(r2) - y };
}
const REF_COLORS = ["#1a73e8", "#d93025", "#9334e6", "#188038", "#e8710a", "#c5221f", "#129eaf", "#f538a0"];
function drawSelection(z) {
  const HWp = hw(), HHp = hh();
  ctx.save(); ctx.beginPath(); ctx.rect(HWp, HHp, viewW - HWp, viewH - HHp); ctx.clip();
  // spilled results of the selected formula get a thin outline, like both apps
  const spill = eng.spillAt(si, active.r, active.c);
  if (spill && !edit) { const b = boxRect(spill.r, spill.c, spill.r + spill.h - 1, spill.c + spill.w - 1); ctx.setLineDash([3, 2]); ctx.strokeStyle = theme.accent; ctx.lineWidth = 1; ctx.strokeRect(b.x + 0.5, b.y + 0.5, b.w - 1, b.h - 1); ctx.setLineDash([]); }
  // references in the formula being typed
  if (edit && edit.text.startsWith("=")) {
    F.refsIn(edit.text).forEach((ref, i) => {
      const sh = ref.sheet == null ? si : eng.sheetIndex(ref.sheet); if (sh !== si) return;
      const r2 = Math.min(ref.r2, geometry().rows - 1), c2 = Math.min(ref.c2, geometry().cols - 1);
      const b = boxRect(ref.r1, ref.c1, r2, c2), col = REF_COLORS[i % REF_COLORS.length];
      ctx.fillStyle = col + "22"; ctx.fillRect(b.x, b.y, b.w, b.h);
      ctx.strokeStyle = col; ctx.lineWidth = 2; if (edit.point && edit.point.idx === i) ctx.setLineDash([4, 3]); ctx.strokeRect(b.x + 1, b.y + 1, b.w - 2, b.h - 2); ctx.setLineDash([]);
    });
  }
  const R = range(), b = boxRect(R.r1, R.c1, R.r2, R.c2);
  if (R.r1 !== R.r2 || R.c1 !== R.c2) { ctx.fillStyle = theme.sel; ctx.fillRect(b.x, b.y, b.w, b.h); const a = boxRect(active.r, active.c, active.r, active.c); ctx.fillStyle = theme.bg; ctx.globalAlpha = 0; ctx.fillRect(a.x, a.y, a.w, a.h); ctx.globalAlpha = 1; }
  if (drag && drag.kind === "fill" && drag.to) { const f = drag.to; const fb = boxRect(f.r1, f.c1, f.r2, f.c2); ctx.setLineDash([4, 3]); ctx.strokeStyle = "#5f6368"; ctx.lineWidth = 1; ctx.strokeRect(fb.x + 0.5, fb.y + 0.5, fb.w - 1, fb.h - 1); ctx.setLineDash([]); }
  ctx.strokeStyle = theme.accent; ctx.lineWidth = 2; ctx.strokeRect(b.x + 1, b.y + 1, b.w - 2, b.h - 2);
  const a = boxRect(active.r, active.c, active.r, active.c);
  if (R.r1 !== R.r2 || R.c1 !== R.c2) { ctx.lineWidth = 1; ctx.strokeRect(a.x + 2.5, a.y + 2.5, a.w - 5, a.h - 5); }
  if (!edit) { ctx.fillStyle = theme.accent; ctx.fillRect(b.x + b.w - 4, b.y + b.h - 4, 7, 7); ctx.strokeStyle = theme.bg; ctx.lineWidth = 1; ctx.strokeRect(b.x + b.w - 4.5, b.y + b.h - 4.5, 8, 8); }
  if (copyBox && copyBox.si === si) { const cb = boxRect(copyBox.r1, copyBox.c1, copyBox.r2, copyBox.c2); ctx.setLineDash([5, 4]); ctx.lineDashOffset = -(Date.now() / 60) % 9; ctx.strokeStyle = theme.accent; ctx.lineWidth = 2; ctx.strokeRect(cb.x + 1, cb.y + 1, cb.w - 2, cb.h - 2); ctx.setLineDash([]); setTimeout(draw, 120); }
  ctx.restore();
}
function drawHeaders(g, z, cStart, cEnd, rStart, rEnd) {
  const HWp = hw(), HHp = hh(), R = range();
  ctx.font = `${Math.round(11 * Math.min(zoom, 1.25))}px ${FONT}`; ctx.textBaseline = "middle"; ctx.textAlign = "center";
  const cols = [...Array(z.fc).keys(), ...range2(cStart, cEnd)], rows = [...Array(z.fr).keys(), ...range2(rStart, rEnd)];
  ctx.fillStyle = theme.head; ctx.fillRect(0, 0, viewW, HHp); ctx.fillRect(0, 0, HWp, viewH);
  for (const c of cols) {
    const x = colX(c), w = colW(c), on = c >= R.c1 && c <= R.c2, all = selKind === "cols" || selKind === "all";
    if (x + w < HWp || x > viewW) continue;
    if (on) { ctx.fillStyle = all ? theme.accent : theme.headOn; ctx.fillRect(x, 0, w, HHp); }
    ctx.fillStyle = on ? (all ? "#fff" : theme.headOnInk) : theme.headInk;
    const named = (sheet().names || {})[c];
    if (named) {
      // Your column name in the header, with the letter formulas use kept small at the left.
      const small = Math.round(9 * Math.min(zoom, 1.25)), big = Math.round(11.5 * Math.min(zoom, 1.25));
      ctx.save(); ctx.beginPath(); ctx.rect(x, 0, w - 1, HHp); ctx.clip();
      ctx.textAlign = "left"; ctx.font = `${small}px ${FONT}`; ctx.globalAlpha = 0.55;
      const letter = F.colName(c); ctx.fillText(letter, x + 4, HHp / 2 + 1);
      const lw = ctx.measureText(letter).width + 8; ctx.globalAlpha = 1;
      ctx.font = `600 ${big}px ${FONT}`; ctx.textAlign = "center";
      let label = named; const room = w - lw * 2 - 4;
      while (label.length > 1 && ctx.measureText(label).width > room) label = label.slice(0, -1);
      if (label !== named) label = label.slice(0, -1) + "…";
      ctx.fillText(label, x + w / 2, HHp / 2 + 1);
      ctx.restore(); ctx.font = `${Math.round(11 * Math.min(zoom, 1.25))}px ${FONT}`; ctx.textAlign = "center";
    } else ctx.fillText(F.colName(c), x + w / 2, HHp / 2 + 1);
    ctx.fillStyle = theme.line; ctx.fillRect(Math.round(x + w) - 1, 0, 1, HHp);
  }
  drawFilterIcons();
  for (const r of rows) {
    const y = rowY(r), hg = rowH(r), on = r >= R.r1 && r <= R.r2, all = selKind === "rows" || selKind === "all";
    if (hg < 0.5 || y + hg < HHp || y > viewH) continue;
    if (on) { ctx.fillStyle = all ? theme.accent : theme.headOn; ctx.fillRect(0, y, HWp, hg); }
    ctx.fillStyle = on ? (all ? "#fff" : theme.headOnInk) : theme.headInk;
    ctx.fillText(String(r + 1), HWp / 2, y + hg / 2 + 1);
    ctx.fillStyle = theme.line; ctx.fillRect(0, Math.round(y + hg) - 1, HWp, 1);
  }
  ctx.fillStyle = theme.head; ctx.fillRect(0, 0, HWp, HHp);
  ctx.fillStyle = theme.line; ctx.fillRect(0, HHp - 1, viewW, 1); ctx.fillRect(HWp - 1, 0, 1, viewH);
  ctx.textAlign = "left";
}
function range2(a, b) { const out = []; for (let i = a; i <= b; i++) out.push(i); return out; }

// ---------------------------------------------------------------- selection
function select(r, c, extend = false, kind = "cells") {
  const g = geometry();
  r = Math.max(0, Math.min(r, g.rows - 1)); c = Math.max(0, Math.min(c, g.cols - 1));
  const mg = !extend && mergeAt(r, c);
  if (extend) cursor = { r, c }; else if (mg) { active = { r: mg.r1, c: mg.c1 }; cursor = { r: mg.r2, c: mg.c2 }; } else { active = { r, c }; cursor = { r, c }; }
  selKind = kind;
  reveal(extend ? r : active.r, extend ? c : active.c);
  updateBars(); positionEditor(); draw();
}
function selectAll() { const g = geometry(); active = { r: 0, c: 0 }; cursor = { r: g.rows - 1, c: g.cols - 1 }; selKind = "all"; updateBars(); draw(); }
function updateBars() {
  if (!doc) return;
  const R = range();
  nameBox.value = R.r1 === R.r2 && R.c1 === R.c2 ? A(active.r, active.c) : selKind === "cols" ? `${F.colName(R.c1)}:${F.colName(R.c2)}` : selKind === "rows" ? `${R.r1 + 1}:${R.r2 + 1}` : `${A(R.r1, R.c1)}:${A(R.r2, R.c2)}`;
  if (!edit) { fbar.value = cellText(active.r, active.c); fbar.style.height = ""; }
  updateToolbar(); updateStatus();
}
function updateStatus() {
  const R = range();
  const cells = (R.r2 - R.r1 + 1) * (R.c2 - R.c1 + 1);
  if (cells < 2) { const v = eng.value(si, active.r, active.c); const sp = eng.spillOwner(si, active.r, active.c); statusEl.replaceChildren(F.isErr(v) && v.msg ? h("span", { class: "errmsg" }, `${v.code} ${v.msg}`) : sp && !(sp.r === active.r && sp.c === active.c) ? h("span", {}, `Result of the formula in ${A(sp.r, sp.c)}`) : ""); return; }
  let sum = 0, count = 0, counta = 0, min = Infinity, max = -Infinity;
  const b = eng.bounds(si);
  for (let r = R.r1; r <= Math.min(R.r2, b.r); r++) for (let c = R.c1; c <= Math.min(R.c2, b.c); c++) {
    const v = eng.value(si, r, c);
    if (v === null || v === undefined || v === "") continue;
    counta++;
    if (typeof v === "number") { sum += v; count++; if (v < min) min = v; if (v > max) max = v; }
  }
  const nf = styleAt(active.r, active.c).nf;
  const fmt = (v) => F.format(v, nf && !F.isDateFormat(nf) ? nf : undefined).text;
  statusEl.replaceChildren(...(count ? [`Sum: ${fmt(sum)}`, `Average: ${fmt(sum / count)}`, `Min: ${fmt(min)}`, `Max: ${fmt(max)}`, `Count: ${count}`] : [`Count: ${counta}`]).map((t) => h("span", {}, t)));
}
function gotoRef(text) {
  const t = text.trim().replace(/^=/, "");
  const m = /^(?:(?:'([^']+)'|([^!]+))!)?(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?$/.exec(t);
  if (!m) { toast(`“${text}” isn't a cell like B12 or a range like A1:C9`, true); updateBars(); return; }
  const name = m[1] || m[2];
  if (name) { const i = eng.sheetIndex(name); if (i < 0) { toast(`There's no sheet called ${name}`, true); return; } switchSheet(i); }
  const a = F.parseAddr(m[3]), b = m[4] ? F.parseAddr(m[4]) : a;
  growTo(Math.max(a.r, b.r), Math.max(a.c, b.c));
  select(a.r, a.c); if (m[4]) select(b.r, b.c, true);
}
function growTo(r, c) { const s = sheet(); if (r >= s.rows) s.rows = Math.min(F.MAX_ROWS, r + 100); if (c >= s.cols) s.cols = Math.min(F.MAX_COLS, c + 5); sizeSpacer(); }

// Ctrl+arrow: to the edge of the block of filled cells, like both apps.
function jump(r, c, dr, dc) {
  const g = geometry(), b = eng.bounds(si);
  const filled = (rr, cc) => { const v = eng.value(si, rr, cc); return v !== null && v !== undefined && v !== ""; };
  const inside = (rr, cc) => rr >= 0 && cc >= 0 && rr < g.rows && cc < g.cols;
  let nr = r + dr, nc = c + dc;
  if (!inside(nr, nc)) return { r, c };
  if (filled(r, c) && filled(nr, nc)) { while (inside(nr + dr, nc + dc) && filled(nr + dr, nc + dc)) { nr += dr; nc += dc; } return { r: nr, c: nc }; }
  while (inside(nr, nc) && !filled(nr, nc)) {
    if ((dr > 0 && nr > b.r) || (dc > 0 && nc > b.c)) return dr ? { r: g.rows - 1, c } : { r, c: g.cols - 1 }; // nothing further: the last row or column
    nr += dr; nc += dc;
  }
  if (!inside(nr, nc)) return { r: dr ? (dr > 0 ? g.rows - 1 : 0) : r, c: dc ? (dc > 0 ? g.cols - 1 : 0) : c };
  return { r: nr, c: nc };
}

// ---------------------------------------------------------------- keys on the grid
function gridKeys(e) {
  const g = geometry(), R = range(), k = e.key;
  if (e.altKey && k === "ArrowDown" && listAt(active.r, active.c)) { e.preventDefault(); openListMenu(active.r, active.c); return; }
  const move = (dr, dc) => {
    e.preventDefault();
    const from = e.shiftKey ? cursor : active;
    const to = mod(e) ? jump(from.r, from.c, dr, dc) : { r: from.r + dr, c: from.c + dc };
    if (to.r >= g.rows - 1 || to.c >= g.cols - 1) growTo(to.r + 1, to.c + 1);
    select(to.r, to.c, e.shiftKey);
  };
  if (mod(e) && !e.altKey) {
    const lk = k.toLowerCase();
    if (lk === "z") { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
    if (lk === "y") { e.preventDefault(); redo(); return; }
    if (lk === "b") { e.preventDefault(); toggle("b"); return; }
    if (lk === "i") { e.preventDefault(); toggle("i"); return; }
    if (lk === "u") { e.preventDefault(); toggle("u"); return; }
    if (k === "5" && !e.shiftKey) { e.preventDefault(); toggle("st"); return; }
    if (lk === "a") { e.preventDefault(); selectAll(); return; }
    if (lk === "d") { e.preventDefault(); fillDir("down"); return; }
    if (lk === "r") { e.preventDefault(); fillDir("right"); return; }
    if (lk === "f") { e.preventDefault(); findDialog(false); return; }
    if (lk === "h") { e.preventDefault(); findDialog(true); return; }
    if (lk === "v" && e.shiftKey) { pasteValuesNext = true; return; } // the paste event follows
    if (k === ";" || k === ":") { e.preventDefault(); const now = new Date(); const v = e.shiftKey ? (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) / 86400 : F.dateSerial(now.getFullYear(), now.getMonth() + 1, now.getDate()); setValues([[v]], R.r1, R.c1, e.shiftKey ? "h:mm:ss AM/PM" : "m/d/yyyy"); return; }
    if (k === "`") { e.preventDefault(); showFormulas = !showFormulas; draw(); return; }
    if (k === "Home") { e.preventDefault(); select(0, 0, e.shiftKey); return; }
    if (k === "End") { e.preventDefault(); const b = eng.bounds(si); select(Math.max(0, b.r), Math.max(0, b.c), e.shiftKey); return; }
    if (k === " ") { e.preventDefault(); active = { r: 0, c: R.c1 }; cursor = { r: g.rows - 1, c: R.c2 }; selKind = "cols"; updateBars(); draw(); return; }
    if (e.shiftKey && (k === "!" || k === "1")) { e.preventDefault(); setStyle({ nf: "#,##0.00" }); return; }
    if (e.shiftKey && (k === "$" || k === "4")) { e.preventDefault(); setStyle({ nf: "$#,##0.00" }); return; }
    if (e.shiftKey && (k === "%" || k === "5")) { e.preventDefault(); setStyle({ nf: "0%" }); return; }
    if (e.shiftKey && (k === "#" || k === "3")) { e.preventDefault(); setStyle({ nf: "m/d/yyyy" }); return; }
    if (e.shiftKey && (k === "~" || k === "`")) { e.preventDefault(); setStyle({ nf: undefined }); return; }
    if (e.shiftKey && lk === "e") { e.preventDefault(); setStyle({ ha: "center" }); return; }
    if (e.shiftKey && lk === "l") { e.preventDefault(); setStyle({ ha: "left" }); return; }
    if (e.shiftKey && lk === "r") { e.preventDefault(); setStyle({ ha: "right" }); return; }
    if (k === "Enter") { e.preventDefault(); startEdit(cellText(active.r, active.c)); return; }
    if (k === "\\") { e.preventDefault(); clearFormat(); return; }
    if (k === "/" || k === "?") { e.preventDefault(); shortcutsDialog(); return; }
    if (lk === "c" || lk === "x" || lk === "v") return; // the copy, cut and paste events do these
  }
  if (e.shiftKey && k === " " && !mod(e)) { e.preventDefault(); active = { r: R.r1, c: 0 }; cursor = { r: R.r2, c: g.cols - 1 }; selKind = "rows"; updateBars(); draw(); return; }
  if (e.altKey && (k === "ArrowDown")) { return; }
  switch (k) {
    case "ArrowUp": return move(-1, 0);
    case "ArrowDown": return move(1, 0);
    case "ArrowLeft": return move(0, -1);
    case "ArrowRight": return move(0, 1);
    case "PageDown": e.preventDefault(); return select(Math.min(g.rows - 1, (e.shiftKey ? cursor : active).r + Math.floor(viewH / (DEF_H * zoom)) - 2), (e.shiftKey ? cursor : active).c, e.shiftKey);
    case "PageUp": e.preventDefault(); return select(Math.max(0, (e.shiftKey ? cursor : active).r - Math.floor(viewH / (DEF_H * zoom)) + 2), (e.shiftKey ? cursor : active).c, e.shiftKey);
    case "Home": e.preventDefault(); return select((e.shiftKey ? cursor : active).r, 0, e.shiftKey);
    case "End": e.preventDefault(); { const b = eng.bounds(si); return select((e.shiftKey ? cursor : active).r, Math.max(0, b.c), e.shiftKey); }
    case "Tab": e.preventDefault(); return within(0, e.shiftKey ? -1 : 1);
    case "Enter": e.preventDefault(); return within(e.shiftKey ? -1 : 1, 0);
    case "F2": e.preventDefault(); return startEdit(cellText(active.r, active.c), { f2: true });
    case "Delete": case "Backspace": e.preventDefault(); return clearValues();
    case "Escape": copyBox = null; clip = null; draw(); return;
    case " ": if (styleAt(active.r, active.c).cb) { e.preventDefault(); toggleChecks(); return; } break;
  }
}
// Enter and Tab move inside a selected block, wrapping, like both apps.
function within(dr, dc) {
  const R = range();
  if (R.r1 === R.r2 && R.c1 === R.c2) { growTo(active.r + dr + 1, active.c + dc + 1); select(active.r + dr, active.c + dc); return; }
  let r = active.r + dr, c = active.c + dc;
  if (dc) { if (c > R.c2) { c = R.c1; r++; } if (c < R.c1) { c = R.c2; r--; } if (r > R.r2) r = R.r1; if (r < R.r1) r = R.r2; }
  if (dr) { if (r > R.r2) { r = R.r1; c++; } if (r < R.r1) { r = R.r2; c--; } if (c > R.c2) c = R.c1; if (c < R.c1) c = R.c2; }
  active = { r, c }; reveal(r, c); updateBars(); positionEditor(); draw();
}

// ---------------------------------------------------------------- editing a cell
function onEditorInput() {
  if (!edit) { startEdit(editor.value, { typed: true }); return; }
  edit.text = editor.value; fbar.value = editor.value; edit.point = null; sizeEditor(); suggest(editor); draw();
}
function startEdit(text, opts = {}) {
  if (edit) return;
  edit = { r: active.r, c: active.c, text, original: cellText(active.r, active.c), viaBar: !!opts.viaBar, f2: !!opts.f2, typed: !!opts.typed, point: null };
  editor.classList.remove("idle");
  editor.value = text; fbar.value = text;
  positionEditor(); sizeEditor();
  const s = styleAt(active.r, active.c);
  editor.style.font = font(s); editor.style.textAlign = s.ha === "right" ? "right" : s.ha === "center" ? "center" : "left";
  editor.style.fontWeight = s.b ? "bold" : ""; editor.style.color = s.fc || "";
  if (!opts.viaBar) { editor.focus(); editor.setSelectionRange(text.length, text.length); }
  if (opts.viaBar && !opts.keep) { fbar.focus(); }
  suggest(opts.viaBar ? fbar : editor);
  draw();
}
function positionEditor() {
  if (!editor || !doc) return;
  const r = edit ? edit.r : active.r, c = edit ? edit.c : active.c;
  const x = colX(c), y = rowY(r);
  editor.style.left = Math.max(hw(), x) + "px"; editor.style.top = Math.max(hh(), y) + "px";
  editor.style.minWidth = colW(c) + 1 + "px"; editor.style.minHeight = rowH(r) + 1 + "px";
  if (!edit) { editor.style.width = colW(c) + "px"; editor.style.height = rowH(r) + "px"; }
}
function sizeEditor() {
  if (!edit) return;
  const s = styleAt(edit.r, edit.c);
  ctx.font = font(s);
  const lines = edit.text.split("\n");
  const w = Math.max(...lines.map((l) => ctx.measureText(l).width)) + 14 * zoom;
  const maxW = viewW - (colX(edit.c)) - 8;
  editor.style.width = Math.max(colW(edit.c) + 1, Math.min(maxW, w)) + "px";
  editor.style.height = "auto";
  editor.style.height = Math.max(rowH(edit.r) + 1, editor.scrollHeight) + "px";
  fbar.style.height = lines.length > 1 ? Math.min(140, lines.length * 20 + 8) + "px" : "";
}
function endEdit() {
  edit = null; editor.classList.add("idle"); editor.value = ""; editor.style.cssText = ""; acEl.hidden = true; tipEl.hidden = true;
  positionEditor(); updateBars(); editor.focus(); draw();
}
function cancelEdit() { if (edit) endEdit(); }
function commit(dr = 0, dc = 0, all = false) {
  if (!edit) return;
  const { r, c, text } = edit;
  let t = text;
  if (t.startsWith("=")) t = autoClose(t);
  endEdit();
  if (all) { const R = range(); const rows = []; for (let rr = R.r1; rr <= R.r2; rr++) { const row = []; for (let cc = R.c1; cc <= R.c2; cc++) row.push(t.startsWith("=") ? F.shiftFormula(t, rr - r, cc - c) : t); rows.push(row); } enterTexts(rows, R.r1, R.c1); }
  else enterTexts([[t]], r, c);
  if (dr || dc) { if (range().r1 !== range().r2 || range().c1 !== range().c2) within(dr, dc); else { growTo(r + dr + 1, c + dc + 1); select(r + dr, c + dc); } }
}
// =SUM(A1:A3 becomes =SUM(A1:A3) when you press Enter, like both apps.
function autoClose(t) {
  let depth = 0, q = false;
  for (const ch of t) { if (ch === '"') q = !q; else if (!q) { if (ch === "(") depth++; else if (ch === ")") depth--; } }
  if (q) t += '"';
  return depth > 0 ? t + ")".repeat(depth) : t;
}
// Typed text becomes values, with formats for money, percents and dates.
function enterTexts(rows, r0, c0, styles) {
  if (!checkList(rows, r0, c0)) return;
  const updates = {};
  rows.forEach((row, i) => row.forEach((text, j) => {
    const r = r0 + i, c = c0 + j, old = cellAt(r, c);
    const p = F.parseInput(text);
    let s = old && old.s ? { ...old.s } : undefined;
    if (p.fmt && !(s && s.nf && (F.isDateFormat(s.nf) === F.isDateFormat(p.fmt)))) s = { ...(s || {}), nf: p.fmt };
    if (s && s.nf === "@" && typeof text === "string" && text !== "") { updates[A(r, c)] = { v: text, s }; return; }
    if (typeof p.v === "number" && s && s.nf && /%/.test(s.nf) && !p.fmt && !/%$/.test(String(text).trim()) && Math.abs(p.v) >= 1) p.v = p.v / 100; // typing 5 in a percent cell means 5%
    if (s && s.cb && typeof p.v !== "boolean" && p.v !== null) s = { ...s, cb: undefined };
    const cell = { ...(old || {}), v: p.v, s };
    delete cell.c;
    const extra = styles && styles[i] && styles[i][j];
    if (extra) cell.s = { ...(cell.s || {}), ...extra };
    updates[A(r, c)] = cell.v === null && !cell.s && !cell.n ? null : clean(cell);
  }));
  change(updates, styles ? "Paste" : "Typing");
  if (rows.some((row) => row.some((t) => String(t).includes("\n")))) fitRows({ r1: r0, r2: r0 + rows.length - 1 });
}
function clean(cell) {
  const out = {};
  if (cell.v !== null && cell.v !== undefined) out.v = cell.v;
  if (cell.s) { const s = {}; for (const [k, v] of Object.entries(cell.s)) if (v !== undefined && v !== null && v !== false && v !== 0 && v !== "") s[k] = v; if (Object.keys(s).length) out.s = s; }
  if (cell.n) out.n = cell.n;
  return Object.keys(out).length ? out : null;
}
// Is the caret where a cell reference could go next (after =, (, an operator or a comma)?
function pointable(text, pos) {
  if (!text.startsWith("=")) return false;
  const before = text.slice(0, pos).replace(/\s+$/, "");
  return /[=(,;+\-*/^&<>:]$/.test(before) || before === "=";
}
function editorKeys(e, el) {
  const k = e.key;
  if (!edit) { if (el === fbar && k === "Escape") { editor.focus(); } return; }
  if (!acEl.hidden) {
    const items = [...acEl.children];
    const on = items.findIndex((x) => x.classList.contains("on"));
    if (k === "ArrowDown" || k === "ArrowUp") { e.preventDefault(); const n = (on + (k === "ArrowDown" ? 1 : -1) + items.length) % items.length; items.forEach((x, i) => x.classList.toggle("on", i === n)); items[n].scrollIntoView({ block: "nearest" }); return; }
    if ((k === "Tab" || k === "Enter") && on >= 0) { e.preventDefault(); items[on].dispatchEvent(new MouseEvent("mousedown")); return; }
    if (k === "Escape") { e.preventDefault(); acEl.hidden = true; return; }
  }
  if (k === "Enter" && e.altKey) { e.preventDefault(); insertAt(el, "\n"); return; }
  if (k === "Enter") { e.preventDefault(); commit(e.shiftKey ? -1 : 1, 0, mod(e)); return; }
  if (k === "Tab") { e.preventDefault(); commit(0, e.shiftKey ? -1 : 1); return; }
  if (k === "Escape") { e.preventDefault(); cancelEdit(); return; }
  if (k === "F4") { e.preventDefault(); cycleAbsolute(el); return; }
  if (k === "F2") { e.preventDefault(); edit.f2 = !edit.f2; return; }
  if (k.startsWith("Arrow")) {
    const dr = k === "ArrowDown" ? 1 : k === "ArrowUp" ? -1 : 0, dc = k === "ArrowRight" ? 1 : k === "ArrowLeft" ? -1 : 0;
    // pointing: arrow keys pick a cell for the formula, like both apps
    if (!edit.f2 && el === editor && (edit.point || pointable(edit.text, el.selectionStart)) && el.selectionStart === el.selectionEnd) {
      e.preventDefault(); pointMove(dr, dc, e.shiftKey, mod(e)); return;
    }
    if (!edit.f2 && !edit.viaBar && !edit.text.startsWith("=") && edit.typed) { e.preventDefault(); commit(dr, dc); return; }
  }
}
function insertAt(el, s) {
  const a = el.selectionStart, b = el.selectionEnd;
  el.value = el.value.slice(0, a) + s + el.value.slice(b);
  el.setSelectionRange(a + s.length, a + s.length);
  edit.text = el.value; (el === editor ? fbar : editor).value = el.value; sizeEditor(); draw();
}
// Clicking or arrowing while typing a formula puts that cell's address in, replacing the one just put in.
function pointInsert(ref) {
  const el = edit.viaBar ? fbar : editor;
  let a, b;
  if (edit.point) { a = edit.point.start; b = edit.point.end; }
  else { a = el.selectionStart; b = el.selectionEnd; }
  const text = edit.text.slice(0, a) + ref + edit.text.slice(b);
  edit.text = text; el.value = text; (el === editor ? fbar : editor).value = text;
  const refs = F.refsIn(text); const idx = refs.findIndex((x) => x.start === a);
  edit.point = { start: a, end: a + ref.length, idx, anchor: edit.point ? edit.point.anchor : null };
  el.setSelectionRange(a + ref.length, a + ref.length);
  sizeEditor(); draw();
}
function rangeText(a, b) { return a.r === b.r && a.c === b.c ? A(b.r, b.c) : `${A(Math.min(a.r, b.r), Math.min(a.c, b.c))}:${A(Math.max(a.r, b.r), Math.max(a.c, b.c))}`; }
function pointMove(dr, dc, extend, far) {
  const p = edit.point;
  let at = p && p.cur ? p.cur : { r: edit.r, c: edit.c };
  const next = far ? jump(at.r, at.c, dr, dc) : { r: Math.max(0, at.r + dr), c: Math.max(0, at.c + dc) };
  const anchor = extend ? (p && p.anchor ? p.anchor : at) : next;
  const text = anchor.r === next.r && anchor.c === next.c ? A(next.r, next.c) : `${A(Math.min(anchor.r, next.r), Math.min(anchor.c, next.c))}:${A(Math.max(anchor.r, next.r), Math.max(anchor.c, next.c))}`;
  pointInsert(text);
  edit.point.cur = next; edit.point.anchor = anchor;
  reveal(next.r, next.c); positionEditor(); draw();
}
// F4 cycles A1 → $A$1 → A$1 → $A1 → A1 on the reference at the caret.
function cycleAbsolute(el) {
  const pos = el.selectionStart, refs = F.refsIn(edit.text);
  const ref = refs.find((x) => pos >= x.start && pos <= x.end) || refs.filter((x) => x.end <= pos).pop();
  if (!ref) return;
  const t = F.tokenize(ref.text)[0]; const p = F.refParts(t);
  const cycle = (x) => { const st = (x.ar ? 1 : 0) + (x.ac ? 2 : 0); const nx = { 0: 3, 3: 1, 1: 2, 2: 0 }[st]; return { ...x, ar: !!(nx & 1), ac: !!(nx & 2) }; };
  const q = { ...p, a: cycle(p.a), b: p.b ? cycle(p.b) : null };
  const repl = F.refText(q);
  edit.text = edit.text.slice(0, ref.start) + repl + edit.text.slice(ref.end);
  el.value = edit.text; (el === editor ? fbar : editor).value = edit.text;
  el.setSelectionRange(ref.start + repl.length, ref.start + repl.length); edit.point = null; draw();
}

// ---------------------------------------------------------------- function suggestions and argument hints
function suggest(el) {
  if (!edit || !edit.text.startsWith("=")) { acEl.hidden = true; tipEl.hidden = true; return; }
  const pos = el.selectionStart, before = edit.text.slice(0, pos);
  const m = /(^|[=(,;+\-*/^&<>:\s])([A-Za-z][A-Za-z0-9.]*)$/.exec(before);
  const box = (el === fbar ? fbar : editor).getBoundingClientRect(), wrap = wrapEl.getBoundingClientRect();
  if (m && !F.parseAddr(m[2])) {
    const word = m[2].toUpperCase();
    const list = F.FUNCTIONS.filter((f) => f.name.startsWith(word)).slice(0, 8);
    if (list.length && edit.text[pos] !== "(") {
      acEl.replaceChildren(...list.map((f, i) => h("div", { class: `aci ${i === 0 ? "on" : ""}`, role: "option", onmousedown: (e) => {
        e.preventDefault();
        const start = pos - m[2].length;
        const text = edit.text.slice(0, start) + f.name + "(" + edit.text.slice(pos);
        edit.text = text; el.value = text; (el === editor ? fbar : editor).value = text;
        el.setSelectionRange(start + f.name.length + 1, start + f.name.length + 1); edit.point = null;
        acEl.hidden = true; sizeEditor(); suggest(el); draw();
      } }, h("b", {}, f.name), h("small", {}, f.desc))));
      acEl.style.left = Math.max(0, box.left - wrap.left) + "px"; acEl.style.top = (el === fbar ? 0 : box.bottom - wrap.top) + 2 + "px";
      acEl.hidden = false; tipEl.hidden = true;
      return;
    }
  }
  acEl.hidden = true;
  // which function and argument the caret is in
  let depth = 0, q = false, argIdx = 0, fname = null;
  for (let i = pos - 1; i >= 0; i--) {
    const ch = edit.text[i];
    if (ch === '"') { q = !q; continue; }
    if (q) continue;
    if (ch === ")") depth++;
    else if (ch === "(") { if (depth === 0) { const nm = /([A-Za-z][A-Za-z0-9.]*)$/.exec(edit.text.slice(0, i)); fname = nm ? nm[1].toUpperCase() : null; break; } depth--; }
    else if ((ch === "," || ch === ";") && depth === 0) argIdx++;
  }
  const fn = fname && F.FUNCTIONS.find((f) => f.name === fname);
  if (!fn) { tipEl.hidden = true; return; }
  const args = fn.args.split(/,\s*/);
  const repeat = args.findIndex((a) => a.includes("…"));
  const cur = repeat >= 0 && argIdx >= repeat ? repeat - 1 + ((argIdx - repeat + 1) % 1) : argIdx;
  tipEl.replaceChildren(h("div", { class: "sig" }, h("b", {}, fn.name), "(", ...args.flatMap((a, i) => [i ? ", " : "", h("span", { class: i === Math.min(cur, args.length - 1) ? "arg on" : "arg" }, a)]), ")"),
    h("div", { class: "desc" }, fn.desc));
  tipEl.style.left = Math.max(0, box.left - wrap.left) + "px";
  tipEl.style.top = (el === fbar ? 0 : box.bottom - wrap.top) + 4 + "px";
  tipEl.hidden = false;
}
document.addEventListener("selectionchange", () => { if (edit && (document.activeElement === editor || document.activeElement === fbar)) suggest(document.activeElement); });

// ---------------------------------------------------------------- changing cells, with undo
// Every change is a list of cells before and after, so undo puts them back exactly.
function change(updates, label, opts = {}) {
  const s = sheet(), before = {};
  for (const key of Object.keys(updates)) before[key] = s.cells[key] ? JSON.parse(JSON.stringify(s.cells[key])) : null;
  apply(si, updates);
  undoStack.push({ kind: "cells", si, before, after: updates, label, sel: { active: { ...active }, cursor: { ...cursor } } });
  if (undoStack.length > 300) undoStack.shift();
  redoStack = [];
  if (!opts.quiet) afterChange();
}
function apply(sheetIndex, updates) {
  const s = doc.sheets[sheetIndex];
  for (const [key, cell] of Object.entries(updates)) { if (cell === null) delete s.cells[key]; else s.cells[key] = JSON.parse(JSON.stringify(cell)); }
  eng.touched(sheetIndex);
}
// Changes to rows, columns or sheets keep a copy of the whole workbook, which is simpler than working out the reverse.
function structural(label, work) {
  const before = JSON.stringify(doc);
  work();
  undoStack.push({ kind: "doc", before, after: JSON.stringify(doc), label, si, sel: { active: { ...active }, cursor: { ...cursor } } });
  redoStack = [];
  relayout(); sizeSpacer();
  afterChange();
}
function afterChange() {
  eng.recalc();
  updateBars(); renderTabs(); draw(); markDirty();
}
function undo() {
  if (edit) cancelEdit();
  const step = undoStack.pop();
  if (!step) return toast("Nothing to undo");
  if (step.kind === "cells") { si = step.si; apply(step.si, step.before); }
  else { replaceDoc(JSON.parse(step.before)); si = Math.min(step.si, doc.sheets.length - 1); }
  redoStack.push(step);
  active = step.sel.active; cursor = step.sel.cursor; reveal(active.r, active.c);
  if (step.kind === "doc") refilter();
  relayout(); sizeSpacer(); afterChange();
}
function redo() {
  if (edit) cancelEdit();
  const step = redoStack.pop();
  if (!step) return toast("Nothing to redo");
  if (step.kind === "cells") { si = step.si; apply(step.si, step.after); }
  else { replaceDoc(JSON.parse(step.after)); si = Math.min(step.si, doc.sheets.length - 1); }
  undoStack.push(step);
  if (step.kind === "doc") refilter();
  relayout(); sizeSpacer(); afterChange();
}
function replaceDoc(next) { doc.sheets = next.sheets; doc.active = next.active; eng = new F.Book(doc); }
function setValues(rows, r0, c0, fmt) {
  const updates = {};
  rows.forEach((row, i) => row.forEach((v, j) => { const old = cellAt(r0 + i, c0 + j); updates[A(r0 + i, c0 + j)] = clean({ ...(old || {}), v, s: fmt ? { ...((old && old.s) || {}), nf: fmt } : old && old.s }); }));
  change(updates, "Insert");
}
function eachInRange(R, fn) {
  const b = eng.bounds(si), s = sheet();
  const r2 = selKind === "cols" || selKind === "all" ? Math.min(R.r2, Math.max(b.r, s.rows - 1)) : R.r2, c2 = selKind === "rows" || selKind === "all" ? Math.min(R.c2, Math.max(b.c, s.cols - 1)) : R.c2;
  for (let r = R.r1; r <= r2; r++) for (let c = R.c1; c <= c2; c++) fn(r, c);
}
function clearValues() {
  const updates = {};
  eachInRange(range(), (r, c) => { const old = cellAt(r, c); if (old && old.v !== undefined) updates[A(r, c)] = clean({ ...old, v: null, c: undefined }); });
  if (Object.keys(updates).length) change(updates, "Delete");
}
function setStyle(patch) {
  const updates = {};
  eachInRange(range(), (r, c) => {
    const old = cellAt(r, c) || {};
    const s = { ...(old.s || {}), ...patch };
    for (const k of Object.keys(patch)) if (patch[k] === undefined) delete s[k];
    updates[A(r, c)] = clean({ ...old, s });
  });
  change(updates, "Format");
}
function toggle(key) { setStyle({ [key]: styleAt(active.r, active.c)[key] ? undefined : 1 }); }
function clearFormat() { const updates = {}; eachInRange(range(), (r, c) => { const old = cellAt(r, c); if (old && old.s) updates[A(r, c)] = clean({ ...old, s: undefined }); }); change(updates, "Clear formatting"); }
function decimals(step) {
  const v = eng.value(si, active.r, active.c), nf = styleAt(active.r, active.c).nf;
  let code = nf;
  if (!code || /^general$/i.test(code)) { const d = typeof v === "number" ? Math.min(10, (F.numText(v).split(".")[1] || "").length) : 0; code = d ? `0.${"0".repeat(d)}` : "0"; }
  const m = /0(\.0*)?(?![^"]*")/.exec(code.replace(/"[^"]*"/g, (q) => " ".repeat(q.length)));
  if (!m) return;
  const cur = m[1] ? m[1].length - 1 : 0, next = Math.max(0, Math.min(10, cur + step));
  const rebuilt = code.slice(0, m.index) + "0" + (next ? "." + "0".repeat(next) : "") + code.slice(m.index + m[0].length);
  setStyle({ nf: rebuilt.split(";").map((sec, i) => (i === 0 ? sec : sec.replace(/0(\.0*)?/, "0" + (next ? "." + "0".repeat(next) : "")))).join(";") });
}
function toggleChecks() {
  const R = range(), updates = {};
  const turnOn = eng.value(si, active.r, active.c) !== true;
  eachInRange(R, (r, c) => { const old = cellAt(r, c); if (old && old.s && old.s.cb) updates[A(r, c)] = clean({ ...old, v: turnOn }); });
  change(updates, "Checkbox");
}

// ---------------------------------------------------------------- mouse
function hit(e) {
  const rect = canvas.getBoundingClientRect(), x = e.clientX - rect.left, y = e.clientY - rect.top;
  const out = { x, y, inColHead: y < hh() && x >= hw(), inRowHead: x < hw() && y >= hh(), corner: x < hw() && y < hh() };
  out.r = rowFromY(Math.max(y, hh())); out.c = colFromX(Math.max(x, hw()));
  if (out.inColHead) { const right = colX(out.c) + colW(out.c), left = colX(out.c); if (Math.abs(x - right) < 5) out.colEdge = out.c; else if (Math.abs(x - left) < 5 && out.c > 0) out.colEdge = out.c - 1; }
  if (out.inRowHead) { const bottom = rowY(out.r) + rowH(out.r), top = rowY(out.r); if (Math.abs(y - bottom) < 4) out.rowEdge = out.r; else if (Math.abs(y - top) < 4 && out.r > 0) out.rowEdge = out.r - 1; }
  const R = range(), b = boxRect(R.r1, R.c1, R.r2, R.c2);
  out.handle = !edit && Math.abs(x - (b.x + b.w)) < 6 && Math.abs(y - (b.y + b.h)) < 6;
  return out;
}
function onHover(e) {
  if (drag) return;
  const p = hit(e);
  canvas.style.cursor = p.colEdge !== undefined ? "col-resize" : p.rowEdge !== undefined ? "row-resize" : p.handle ? "crosshair" : p.inColHead || p.inRowHead ? "default" : "cell";
  const cell = !p.inColHead && !p.inRowHead && !p.corner ? cellAt(p.r, p.c) : null;
  const key = cell && cell.n ? A(p.r, p.c) : null;
  if (key !== hoverNote) {
    hoverNote = key;
    if (key) { tipEl.replaceChildren(h("div", { class: "note" }, cell.n)); tipEl.style.left = colX(p.c) + colW(p.c) + 4 + "px"; tipEl.style.top = rowY(p.r) + "px"; tipEl.hidden = false; }
    else if (!edit) tipEl.hidden = true;
  }
}
function onMouseDown(e) {
  if (e.button === 2) return;
  menuEl.hidden = true;
  const p = hit(e);
  e.preventDefault();
  if (p.corner) { if (edit) commit(); selectAll(); editor.focus(); return; }
  if (p.colEdge !== undefined) { drag = { kind: "colsize", c: p.colEdge, x0: e.clientX, w0: colW(p.colEdge) / zoom }; trackDrag(); return; }
  if (p.rowEdge !== undefined) { drag = { kind: "rowsize", r: p.rowEdge, y0: e.clientY, h0: rowH(p.rowEdge) / zoom }; trackDrag(); return; }
  // pointing at cells while typing a formula: the cell's address goes in (shift-click makes a range)
  if (edit && edit.text.startsWith("=") && !p.inColHead && !p.inRowHead) {
    const el = edit.viaBar ? fbar : editor;
    if (edit.point || pointable(edit.text, el.selectionStart)) {
      const anchor = e.shiftKey && edit.point && edit.point.anchor ? edit.point.anchor : { r: p.r, c: p.c };
      pointInsert(rangeText(anchor, p));
      edit.point.anchor = anchor; edit.point.cur = { r: p.r, c: p.c };
      drag = { kind: "point" }; trackDrag(); el.focus();
      return;
    }
  }
  if (edit) commit();
  if (p.inColHead && p.colEdge === undefined) { const fi = filterIconAt(p.c); if (fi && p.x >= fi.x - 2 && p.x <= fi.x + fi.s + 2) { filterMenu(p.c, e); return; } }
  if (p.inColHead) { if (e.shiftKey) { cursor = { r: geometry().rows - 1, c: p.c }; } else { active = { r: 0, c: p.c }; cursor = { r: geometry().rows - 1, c: p.c }; } selKind = "cols"; drag = { kind: "cols" }; updateBars(); draw(); trackDrag(); editor.focus(); return; }
  if (p.inRowHead) { if (e.shiftKey) { cursor = { r: p.r, c: geometry().cols - 1 }; } else { active = { r: p.r, c: 0 }; cursor = { r: p.r, c: geometry().cols - 1 }; } selKind = "rows"; drag = { kind: "rows" }; updateBars(); draw(); trackDrag(); editor.focus(); return; }
  if (p.handle) { drag = { kind: "fill", from: range() }; trackDrag(); return; }
  const st = styleAt(p.r, p.c);
  if (st.cb && !e.shiftKey) {
    const b = boxRect(p.r, p.c, p.r, p.c), s = 16 * zoom;
    if (Math.abs(p.x - (b.x + b.w / 2)) < s && Math.abs(p.y - (b.y + b.h / 2)) < s) { select(p.r, p.c); toggleChecks(); editor.focus(); return; }
  }
  if (!e.shiftKey && listAt(p.r, p.c)) { const b = boxRect(p.r, p.c, p.r, p.c); if (p.x > b.x + b.w - 20 * zoom) { select(p.r, p.c); openListMenu(p.r, p.c); editor.focus(); return; } }
  select(p.r, p.c, e.shiftKey);
  drag = { kind: "cells" }; trackDrag();
  editor.focus();
}
function trackDrag() {
  let scrollTimer = null, last = null;
  const moveAt = (e) => {
    if (!drag) return;
    last = e;
    const p = hit(e);
    if (drag.kind === "colsize") { const w = Math.max(4, Math.round(drag.w0 + (e.clientX - drag.x0) / zoom)); sheet().colW[drag.c] = w; relayout(); sizeSpacer(); positionEditor(); draw(); return; }
    if (drag.kind === "rowsize") { const hgt = Math.max(4, Math.round(drag.h0 + (e.clientY - drag.y0) / zoom)); sheet().rowH[drag.r] = hgt; relayout(); sizeSpacer(); positionEditor(); draw(); return; }
    if (drag.kind === "cells") { cursor = { r: p.r, c: p.c }; selKind = "cells"; }
    if (drag.kind === "cols") cursor = { r: geometry().rows - 1, c: p.c };
    if (drag.kind === "rows") cursor = { r: p.r, c: geometry().cols - 1 };
    if (drag.kind === "point" && edit && edit.point) { const a = edit.point.anchor; pointInsert(rangeText(a, p)); edit.point.anchor = a; edit.point.cur = { r: p.r, c: p.c }; }
    if (drag.kind === "fill") {
      const f = drag.from;
      const down = p.r > f.r2 ? p.r - f.r2 : 0, up = p.r < f.r1 ? f.r1 - p.r : 0, right = p.c > f.c2 ? p.c - f.c2 : 0, left = p.c < f.c1 ? f.c1 - p.c : 0;
      const vert = Math.max(down, up), horiz = Math.max(right, left);
      if (!vert && !horiz) drag.to = null;
      else if (vert >= horiz) drag.to = down ? { r1: f.r1, c1: f.c1, r2: p.r, c2: f.c2, dir: "down" } : { r1: p.r, c1: f.c1, r2: f.r2, c2: f.c2, dir: "up" };
      else drag.to = right ? { r1: f.r1, c1: f.c1, r2: f.r2, c2: p.c, dir: "right" } : { r1: f.r1, c1: p.c, r2: f.r2, c2: f.c2, dir: "left" };
    }
    updateBars(); draw();
  };
  const autoscroll = () => {
    if (!last || !drag) return;
    const rect = scroller.getBoundingClientRect();
    let dx = 0, dy = 0;
    if (last.clientX > rect.right - 20) dx = 20; else if (last.clientX < rect.left + hw() + 10) dx = -20;
    if (last.clientY > rect.bottom - 20) dy = 20; else if (last.clientY < rect.top + hh() + 10) dy = -20;
    if (dx || dy) { scroller.scrollLeft += dx; scroller.scrollTop += dy; moveAt(last); }
  };
  const up = () => {
    window.removeEventListener("mousemove", moveAt); window.removeEventListener("mouseup", up); clearInterval(scrollTimer);
    const d = drag; drag = null;
    if (!d) return; // a release that was already handled (say the button came up outside the window)
    if (d.kind === "colsize" || d.kind === "rowsize") { const s = sheet(); const key = d.kind === "colsize" ? "colW" : "rowH"; const before = JSON.stringify(doc); const val = s[key][d.kind === "colsize" ? d.c : d.r];
      // the size was changed live; record it as one undo step
      const range0 = range();
      if (d.kind === "colsize" && (selKind === "cols" || selKind === "all") && d.c >= range0.c1 && d.c <= range0.c2) for (let c = range0.c1; c <= range0.c2; c++) s.colW[c] = val;
      if (d.kind === "rowsize" && (selKind === "rows" || selKind === "all") && d.r >= range0.r1 && d.r <= range0.r2) for (let r = range0.r1; r <= range0.r2; r++) s.rowH[r] = val;
      undoStack.push({ kind: "doc", before: before, after: JSON.stringify(doc), label: "Resize", si, sel: { active: { ...active }, cursor: { ...cursor } } }); relayout(); sizeSpacer(); markDirty(); draw(); }
    if (d.kind === "fill" && d.to) fillTo(d.from, d.to);
    if (d.kind === "point" && edit) (edit.viaBar ? fbar : editor).focus();
    draw();
  };
  window.addEventListener("mousemove", moveAt); window.addEventListener("mouseup", up);
  scrollTimer = setInterval(autoscroll, 50);
}
function onDblClick(e) {
  const p = hit(e);
  if (p.colEdge !== undefined) { autoFitCols(p.colEdge); return; }
  if (p.rowEdge !== undefined) { fitRows({ r1: p.rowEdge, r2: p.rowEdge }, true); return; }
  if (p.handle) { fillDownToData(); return; }
  if (p.inColHead) { renameColumn(p.c); return; }
  if (p.inRowHead || p.corner) return;
  select(p.r, p.c);
  startEdit(cellText(p.r, p.c), { f2: true });
}
function autoFitCols(c0) {
  const R = range(), cols = (selKind === "cols" || selKind === "all") && c0 >= R.c1 && c0 <= R.c2 ? range2(R.c1, R.c2) : [c0];
  structural("Fit column width", () => {
    const b = eng.bounds(si);
    for (const c of cols) {
      let w = 30;
      for (let r = 0; r <= Math.min(b.r, 5000); r++) { const d = display(r, c); if (!d.text) continue; ctx.font = font(d.s); w = Math.max(w, ctx.measureText(d.text).width / zoom + 10); }
      sheet().colW[c] = Math.min(600, Math.ceil(w));
    }
  });
}
// Rows with wrapped text grow to fit, like both apps.
function fitRows(R, force) {
  const s = sheet();
  let changed = false;
  const byRow = new Map();
  for (const key of Object.keys(s.cells)) { const p = F.parseAddr(key); if (p && p.r >= R.r1 && p.r <= Math.min(R.r2, R.r1 + 500)) { if (!byRow.has(p.r)) byRow.set(p.r, []); byRow.get(p.r).push([key, p]); } }
  for (let r = R.r1; r <= Math.min(R.r2, R.r1 + 500); r++) {
    let need = DEF_H;
    for (const [key, p] of byRow.get(r) || []) {
      const st = s.cells[key].s || {}; const d = display(p.r, p.c); if (!d.text || (!st.wr && !d.text.includes("\n"))) { if (st.fs) need = Math.max(need, Math.round(st.fs * 1.33 + 8)); continue; }
      ctx.font = font(st); const size = (st.fs || 10) * 1.33;
      const width = colW(p.c) / zoom - 6; let lines = 0;
      for (const para of d.text.split("\n")) { let line = ""; lines++; for (const word of para.split(/(?<=\s)/)) { if (ctx.measureText(line + word).width / zoom > width && line) { lines++; line = word; } else line += word; } }
      need = Math.max(need, Math.ceil(lines * size * 1.25 + 6));
    }
    if (force || (s.rowH[r] ?? DEF_H) < need || (s.rowH[r] && s.rowH[r] > need && force)) { if (need === DEF_H) delete s.rowH[r]; else s.rowH[r] = need; changed = true; }
  }
  if (changed) { relayout(); sizeSpacer(); markDirty(); draw(); }
}

// ---------------------------------------------------------------- fill handle
const MONTHS_L = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
const DAYS_L = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
// Values for extending a list: numbers keep their step, dates their gap, "Item 1" counts on, month and day names continue.
function seriesFor(vals, n) {
  const k = vals.length;
  const nums = vals.map((x) => x.v);
  if (k && nums.every((v) => typeof v === "number")) {
    const step = k >= 2 ? (nums[k - 1] - nums[0]) / (k - 1) : (vals[0].date ? 1 : 0);
    const steady = k < 2 || nums.every((v, i) => !i || Math.abs(v - nums[i - 1] - step) < 1e-9);
    if (steady && (k >= 2 || vals[0].date)) return Array.from({ length: n }, (_, i) => nums[k - 1] + step * (i + 1));
  }
  const names = (list, short) => {
    const idx = nums.map((v) => (typeof v === "string" ? list.findIndex((x) => (short ? x.slice(0, 3) : x) === v.toLowerCase()) : -1));
    if (idx.every((i) => i >= 0)) { const step = k >= 2 ? idx[1] - idx[0] : 1; const cap = (s) => (nums[0] === nums[0].toUpperCase() ? s.toUpperCase() : s[0].toUpperCase() + s.slice(1)); return Array.from({ length: n }, (_, i) => cap((short ? list[((idx[k - 1] + step * (i + 1)) % 12 + 12) % 12].slice(0, 3) : list[((idx[k - 1] + step * (i + 1)) % list.length + list.length) % list.length]))); }
    return null;
  };
  const named = names(MONTHS_L, false) || names(MONTHS_L.map((m) => m), true) || names(DAYS_L, false) || names(DAYS_L, true);
  if (named) return named;
  if (k && nums.every((v) => typeof v === "string" && /\d+$/.test(v) && !v.startsWith("="))) {
    const parts = nums.map((v) => /^(.*?)(\d+)$/.exec(v));
    if (parts.every((p) => p[1] === parts[0][1])) {
      const ns = parts.map((p) => Number(p[2])); const step = k >= 2 ? ns[1] - ns[0] : 1; const width = parts[k - 1][2].length;
      return Array.from({ length: n }, (_, i) => parts[0][1] + String(ns[k - 1] + step * (i + 1)).padStart(parts[0][2].startsWith("0") ? width : 0, "0"));
    }
  }
  return null;
}
function fillTo(from, to) {
  const updates = {}, s = sheet();
  const vertical = to.dir === "down" || to.dir === "up";
  const lanes = vertical ? range2(from.c1, from.c2) : range2(from.r1, from.r2);
  for (const lane of lanes) {
    const src = vertical ? range2(from.r1, from.r2).map((r) => ({ r, c: lane })) : range2(from.c1, from.c2).map((c) => ({ r: lane, c }));
    const targets = vertical ? (to.dir === "down" ? range2(from.r2 + 1, to.r2) : range2(to.r1, from.r1 - 1).reverse()).map((r) => ({ r, c: lane }))
      : (to.dir === "right" ? range2(from.c2 + 1, to.c2) : range2(to.c1, from.c1 - 1).reverse()).map((c) => ({ r: lane, c }));
    const ordered = to.dir === "up" || to.dir === "left" ? [...src].reverse() : src;
    const vals = ordered.map((p) => { const cell = s.cells[A(p.r, p.c)]; return { v: cell ? cell.v : null, date: !!(cell && cell.s && cell.s.nf && F.isDateFormat(cell.s.nf)) }; });
    const series = !vals.some((x) => typeof x.v === "string" && x.v.startsWith("=")) ? seriesFor(vals, targets.length) : null;
    targets.forEach((t, i) => {
      const from0 = ordered[i % ordered.length], cell = s.cells[A(from0.r, from0.c)];
      if (!cell) { updates[A(t.r, t.c)] = null; return; }
      let v = cell.v;
      if (series && (vals.length > 1 || typeof v !== "number" || vals[0].date)) v = series[i];
      else if (typeof v === "string" && v.startsWith("=")) v = F.shiftFormula(v, t.r - from0.r, t.c - from0.c);
      updates[A(t.r, t.c)] = clean({ v, s: cell.s, n: undefined });
    });
  }
  change(updates, "Fill");
  active = { r: Math.min(from.r1, to.r1), c: Math.min(from.c1, to.c1) }; cursor = { r: Math.max(from.r2, to.r2), c: Math.max(from.c2, to.c2) };
  updateBars(); draw();
}
function fillDir(dir) {
  const R = range();
  if (dir === "down") { if (R.r1 === R.r2) { if (R.r1 === 0) return; fillTo({ ...R, r1: R.r1 - 1, r2: R.r1 - 1 }, { ...R, r1: R.r1 - 1, dir: "down" }); } else fillTo({ ...R, r2: R.r1 }, { ...R, dir: "down" }); }
  else { if (R.c1 === R.c2) { if (R.c1 === 0) return; fillTo({ ...R, c1: R.c1 - 1, c2: R.c1 - 1 }, { ...R, c1: R.c1 - 1, dir: "right" }); } else fillTo({ ...R, c2: R.c1 }, { ...R, dir: "right" }); }
}
// Double-clicking the fill handle fills down as far as the column next to it goes, like both apps.
function fillDownToData() {
  const R = range();
  const filled = (r, c) => { const v = eng.value(si, r, c); return v !== null && v !== ""; };
  let end = R.r2;
  for (const c of [R.c1 - 1, R.c2 + 1]) { if (c < 0) continue; let r = R.r2; while (filled(r + 1, c)) r++; end = Math.max(end, r); }
  if (end > R.r2) fillTo(R, { ...R, r2: end, dir: "down" });
}

// ---------------------------------------------------------------- copy and paste (plain text works with Excel and Google Sheets)
function shownText(r, c) { const d = display(r, c); return d.check !== undefined ? String(!!d.check).toUpperCase() : d.text; }
function copy(data, cut) {
  const R = range(), b = eng.bounds(si);
  const r2 = Math.min(R.r2, Math.max(R.r1, b.r)), c2 = Math.min(R.c2, Math.max(R.c1, b.c));
  const rows = [], cellsCopy = [];
  for (let r = R.r1; r <= r2; r++) { const row = [], crow = []; for (let c = R.c1; c <= c2; c++) { row.push(shownText(r, c)); const cell = cellAt(r, c); crow.push(cell ? JSON.parse(JSON.stringify(cell)) : null); } rows.push(row); cellsCopy.push(crow); }
  const tsv = rows.map((row) => row.map((t) => (/[\t\n"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t)).join("\t")).join("\n");
  const esc = (t) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const html = `<table>${rows.map((row, i) => `<tr>${row.map((t, j) => { const st = (cellsCopy[i][j] || {}).s || {}; const css = [st.b && "font-weight:bold", st.i && "font-style:italic", st.fc && `color:${st.fc}`, st.bg && `background:${st.bg}`].filter(Boolean).join(";"); return `<td${css ? ` style="${css}"` : ""}>${esc(t)}</td>`; }).join("")}</tr>`).join("")}</table>`;
  data.setData("text/plain", tsv); data.setData("text/html", html);
  clip = { si, r1: R.r1, c1: R.c1, rows: cellsCopy, text: tsv, cut };
  copyBox = { si, r1: R.r1, c1: R.c1, r2, c2 };
  toast(cut ? "Cut. Paste where you want it." : `Copied ${rows.length * rows[0].length} cell${rows.length * rows[0].length === 1 ? "" : "s"}`);
  draw();
}
function parseTSV(text) {
  const rows = []; let row = [], cur = "", q = false;
  text = text.replace(/\r\n?/g, "\n").replace(/\n$/, "");
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; continue; }
    if (ch === '"' && cur === "") { q = true; continue; }
    if (ch === "\t") { row.push(cur); cur = ""; continue; }
    if (ch === "\n") { row.push(cur); rows.push(row); row = []; cur = ""; continue; }
    cur += ch;
  }
  row.push(cur); rows.push(row);
  return rows;
}
function paste(data, valuesOnly) {
  const text = data.getData("text/plain");
  const R = range();
  if (clip && text === clip.text) {
    const h0 = clip.rows.length, w0 = clip.rows[0].length;
    // a selection that's a multiple of what was copied gets it repeated, like both apps
    const reps = (R.r2 - R.r1 + 1) % h0 === 0 && (R.c2 - R.c1 + 1) % w0 === 0 ? [(R.r2 - R.r1 + 1) / h0, (R.c2 - R.c1 + 1) / w0] : [1, 1];
    const updates = {};
    if (clip.cut && clip.si === si) clip.rows.forEach((row, i) => row.forEach((_, j) => { updates[A(clip.r1 + i, clip.c1 + j)] = null; }));
    for (let ri = 0; ri < reps[0]; ri++) for (let ci = 0; ci < reps[1]; ci++) clip.rows.forEach((row, i) => row.forEach((cell, j) => {
      const r = R.r1 + ri * h0 + i, c = R.c1 + ci * w0 + j;
      if (!cell) { updates[A(r, c)] = null; return; }
      let v = cell.v;
      if (valuesOnly) { const val = eng.value(clip.si, clip.r1 + i, clip.c1 + j); v = F.isErr(val) ? val.code : val; updates[A(r, c)] = clean({ v, s: (cellAt(r, c) || {}).s || (cell.s && cell.s.nf ? { nf: cell.s.nf } : undefined) }); return; }
      if (typeof v === "string" && v.startsWith("=") && !clip.cut) v = F.shiftFormula(v, r - (clip.r1 + i), c - (clip.c1 + j));
      updates[A(r, c)] = clean({ ...cell, v, c: undefined });
    }));
    if (clip.cut && clip.si !== si) { const other = {}; clip.rows.forEach((row, i) => row.forEach((_, j) => { other[A(clip.r1 + i, clip.c1 + j)] = null; })); apply(clip.si, other); }
    growTo(R.r1 + h0 * reps[0], R.c1 + w0 * reps[1]);
    change(updates, "Paste");
    active = { r: R.r1, c: R.c1 }; cursor = { r: R.r1 + h0 * reps[0] - 1, c: R.c1 + w0 * reps[1] - 1 };
    if (clip.cut) { clip = null; copyBox = null; }
    updateBars(); draw();
    return;
  }
  if (!text) return;
  const rows = parseTSV(text);
  growTo(R.r1 + rows.length, R.c1 + Math.max(...rows.map((r) => r.length)));
  const html = data.getData("text/html");
  const styles = html && !valuesOnly ? htmlStyles(html) : null;
  enterTexts(rows, R.r1, R.c1, styles || []);
  active = { r: R.r1, c: R.c1 }; cursor = { r: R.r1 + rows.length - 1, c: R.c1 + Math.max(...rows.map((r) => r.length)) - 1 };
  copyBox = null; updateBars(); draw();
}
// Bold, italic and colours from a table copied out of Excel or Google Sheets.
function htmlStyles(html) {
  try {
    const parsed = new DOMParser().parseFromString(html, "text/html");
    const rows = [...parsed.querySelectorAll("tr")];
    if (!rows.length) return null;
    return rows.map((tr) => [...tr.querySelectorAll("td,th")].map((td) => {
      const css = td.getAttribute("style") || "", st = {};
      const inner = td.querySelector("b,strong");
      if (/font-weight:\s*(bold|[6-9]00)/i.test(css) || inner) st.b = 1;
      if (/font-style:\s*italic/i.test(css) || td.querySelector("i,em")) st.i = 1;
      const color = /(?:^|;)\s*color:\s*([^;]+)/i.exec(css), bg = /background(?:-color)?:\s*([^;]+)/i.exec(css);
      if (color && !/^(#000|#000000|black|rgb\(0,\s*0,\s*0\)|windowtext)$/i.test(color[1].trim())) st.fc = color[1].trim();
      if (bg && !/^(#fff|#ffffff|white|transparent|rgb\(255,\s*255,\s*255\))$/i.test(bg[1].trim())) st.bg = bg[1].trim();
      return Object.keys(st).length ? st : null;
    }));
  } catch (e) { return null; }
}

// ---------------------------------------------------------------- rows, columns and sheets
function rewriteAll(fn) {
  doc.sheets.forEach((s) => { for (const cell of Object.values(s.cells)) if (typeof cell.v === "string" && cell.v.startsWith("=")) cell.v = fn(cell.v, s.name); });
}
function shiftCells(s, axis, at, n) {
  const next = {};
  for (const [key, cell] of Object.entries(s.cells)) {
    const p = F.parseAddr(key); let { r, c } = p;
    const v = axis === "r" ? r : c;
    if (n < 0 && v >= at && v < at - n) continue;
    if (v >= at) { if (axis === "r") r += n; else c += n; }
    next[A(r, c)] = cell;
  }
  s.cells = next;
  const sizes = axis === "r" ? s.rowH : s.colW, moved = {};
  for (const [k, w] of Object.entries(sizes)) { const i = Number(k); if (n < 0 && i >= at && i < at - n) continue; moved[i >= at ? i + n : i] = w; }
  if (axis === "r") { s.rowH = moved; s.rows = Math.max(1, s.rows + n); } else { s.colW = moved; s.cols = Math.max(1, s.cols + n); }
  shiftExtras(s, axis, at, n);
  if (axis === "c" && s.names) {   // column names move with their columns
    const names = {};
    for (const [k, v] of Object.entries(s.names)) { const i = Number(k); if (n < 0 && i >= at && i < at - n) continue; names[i >= at ? i + n : i] = v; }
    s.names = names;
  }
}
// ---------------------------------------------------------------- column names
// A column can have a name (Company, Position…) shown in the header bar instead of its letter, so a header
// doesn't take up row 1. Formulas still use the letters (A1, B2).
function renameColumn(c) {
  const s = sheet(), now = (s.names || {})[c] || "";
  promptDialog(`Name for column ${F.colName(c)} (leave empty to show the letter)`, now, (v) => {
    const name = v.trim().replace(/\s+/g, " ").slice(0, 100);
    if (name === now) return;
    structural("Rename column", () => { s.names = { ...(s.names || {}) }; if (name) s.names[c] = name; else delete s.names[c]; });
  });
}
function row1Texts() {
  const out = {};
  for (let c = 0; c < sheet().cols; c++) { const t = display(0, c).text; if (t !== "" && t != null) out[c] = String(t).replace(/\s+/g, " ").trim().slice(0, 100); }
  return out;
}
// Row 1's text becomes the column names, and row 1 goes, so the data starts at row 1. Formulas are adjusted.
function useRow1AsNames() {
  const names = row1Texts();
  if (!Object.keys(names).length) { toast("Row 1 is empty, so there's nothing to use as column names"); return; }
  const name = sheet().name;
  structural("Use row 1 as column names", () => {
    const s = sheet();
    s.names = { ...(s.names || {}), ...names };
    rewriteAll((f, home) => F.structFormula(f, home, name, "r", 0, -1));
    shiftCells(s, "r", 0, -1);
    if (s.freeze && s.freeze.r) s.freeze.r = Math.max(0, s.freeze.r - 1);
  });
  toast("Row 1 is now the column names. Undo with Ctrl+Z if that wasn't right.");
}
// The other way: the names go back into a new row 1.
function namesToRow1() {
  const s = sheet(), names = s.names || {};
  if (!Object.keys(names).length) return;
  const name = s.name;
  structural("Put column names in row 1", () => {
    rewriteAll((f, home) => F.structFormula(f, home, name, "r", 0, 1));
    shiftCells(s, "r", 0, 1);
    for (const [c, v] of Object.entries(names)) s.cells[A(0, Number(c))] = { v, s: { b: 1 } };
    s.names = {};
  });
}
function columnNameItems() {
  const s = sheet(), has = Object.keys(s.names || {}).length > 0, R = range();
  return [
    { label: `Rename column ${F.colName(R.c1)}…`, action: () => renameColumn(R.c1) },
    { label: "Use row 1 as column names", action: useRow1AsNames },
    { label: "Put column names back in row 1", disabled: !has, action: namesToRow1 },
    { label: "Clear all column names", disabled: !has, action: () => structural("Clear column names", () => { sheet().names = {}; }) },
  ];
}
function insertRows(at, n) { const name = sheet().name; structural(n > 0 ? "Insert rows" : "Delete rows", () => { rewriteAll((f, home) => F.structFormula(f, home, name, "r", at, n)); shiftCells(sheet(), "r", at, n); }); }
function insertCols(at, n) { const name = sheet().name; structural(n > 0 ? "Insert columns" : "Delete columns", () => { rewriteAll((f, home) => F.structFormula(f, home, name, "c", at, n)); shiftCells(sheet(), "c", at, n); }); }
function deleteSelectedRows() { const R = range(); insertRows(R.r1, -(R.r2 - R.r1 + 1)); select(R.r1, active.c); }
function deleteSelectedCols() { const R = range(); insertCols(R.c1, -(R.c2 - R.c1 + 1)); select(active.r, R.c1); }
function uniqueSheetName(base) { let n = 1, name = base; while (doc.sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) name = `${base.replace(/ \d+$/, "")} ${++n}`; return name; }
function addSheet() {
  structural("Add sheet", () => {
    let n = doc.sheets.length + 1; while (doc.sheets.some((s) => s.name === `Sheet${n}`)) n++;
    doc.sheets.splice(si + 1, 0, { name: `Sheet${n}`, rows: 1000, cols: 26, cells: {}, colW: {}, rowH: {}, freeze: { r: 0, c: 0 }, names: {} });
    si += 1;
  });
  switchSheet(si);
}
function switchSheet(i) {
  if (edit && !(edit.text.startsWith("=") && edit.point !== undefined)) commit();
  si = Math.max(0, Math.min(i, doc.sheets.length - 1)); doc.active = si;
  active = { r: 0, c: 0 }; cursor = { r: 0, c: 0 }; selKind = "cells"; copyBox = copyBox && copyBox.si === si ? copyBox : copyBox;
  scroller.scrollLeft = 0; scroller.scrollTop = 0; scrollX = scrollY = 0;
  clearCharts(); refilter();
  relayout(); sizeSpacer(); renderTabs(); updateBars(); positionEditor(); draw(); editor.focus();
}
function renameSheet(i) {
  promptDialog("Rename sheet", doc.sheets[i].name, (name) => {
    name = name.trim();
    if (!name || /[\[\]:*?/\\]/.test(name) || /^'|'$/.test(name)) return toast("A sheet name can't be empty or use [ ] : * ? / \\", true);
    if (doc.sheets.some((s, j) => j !== i && s.name.toLowerCase() === name.toLowerCase())) return toast("Another sheet already has that name", true);
    const old = doc.sheets[i].name;
    structural("Rename sheet", () => { rewriteAll((f) => F.renameSheetIn(f, old, name)); doc.sheets[i].name = name; });
  });
}
function deleteSheet(i) {
  if (doc.sheets.filter((s) => !s.hidden).length <= 1) return toast("A workbook needs at least one sheet", true);
  confirmDialog(`Delete the sheet “${doc.sheets[i].name}”? Formulas that use it will show #REF!. You can undo this.`, () => {
    const name = doc.sheets[i].name;
    structural("Delete sheet", () => { doc.sheets.splice(i, 1); rewriteAll((f) => F.dropSheetIn(f, name)); });
    eng = new F.Book(doc); switchSheet(Math.min(i, doc.sheets.length - 1)); afterChange();
  });
}
function duplicateSheet(i) {
  const copyName = uniqueSheetName(`Copy of ${doc.sheets[i].name}`.slice(0, 100));
  structural("Duplicate sheet", () => { const c = JSON.parse(JSON.stringify(doc.sheets[i])); c.name = copyName; doc.sheets.splice(i + 1, 0, c); });
  switchSheet(i + 1);
}
function moveSheet(i, d) {
  const j = i + d; if (j < 0 || j >= doc.sheets.length) return;
  structural("Move sheet", () => { const [s] = doc.sheets.splice(i, 1); doc.sheets.splice(j, 0, s); });
  eng = new F.Book(doc); afterChange(); switchSheet(j);
}
function renderTabs() {
  if (!tabsEl || !doc) return;
  tabsEl.replaceChildren(...doc.sheets.flatMap((s, i) => s.hidden ? [] : h("button", { class: `stab ${i === si ? "on" : ""}`, style: s.tab ? `--tab:${s.tab}` : "",
    onclick: () => { if (i !== si) switchSheet(i); }, ondblclick: () => renameSheet(i),
    oncontextmenu: (e) => { e.preventDefault(); openMenu(e.currentTarget, sheetMenu(i), false, e); } }, s.name, h("span", { class: "caret", onclick: (e) => { e.stopPropagation(); openMenu(e.currentTarget, sheetMenu(i)); } }, "▾"))));
}
function sheetMenu(i) {
  return [
    { label: "Rename", action: () => renameSheet(i) }, { label: "Duplicate", action: () => duplicateSheet(i) },
    { label: "Delete", danger: true, action: () => deleteSheet(i) },
    { label: "Tab colour", sub: () => ["#ea4335", "#fbbc04", "#34a853", "#4285f4", "#a142f4", "#ff6d01", null].map((c) => ({ label: c ? "■ " + c : "No colour", color: c, action: () => structural("Tab colour", () => { if (c) doc.sheets[i].tab = c; else delete doc.sheets[i].tab; }) })) },
    { label: "Hide sheet", action: () => { if (doc.sheets.filter((s) => !s.hidden).length <= 1) return toast("At least one sheet has to show", true); structural("Hide sheet", () => { doc.sheets[i].hidden = true; }); switchSheet(doc.sheets.findIndex((s) => !s.hidden)); } },
    { sep: true }, { label: "Move left", disabled: i === 0, action: () => moveSheet(i, -1) }, { label: "Move right", disabled: i === doc.sheets.length - 1, action: () => moveSheet(i, 1) },
  ];
}

// ---------------------------------------------------------------- sorting, freezing, data tools
// The block around the selected cell, for sorting when only one cell is picked.
function currentRegion() {
  const R = range();
  if (R.r1 !== R.r2 || R.c1 !== R.c2) return R;
  const b = eng.bounds(si);
  let r1 = active.r, r2 = active.r, c1 = active.c, c2 = active.c;
  const filled = (r, c) => { const v = eng.value(si, r, c); return v !== null && v !== ""; };
  let grew = true;
  while (grew) {
    grew = false;
    if (r1 > 0 && range2(c1, c2).some((c) => filled(r1 - 1, c))) { r1--; grew = true; }
    if (r2 < b.r && range2(c1, c2).some((c) => filled(r2 + 1, c))) { r2++; grew = true; }
    if (c1 > 0 && range2(r1, r2).some((r) => filled(r, c1 - 1))) { c1--; grew = true; }
    if (c2 < b.c && range2(r1, r2).some((r) => filled(r, c2 + 1))) { c2++; grew = true; }
  }
  return { r1, c1, r2, c2 };
}
function sortBy(asc, R0) {
  let R = R0 || currentRegion();
  const s = sheet(), fr = (s.freeze || {}).r || 0;
  // a header row (frozen, or text above numbers) stays put
  if (!R0) {
    if (R.r1 < fr) R = { ...R, r1: fr };
    else { const top = range2(R.c1, R.c2).map((c) => eng.value(si, R.r1, c)), below = range2(R.c1, R.c2).map((c) => eng.value(si, R.r1 + 1, c)); if (top.every((v) => typeof v === "string" || v === null) && below.some((v) => typeof v === "number") && R.r2 > R.r1) R = { ...R, r1: R.r1 + 1 }; }
  }
  if (R.r2 <= R.r1) return toast("Nothing to sort");
  const col = active.c >= R.c1 && active.c <= R.c2 ? active.c : R.c1;
  const rows = range2(R.r1, R.r2).map((r) => ({ r, key: eng.value(si, r, col), cells: range2(R.c1, R.c2).map((c) => cellAt(r, c)) }));
  rows.sort((a, b) => { const ea = a.key === null || a.key === "", eb = b.key === null || b.key === ""; if (ea || eb) return ea - eb; const c = F.compare(a.key, b.key); return asc ? c : -c; });
  const updates = {};
  rows.forEach((row, i) => row.cells.forEach((cell, j) => {
    const r = R.r1 + i, c = R.c1 + j;
    if (!cell) { updates[A(r, c)] = null; return; }
    const v = typeof cell.v === "string" && cell.v.startsWith("=") ? F.shiftFormula(cell.v, r - row.r, 0) : cell.v;
    updates[A(r, c)] = clean({ ...cell, v, c: undefined });
  }));
  change(updates, "Sort");
  toast(`Sorted rows ${R.r1 + 1} to ${R.r2 + 1} by column ${F.colName(col)}, ${asc ? "A to Z" : "Z to A"}`);
}
function toggleFreeze() {
  const s = sheet(), f = s.freeze || { r: 0, c: 0 };
  const next = f.r || f.c ? { r: 0, c: 0 } : { r: active.r, c: active.c };
  if (!next.r && !next.c && !(f.r || f.c)) return toast("Pick a cell below and to the right of what should stay in view, then freeze");
  structural("Freeze", () => { s.freeze = next; });
  toast(next.r || next.c ? `Froze ${next.r ? `${next.r} row${next.r > 1 ? "s" : ""}` : ""}${next.r && next.c ? " and " : ""}${next.c ? `${next.c} column${next.c > 1 ? "s" : ""}` : ""}` : "Unfroze");
}
function setFreeze(r, c) { structural("Freeze", () => { sheet().freeze = { r: Math.min(r, 50), c: Math.min(c, 26) }; }); }
function removeDuplicates() {
  const R = currentRegion(), seen = new Set(), keep = [];
  for (let r = R.r1; r <= R.r2; r++) { const key = JSON.stringify(range2(R.c1, R.c2).map((c) => { const v = eng.value(si, r, c); return typeof v === "string" ? v.toLowerCase().trim() : v; })); if (!seen.has(key)) { seen.add(key); keep.push(r); } }
  const removed = R.r2 - R.r1 + 1 - keep.length;
  if (!removed) return toast("No duplicate rows found");
  const updates = {};
  keep.forEach((r, i) => range2(R.c1, R.c2).forEach((c) => { const cell = cellAt(r, c); updates[A(R.r1 + i, c)] = cell ? clean({ ...cell, c: undefined }) : null; }));
  for (let r = R.r1 + keep.length; r <= R.r2; r++) for (const c of range2(R.c1, R.c2)) updates[A(r, c)] = null;
  change(updates, "Remove duplicates");
  toast(`Removed ${removed} duplicate row${removed === 1 ? "" : "s"}; ${keep.length} left`);
}
function trimSpaces() {
  const updates = {}; let n = 0;
  eachInRange(range(), (r, c) => { const cell = cellAt(r, c); if (cell && typeof cell.v === "string" && !cell.v.startsWith("=")) { const t = cell.v.replace(/\s+/g, " ").trim(); if (t !== cell.v) { updates[A(r, c)] = clean({ ...cell, v: t }); n++; } } });
  if (n) change(updates, "Trim"); toast(n ? `Trimmed ${n} cell${n === 1 ? "" : "s"}` : "No extra spaces found");
}
function splitColumns() {
  promptDialog("Split text into columns at", ",", (sep) => {
    if (!sep) return;
    const R = range(), rows = [];
    const sepReal = sep === "\\t" ? "\t" : sep;
    for (let r = R.r1; r <= R.r2; r++) { const v = eng.value(si, r, R.c1); rows.push(typeof v === "string" ? v.split(sepReal).map((x) => x.trim()) : [v === null ? "" : String(v)]); }
    enterTexts(rows, R.r1, R.c1);
  });
}

// ---------------------------------------------------------------- menus
function fileMenu() {
  return [
    { label: "New spreadsheet", action: () => newBook(null) },
    { label: "Open…", shortcut: "", action: () => { location.hash = ""; } },
    { label: "Make a copy", action: async () => { await flushSave(); act("duplicate", { id: meta.id }).then((o) => { location.hash = `book/${o.id}`; toast("This is the copy"); }).catch(fail); } },
    { label: "Import a file…", action: () => { const f = h("input", { type: "file", accept: ".xlsx,.xlsm,.csv,.tsv,.txt", onchange: (e) => importFile(e.target.files[0]) }); f.click(); } },
    { sep: true },
    { label: "Download", sub: () => [
      { label: "Microsoft Excel (.xlsx)", action: async () => { await flushSave(); download(meta.id, "xlsx"); } },
      { label: "Comma-separated values (.csv, this sheet)", action: async () => { await flushSave(); download(meta.id, "csv", si); } },
      { label: "Tab-separated values (.tsv, this sheet)", action: async () => { await flushSave(); download(meta.id, "tsv", si); } }] },
    { label: "Version history", action: versionsDialog },
    { sep: true },
    { label: "Rename", action: () => { titleEl.focus(); titleEl.select(); } },
    { label: "Delete", danger: true, action: () => confirmDialog(`Delete “${meta.name}”? Its version history goes too.`, async () => { await act("delete", { id: meta.id }).catch(fail); meta = null; location.hash = ""; }) },
  ];
}
function editMenu() {
  return [
    { label: "Undo", shortcut: `${KEY}Z`, disabled: !undoStack.length, action: undo }, { label: "Redo", shortcut: `${KEY}${isMac ? "⇧Z" : "Y"}`, disabled: !redoStack.length, action: redo },
    { sep: true },
    { label: "Cut", shortcut: `${KEY}X`, action: () => document.execCommand("cut") }, { label: "Copy", shortcut: `${KEY}C`, action: () => document.execCommand("copy") },
    { label: "Paste", shortcut: `${KEY}V`, action: () => pasteFromMenu(false) }, { label: "Paste values only", shortcut: `${KEY}⇧V`, action: () => pasteFromMenu(true) },
    { sep: true },
    { label: "Find and replace", shortcut: `${KEY}H`, action: () => findDialog(true) },
    { label: "Fill down", shortcut: `${KEY}D`, action: () => fillDir("down") }, { label: "Fill right", shortcut: `${KEY}R`, action: () => fillDir("right") },
    { label: "Delete values", shortcut: "Delete", action: clearValues },
    { label: "Delete rows " + rowsLabel(), action: deleteSelectedRows }, { label: "Delete columns " + colsLabel(), action: deleteSelectedCols },
  ];
}
async function pasteFromMenu(values) {
  try {
    const text = await navigator.clipboard.readText();
    paste({ getData: (t) => (t === "text/plain" ? text : "") }, values);
  } catch (e) { toast(`Use ${KEY}V to paste${values ? ` (${KEY}⇧V for values only)` : ""}; this browser doesn't let menus read the clipboard.`); }
}
const rowsLabel = () => { const R = range(); return R.r1 === R.r2 ? `${R.r1 + 1}` : `${R.r1 + 1}–${R.r2 + 1}`; };
const colsLabel = () => { const R = range(); return R.c1 === R.c2 ? F.colName(R.c1) : `${F.colName(R.c1)}–${F.colName(R.c2)}`; };
function viewMenu() {
  const f = sheet().freeze || { r: 0, c: 0 };
  return [
    { label: "Freeze", sub: () => [
      { label: "No rows", check: !f.r, action: () => setFreeze(0, f.c) }, { label: "1 row", check: f.r === 1, action: () => setFreeze(1, f.c) }, { label: "2 rows", check: f.r === 2, action: () => setFreeze(2, f.c) },
      { label: `Up to row ${active.r + 1}`, action: () => setFreeze(active.r + 1, f.c) }, { sep: true },
      { label: "No columns", check: !f.c, action: () => setFreeze(f.r, 0) }, { label: "1 column", check: f.c === 1, action: () => setFreeze(f.r, 1) }, { label: "2 columns", check: f.c === 2, action: () => setFreeze(f.r, 2) },
      { label: `Up to column ${F.colName(active.c)}`, action: () => setFreeze(f.r, active.c + 1) }] },
    { label: "Gridlines", check: gridlines, action: () => { gridlines = !gridlines; try { localStorage.setItem("sheets-grid", gridlines ? "1" : "0"); } catch (e) { /* fine */ } draw(); } },
    { label: "Formulas instead of results", shortcut: `${KEY}\``, check: showFormulas, action: () => { showFormulas = !showFormulas; draw(); } },
    { sep: true },
    { label: "Zoom", sub: () => [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2].map((z) => ({ label: `${Math.round(z * 100)}%`, check: z === zoom, action: () => { zoom = z; tb.zoom.value = String(z); try { localStorage.setItem("sheets-zoom", z); } catch (e) { /* fine */ } onResize(); } })) },
  ];
}
function insertMenu() {
  const R = range(), nr = R.r2 - R.r1 + 1, nc = R.c2 - R.c1 + 1;
  return [
    { label: `${nr > 1 ? nr + " rows" : "Row"} above`, action: () => insertRows(R.r1, nr) }, { label: `${nr > 1 ? nr + " rows" : "Row"} below`, action: () => insertRows(R.r2 + 1, nr) },
    { label: `${nc > 1 ? nc + " columns" : "Column"} left`, action: () => insertCols(R.c1, nc) }, { label: `${nc > 1 ? nc + " columns" : "Column"} right`, action: () => insertCols(R.c2 + 1, nc) },
    { sep: true },
    { label: "Sheet", action: addSheet },
    { label: "Chart", action: () => chartDialog() },
    { label: "Dropdown", action: listDialog },
    { label: "Function", sub: () => sumMenu() },
    { label: "Checkbox", action: () => { const updates = {}; eachInRange(range(), (r, c) => { const old = cellAt(r, c) || {}; updates[A(r, c)] = clean({ ...old, v: typeof old.v === "boolean" ? old.v : false, s: { ...(old.s || {}), cb: 1, ha: "center" } }); }); change(updates, "Checkbox"); } },
    { label: "Note", action: noteDialog },
    { label: "Date (today)", shortcut: `${KEY};`, action: () => { const n = new Date(); setValues([[F.dateSerial(n.getFullYear(), n.getMonth() + 1, n.getDate())]], active.r, active.c, "m/d/yyyy"); } },
  ];
}
function formatMenu() {
  return [
    { label: "Number", sub: numberMenu },
    { label: "Text", sub: () => [{ label: "Bold", shortcut: `${KEY}B`, action: () => toggle("b") }, { label: "Italic", shortcut: `${KEY}I`, action: () => toggle("i") }, { label: "Underline", shortcut: `${KEY}U`, action: () => toggle("u") }, { label: "Strikethrough", shortcut: `${KEY}5`, action: () => toggle("st") }] },
    { label: "Alignment", sub: () => [{ label: "Left", action: () => setStyle({ ha: "left" }) }, { label: "Centre", action: () => setStyle({ ha: "center" }) }, { label: "Right", action: () => setStyle({ ha: "right" }) }, { sep: true },
      { label: "Top", action: () => setStyle({ va: "top" }) }, { label: "Middle", action: () => setStyle({ va: "middle" }) }, { label: "Bottom", action: () => setStyle({ va: undefined }) }] },
    { label: "Wrapping", sub: () => [{ label: "Overflow", check: !styleAt(active.r, active.c).wr && !styleAt(active.r, active.c).fit, action: () => setStyle({ wr: undefined, fit: undefined }) },
      { label: "Wrap", check: !!styleAt(active.r, active.c).wr, action: () => { setStyle({ wr: 1, fit: undefined }); fitRows(range()); } },
      { label: "Shrink to fit", check: !!styleAt(active.r, active.c).fit, action: () => setStyle({ fit: 1, wr: undefined }) }] },
    { label: "Merge cells", sub: () => [{ label: "Merge all", action: () => mergeCells("all") }, { label: "Merge horizontally", action: () => mergeCells("rows") }, { label: "Merge vertically", action: () => mergeCells("cols") }, { sep: true },
      { label: "Unmerge", disabled: !(sheet().merges || []).some((m) => { const R = range(); return m.r1 <= R.r2 && m.r2 >= R.r1 && m.c1 <= R.c2 && m.c2 >= R.c1; }), action: unmerge }] },
    { label: "Font size", sub: () => [8, 9, 10, 11, 12, 14, 18, 24, 36].map((n) => ({ label: String(n), check: (styleAt(active.r, active.c).fs || 10) === n, action: () => setStyle({ fs: n === 10 ? undefined : n }) })) },
    { label: "Row height…", action: () => promptDialog("Row height in pixels", String(Math.round(rowH(active.r) / zoom)), (v) => { const n = Number(v); if (!(n >= 4 && n <= 2000)) return; const R = range(); structural("Row height", () => { for (let r = R.r1; r <= Math.min(R.r2, R.r1 + 5000); r++) sheet().rowH[r] = n; }); }) },
    { label: "Column width…", action: () => promptDialog("Column width in pixels", String(Math.round(colW(active.c) / zoom)), (v) => { const n = Number(v); if (!(n >= 4 && n <= 2000)) return; const R = range(); structural("Column width", () => { for (let c = R.c1; c <= R.c2; c++) sheet().colW[c] = n; }); }) },
    { sep: true }, { label: "Clear formatting", shortcut: `${KEY}\\`, action: clearFormat },
  ];
}
function numberMenu() {
  const nf = styleAt(active.r, active.c).nf;
  const item = (label, code, sample) => ({ label, hint: sample, check: (nf || undefined) === code, action: () => setStyle({ nf: code }) });
  return [
    item("Automatic", undefined, "1000.12"), item("Plain text", "@", "1000.12"), { sep: true },
    item("Number", "#,##0.00", "1,000.12"), item("Percent", "0.00%", "10.12%"), item("Scientific", "0.00E+00", "1.01E+03"), { sep: true },
    item("Financial", "#,##0.00;(#,##0.00)", "(1,000.12)"), item("Currency", "$#,##0.00", "$1,000.12"), item("Currency rounded", "$#,##0", "$1,000"), { sep: true },
    item("Date", "m/d/yyyy", "9/26/2008"), item("Long date", "mmmm d, yyyy", "September 26, 2008"), item("Time", "h:mm:ss AM/PM", "3:59:00 PM"), item("Date time", "m/d/yyyy h:mm", "9/26/2008 15:59"), item("Duration", "[h]:mm:ss", "24:01:00"),
    { sep: true }, { label: "Custom number format…", action: () => promptDialog("Format code (like #,##0.00 or mmm d, yyyy)", nf || "", (code) => setStyle({ nf: code.trim() || undefined })) },
  ];
}
function dataMenu() {
  return [
    { label: "Sort range A to Z", action: () => sortBy(true) }, { label: "Sort range Z to A", action: () => sortBy(false) },
    { label: "Sort selected cells only…", sub: () => [{ label: "A to Z", action: () => sortBy(true, range()) }, { label: "Z to A", action: () => sortBy(false, range()) }] },
    { sep: true },
    { label: filterOn() ? "Turn off filter" : "Create a filter", action: toggleFilter },
    { label: listAt(active.r, active.c) ? "Edit dropdown…" : "Dropdown list…", action: listDialog }, { sep: true },
    { label: "Remove duplicate rows", action: removeDuplicates }, { label: "Trim extra spaces", action: trimSpaces }, { label: "Split text into columns…", action: splitColumns },
    { sep: true }, { label: "Column names", sub: columnNameItems },
  ];
}
function helpMenu() { return [{ label: "Keyboard shortcuts", shortcut: `${KEY}/`, action: shortcutsDialog }, { label: "Every function", action: functionsDialog }]; }
function sumMenu() {
  const pick = (name) => () => insertFunction(name);
  return [{ label: "SUM", action: pick("SUM") }, { label: "AVERAGE", action: pick("AVERAGE") }, { label: "COUNT", action: pick("COUNT") }, { label: "MAX", action: pick("MAX") }, { label: "MIN", action: pick("MIN") },
    { sep: true }, ...["Math", "Statistics", "Logical", "Lookup", "Text", "Date", "Finance", "Info"].map((cat) => ({ label: cat, sub: () => F.FUNCTIONS.filter((f) => f.cat === cat).map((f) => ({ label: f.name, hint: f.desc.length > 46 ? f.desc.slice(0, 45) + "…" : f.desc, action: pick(f.name) })) })),
    { sep: true }, { label: "Every function…", action: functionsDialog }];
}
// =SUM( with the numbers above or to the left already picked, like both apps' AutoSum.
function insertFunction(name) {
  if (edit) { const el = edit.viaBar ? fbar : editor; insertAt(el, `${edit.text.startsWith("=") ? "" : "="}${name}(`); return; }
  const R = range();
  let arg = "";
  if (["SUM", "AVERAGE", "COUNT", "MAX", "MIN"].includes(name)) {
    if (R.r1 !== R.r2 || R.c1 !== R.c2) arg = `${A(R.r1, R.c1)}:${A(R.r2, R.c2)}`;
    else {
      let r = active.r - 1; while (r >= 0 && typeof eng.value(si, r, active.c) === "number") r--;
      if (r < active.r - 1) arg = `${A(r + 1, active.c)}:${A(active.r - 1, active.c)}`;
      else { let c = active.c - 1; while (c >= 0 && typeof eng.value(si, active.r, c) === "number") c--; if (c < active.c - 1) arg = `${A(active.r, c + 1)}:${A(active.r, active.c - 1)}`; }
    }
    if (R.r1 !== R.r2 || R.c1 !== R.c2) { const target = { r: R.r2 + 1, c: R.c1 }; if (R.c1 === R.c2) { growTo(target.r + 1, target.c + 1); select(target.r, target.c); enterTexts([[`=${name}(${arg})`]], target.r, target.c); return; } }
  }
  startEdit(`=${name}(${arg}`);
}
function colorMenu(key) {
  const colors = ["#000000", "#434343", "#666666", "#999999", "#b7b7b7", "#cccccc", "#d9d9d9", "#efefef", "#f3f3f3", "#ffffff",
    "#980000", "#ff0000", "#ff9900", "#ffff00", "#00ff00", "#00ffff", "#4a86e8", "#0000ff", "#9900ff", "#ff00ff",
    "#e6b8af", "#f4cccc", "#fce5cd", "#fff2cc", "#d9ead3", "#d0e0e3", "#c9daf8", "#cfe2f3", "#d9d2e9", "#ead1dc",
    "#dd7e6b", "#ea9999", "#f9cb9c", "#ffe599", "#b6d7a8", "#a2c4c9", "#a4c2f4", "#9fc5e8", "#b4a7d6", "#d5a6bd",
    "#a61c00", "#cc0000", "#e69138", "#f1c232", "#6aa84f", "#45818e", "#3c78d8", "#3d85c6", "#674ea7", "#a64d79"];
  return [{ palette: colors, onpick: (c) => setStyle({ [key]: c }) }, { label: key === "bg" ? "No fill" : "Automatic", action: () => setStyle({ [key]: undefined }) }];
}
function borderMenu() {
  const set = (mode) => {
    const R = range(), updates = {};
    eachInRange(R, (r, c) => {
      const old = cellAt(r, c) || {}; const bd = mode === "none" ? undefined : { ...((old.s || {}).bd || {}) };
      if (mode === "all") Object.assign(bd, { t: 1, r: 1, b: 1, l: 1 });
      if (mode === "outer") { if (r === R.r1) bd.t = 1; if (r === R.r2) bd.b = 1; if (c === R.c1) bd.l = 1; if (c === R.c2) bd.r = 1; }
      if (mode === "bottom" && r === R.r2) bd.b = 1;
      if (mode === "top" && r === R.r1) bd.t = 1;
      updates[A(r, c)] = clean({ ...old, s: { ...(old.s || {}), bd: bd && Object.keys(bd).length ? bd : undefined } });
    });
    change(updates, "Borders");
  };
  return [{ label: "All borders", action: () => set("all") }, { label: "Outer border", action: () => set("outer") }, { label: "Top border", action: () => set("top") }, { label: "Bottom border", action: () => set("bottom") }, { label: "No borders", action: () => set("none") }];
}
function onContext(e) {
  e.preventDefault();
  const p = hit(e), R = range();
  const inside = p.r >= R.r1 && p.r <= R.r2 && p.c >= R.c1 && p.c <= R.c2;
  if (!inside && !p.inColHead && !p.inRowHead) select(p.r, p.c);
  if (p.inColHead && !(p.c >= R.c1 && p.c <= R.c2 && selKind === "cols")) { active = { r: 0, c: p.c }; cursor = { r: geometry().rows - 1, c: p.c }; selKind = "cols"; updateBars(); draw(); }
  if (p.inRowHead && !(p.r >= R.r1 && p.r <= R.r2 && selKind === "rows")) { active = { r: p.r, c: 0 }; cursor = { r: p.r, c: geometry().cols - 1 }; selKind = "rows"; updateBars(); draw(); }
  const Q = range(), nr = Q.r2 - Q.r1 + 1, nc = Q.c2 - Q.c1 + 1;
  const items = [
    { label: "Cut", shortcut: `${KEY}X`, action: () => document.execCommand("cut") }, { label: "Copy", shortcut: `${KEY}C`, action: () => document.execCommand("copy") },
    { label: "Paste", shortcut: `${KEY}V`, action: () => pasteFromMenu(false) }, { label: "Paste values only", shortcut: `${KEY}⇧V`, action: () => pasteFromMenu(true) }, { sep: true },
  ];
  if (selKind !== "cols") items.push({ label: `Insert ${nr > 1 ? nr + " rows" : "1 row"} above`, action: () => insertRows(Q.r1, nr) }, { label: `Insert ${nr > 1 ? nr + " rows" : "1 row"} below`, action: () => insertRows(Q.r2 + 1, nr) });
  if (selKind !== "rows") items.push({ label: `Insert ${nc > 1 ? nc + " columns" : "1 column"} left`, action: () => insertCols(Q.c1, nc) }, { label: `Insert ${nc > 1 ? nc + " columns" : "1 column"} right`, action: () => insertCols(Q.c2 + 1, nc) });
  items.push({ sep: true });
  if (selKind !== "cols") items.push({ label: `Delete row${nr > 1 ? "s " : " "}${rowsLabel()}`, action: deleteSelectedRows });
  if (selKind !== "rows") items.push({ label: `Delete column${nc > 1 ? "s " : " "}${colsLabel()}`, action: deleteSelectedCols });
  items.push({ label: "Clear values", action: clearValues }, { sep: true },
    { label: "Sort range A to Z", action: () => sortBy(true) }, { label: "Sort range Z to A", action: () => sortBy(false) }, { sep: true },
    { label: cellAt(active.r, active.c) && cellAt(active.r, active.c).n ? "Edit note" : "Insert note", action: noteDialog },
    { label: listAt(active.r, active.c) ? "Edit dropdown…" : "Dropdown list…", action: listDialog });
  if (selKind === "cols") items.push({ label: "Fit to data", action: () => autoFitCols(Q.c1) }, { sep: true }, ...columnNameItems());
  openMenu(null, items, false, e);
}
// One menu at a time, with keyboard support and submenus.
let menuStack = [];
// The outside-click listener runs in the capture phase, so an open menu closes before the button that was pressed
// acts. Pressing the button whose menu was just closed leaves it closed (a toggle).
let menuAnchor = null, menuFromBar = false, lastClosed = { anchor: null, at: 0 };
function openMenu(anchor, items, fromBar, ev) {
  if (anchor && lastClosed.anchor === anchor && performance.now() - lastClosed.at < 500) { lastClosed.anchor = null; return; }
  if (!menuEl.hidden) closeMenu();
  menuAnchor = anchor; menuFromBar = !!fromBar;
  menuEl.replaceChildren(); menuEl.hidden = false; menuStack = [];
  buildMenu(menuEl, items);
  const rect = anchor ? anchor.getBoundingClientRect() : { left: ev.clientX, bottom: ev.clientY, top: ev.clientY };
  menuEl.style.left = "0px"; menuEl.style.top = "0px";
  const w = menuEl.offsetWidth, hgt = menuEl.offsetHeight;
  menuEl.style.left = Math.max(4, Math.min(rect.left, window.innerWidth - w - 4)) + "px";
  menuEl.style.top = (rect.bottom + hgt + 4 > window.innerHeight ? Math.max(4, rect.top - hgt) : rect.bottom + 2) + "px";
  document.addEventListener("mousedown", closeMenuOutside, true);
}
function buildMenu(el, items) {
  for (const it of items) {
    if (it.sep) { el.append(h("div", { class: "msep" })); continue; }
    if (it.palette) { el.append(h("div", { class: "palette" }, it.palette.map((c) => h("button", { class: "sw", style: `background:${c}`, title: c, "aria-label": c, onmousedown: (e) => { e.preventDefault(); closeMenu(); it.onpick(c); editor.focus(); } })))); continue; }
    const row = h("button", { class: `mi ${it.danger ? "danger" : ""}`, disabled: it.disabled, onmousedown: (e) => {
      e.preventDefault();
      if (it.sub) return;
      closeMenu(); it.action && it.action(); if (editor && !dialog.open && document.activeElement !== titleEl) editor.focus();
    } }, h("span", { class: "mcheck" }, it.check ? "✓" : ""), it.color ? h("i", { class: "mcolor", style: `background:${it.color}` }) : null, h("span", { class: "mlabel" }, it.label),
      it.hint ? h("small", { class: "mhint" }, it.hint) : null, it.shortcut ? h("small", { class: "mkey" }, it.shortcut) : null, it.sub ? h("span", { class: "marrow" }, "▸") : null);
    if (it.sub) {
      row.addEventListener("mouseenter", () => {
        document.querySelectorAll(".menu.sub").forEach((m) => { if (Number(m.dataset.depth) >= Number(el.dataset.depth || 0) + 1) m.remove(); });
        const sub = h("div", { class: "menu sub", "data-depth": Number(el.dataset.depth || 0) + 1 });
        buildMenu(sub, it.sub()); document.body.append(sub);
        const r = row.getBoundingClientRect();
        sub.style.left = (r.right + sub.offsetWidth > window.innerWidth ? r.left - sub.offsetWidth : r.right) + "px";
        sub.style.top = Math.max(4, Math.min(r.top, window.innerHeight - sub.offsetHeight - 4)) + "px";
      });
    } else row.addEventListener("mouseenter", () => document.querySelectorAll(".menu.sub").forEach((m) => { if (Number(m.dataset.depth) > Number(el.dataset.depth || 0)) m.remove(); }));
    el.append(row);
  }
}
function closeMenu() {
  menuEl.hidden = true; document.querySelectorAll(".menu.sub").forEach((m) => m.remove());
  document.removeEventListener("mousedown", closeMenuOutside, true);
  menuAnchor = null; menuFromBar = false;
}
function closeMenuOutside(e) {
  if (e.target.closest && e.target.closest(".menu")) return;
  const pressed = e.target.closest ? e.target.closest("button") : null;
  lastClosed = { anchor: pressed && pressed === menuAnchor ? pressed : null, at: performance.now() };
  closeMenu();
}
document.addEventListener("keydown", (e) => { if (e.key === "Escape" && !menuEl.hidden) { e.stopPropagation(); closeMenu(); } }, true);

// ---------------------------------------------------------------- dialogs
function openDialog(...kids) { dialog.replaceChildren(h("form", { class: "dlg", method: "dialog" }, ...kids)); if (!dialog.open) dialog.showModal(); }
dialog.addEventListener("close", () => { if (editor && doc) editor.focus(); });
function promptDialog(title, value, then) {
  const input = h("input", { value: value || "", "aria-label": title });
  openDialog(h("h2", {}, title), input, h("div", { class: "btns" }, h("button", { class: "ghost", value: "cancel" }, "Cancel"),
    h("button", { class: "primary", onclick: (e) => { e.preventDefault(); dialog.close(); then(input.value); } }, "OK")));
  input.focus(); input.select();
}
function confirmDialog(text, then) {
  openDialog(h("h2", {}, "Are you sure?"), h("p", {}, text), h("div", { class: "btns" }, h("button", { class: "ghost", value: "cancel" }, "Cancel"),
    h("button", { class: "primary danger", onclick: (e) => { e.preventDefault(); dialog.close(); then(); } }, "Delete")));
}
function noteDialog() {
  const cell = cellAt(active.r, active.c) || {};
  const box = h("textarea", { rows: 4, value: cell.n || "", "aria-label": "Note", placeholder: "A note for this cell" });
  openDialog(h("h2", {}, `Note on ${A(active.r, active.c)}`), box, h("div", { class: "btns" },
    cell.n ? h("button", { class: "ghost bad", onclick: (e) => { e.preventDefault(); dialog.close(); change({ [A(active.r, active.c)]: clean({ ...cell, n: undefined }) }, "Note"); } }, "Remove") : null, h("span", { class: "grow" }),
    h("button", { class: "ghost", value: "cancel" }, "Cancel"), h("button", { class: "primary", onclick: (e) => { e.preventDefault(); dialog.close(); change({ [A(active.r, active.c)]: clean({ ...cell, n: box.value.trim() || undefined }) }, "Note"); } }, "Save")));
  box.focus();
}
function findDialog(replace) {
  const what = h("input", { "aria-label": "Find", placeholder: "Find" }), with_ = h("input", { "aria-label": "Replace with", placeholder: "Replace with" });
  const matchCase = h("input", { type: "checkbox" }), whole = h("input", { type: "checkbox" }), inFormulas = h("input", { type: "checkbox" }), everywhere = h("input", { type: "checkbox" });
  const status = h("p", { class: "hint" });
  let last = null;
  const test = (t) => { const a = matchCase.checked ? t : t.toLowerCase(), b = matchCase.checked ? what.value : what.value.toLowerCase(); return whole.checked ? a === b : a.includes(b); };
  const hits = () => {
    const out = [];
    const sheets = everywhere.checked ? doc.sheets.map((_, i) => i) : [si];
    for (const s of sheets) {
      const keys = Object.keys(doc.sheets[s].cells).map((k) => ({ k, p: F.parseAddr(k) })).sort((x, y) => x.p.r - y.p.r || x.p.c - y.p.c);
      for (const { k, p } of keys) {
        const cell = doc.sheets[s].cells[k];
        const text = inFormulas.checked && typeof cell.v === "string" ? cell.v : (() => { const v = eng.value(s, p.r, p.c); return F.format(v, (cell.s || {}).nf).text; })();
        if (what.value && text && test(text)) out.push({ s, r: p.r, c: p.c });
      }
    }
    return out;
  };
  const next = (dir) => {
    const list = hits();
    if (!list.length) { status.textContent = "Nothing found"; return; }
    let i = last ? list.findIndex((x) => x.s === last.s && (x.r > last.r || (x.r === last.r && x.c > last.c))) : 0;
    if (dir < 0) { i = last ? list.length - 1 - [...list].reverse().findIndex((x) => x.s === last.s && (x.r < last.r || (x.r === last.r && x.c < last.c))) : list.length - 1; if (i >= list.length) i = list.length - 1; }
    if (i < 0) i = dir > 0 ? 0 : list.length - 1;
    last = list[i]; if (last.s !== si) switchSheet(last.s); select(last.r, last.c);
    status.textContent = `${i + 1} of ${list.length}`;
  };
  const doReplace = (all) => {
    const list = all ? hits() : last ? [last] : [];
    if (!list.length) { status.textContent = "Nothing to replace"; return; }
    const re = new RegExp(what.value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), matchCase.checked ? "g" : "gi");
    const bySheet = new Map();
    for (const x of list) {
      const cell = doc.sheets[x.s].cells[A(x.r, x.c)]; if (!cell || typeof cell.v !== "string") continue;
      if (cell.v.startsWith("=") && !inFormulas.checked) continue;
      const v = whole.checked ? with_.value : cell.v.replace(re, () => with_.value);
      if (!bySheet.has(x.s)) bySheet.set(x.s, {});
      bySheet.get(x.s)[A(x.r, x.c)] = clean({ ...cell, v: cell.v.startsWith("=") ? v : F.parseInput(v).v, c: undefined });
    }
    let n = 0;
    for (const [s, updates] of bySheet) { const keep = si; si = s; change(updates, "Replace"); si = keep; n += Object.keys(updates).length; }
    status.textContent = `Replaced ${n} cell${n === 1 ? "" : "s"}`;
    if (!all) next(1);
  };
  openDialog(h("h2", {}, replace ? "Find and replace" : "Find"), what, replace ? with_ : null,
    h("div", { class: "checks" }, h("label", { class: "check" }, matchCase, "Match case"), h("label", { class: "check" }, whole, "Match the whole cell"),
      h("label", { class: "check" }, inFormulas, "Also look inside formulas"), h("label", { class: "check" }, everywhere, "Every sheet")),
    status,
    h("div", { class: "btns" }, h("button", { type: "button", class: "ghost", onclick: () => next(-1) }, "Previous"), h("button", { type: "button", class: "primary", onclick: () => next(1) }, "Find next"),
      replace ? h("button", { type: "button", class: "ghost", onclick: () => doReplace(false) }, "Replace") : null, replace ? h("button", { type: "button", class: "ghost", onclick: () => doReplace(true) }, "Replace all") : null,
      h("button", { class: "ghost", value: "cancel" }, "Done")));
  what.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); next(e.shiftKey ? -1 : 1); } });
  what.focus();
}
async function versionsDialog() {
  await flushSave();
  let out; try { out = await act("versions", { id: meta.id }); } catch (e) { return fail(e); }
  openDialog(h("h2", {}, "Version history"), h("p", { class: "hint" }, "A copy is kept at most every 10 minutes while you work, plus one before each restore. The newest 60 are kept."),
    out.versions.length ? h("div", { class: "vlist" }, out.versions.map((v) => h("div", { class: "vrow" }, h("span", { class: "grow" }, h("b", {}, new Date(v.created.replace(" ", "T")).toLocaleString()), h("small", {}, ` · ${v.note}`)),
      h("button", { type: "button", class: "ghost", onclick: () => restoreVersion(v.id) }, "Restore"))))
      : h("div", { class: "empty" }, "No earlier versions yet. Keep editing; one is saved every 10 minutes."),
    h("div", { class: "btns" }, h("button", { class: "ghost", value: "cancel" }, "Close")));
}
function restoreVersion(version) {
  const go = async () => {
    try {
      const res = await act("restore", { id: meta.id, version });
      const b = await get(`/api/sheets/book?id=${meta.id}`);
      meta.rev = res.rev;
      for (const sh of b.doc.sheets) { sh.colW = sh.colW || {}; sh.rowH = sh.rowH || {}; sh.freeze = sh.freeze || { r: 0, c: 0 }; sh.rows = sh.rows || 1000; sh.cols = sh.cols || 26; sh.names = sh.names || {}; }
      replaceDoc(b.doc); undoStack = []; redoStack = []; eng.recalc(); dirty = false; switchSheet(0); toast("Restored that version");
    } catch (e) { fail(e); }
  };
  confirmDialog("Go back to this version? What you have now is kept in the history first.", go);
}
function functionsDialog() {
  const q = h("input", { type: "search", placeholder: "Search functions, like lookup or date", "aria-label": "Search functions" });
  const list = h("div", { class: "flist" });
  const cats = ["All", ...new Set(F.FUNCTIONS.map((f) => f.cat))];
  let cat = "All";
  const chips = h("div", { class: "fcats" });
  const show = () => {
    const w = q.value.trim().toLowerCase();
    chips.replaceChildren(...cats.map((c) => h("button", { type: "button", class: `chip ${c === cat ? "on" : ""}`, onclick: () => { cat = c; show(); } }, c)));
    list.replaceChildren(...F.FUNCTIONS.filter((f) => (cat === "All" || f.cat === cat) && (!w || `${f.name} ${f.desc}`.toLowerCase().includes(w)))
      .map((f) => h("button", { type: "button", class: "frow", onclick: () => { dialog.close(); insertFunction(f.name); } }, h("b", {}, `${f.name}(${f.args})`), h("small", {}, f.desc), h("span", { class: "fcat" }, f.cat))));
  };
  q.addEventListener("input", show);
  openDialog(h("h2", {}, `Functions (${F.FUNCTIONS.length})`), q, chips, list, h("div", { class: "btns" }, h("button", { class: "ghost", value: "cancel" }, "Close")));
  show(); q.focus();
}
function shortcutsDialog() {
  const rows = [["Edit the cell", "F2, or Enter with " + KEY], ["New line in a cell", isMac ? "⌥Enter" : "Alt+Enter"], ["Fill the selection with what you typed", `${KEY}Enter`],
    ["Undo / redo", `${KEY}Z / ${KEY}${isMac ? "⇧Z" : "Y"}`], ["Copy, cut, paste", `${KEY}C, ${KEY}X, ${KEY}V`], ["Paste values only", `${KEY}⇧V`],
    ["Bold, italic, underline", `${KEY}B, ${KEY}I, ${KEY}U`], ["Strikethrough", `${KEY}5`], ["Currency, percent, number, date", `${KEY}⇧4, ${KEY}⇧5, ${KEY}⇧1, ${KEY}⇧3`],
    ["Centre / left / right", `${KEY}⇧E, ${KEY}⇧L, ${KEY}⇧R`], ["Clear formatting", `${KEY}\\`], ["Jump to the edge of the data", `${KEY}arrow`], ["Select to the edge", `${KEY}⇧arrow`],
    ["Select column / row", `${KEY}Space / ⇧Space`], ["Select everything", `${KEY}A`], ["Go to A1 / the last cell", `${KEY}Home / ${KEY}End`],
    ["Fill down / right", `${KEY}D / ${KEY}R`], ["Today's date / the time now", `${KEY}; / ${KEY}⇧;`], ["Find / find and replace", `${KEY}F / ${KEY}H`],
    ["Show formulas", `${KEY}\``], ["Switch $ in a reference while typing", "F4"], ["Pick a cell while typing a formula", "click it, or arrow keys"], ["Repeat the last change", "F4 (outside editing)"]];
  openDialog(h("h2", {}, "Keyboard shortcuts"), h("table", { class: "keys" }, rows.map(([a, b]) => h("tr", {}, h("td", {}, a), h("td", {}, h("kbd", {}, b))))),
    h("div", { class: "btns" }, h("button", { class: "ghost", value: "cancel" }, "Close")));
}

// ---------------------------------------------------------------- saving
// Every change is saved a moment later, with each formula's result kept alongside it for exports.
function markDirty() {
  if (!meta) return;
  dirty = true; savedEl.textContent = "Saving…"; savedEl.className = "saved busy";
  clearTimeout(saveTimer); saveTimer = setTimeout(save, 900);
}
function withCache() {
  doc.sheets.forEach((s, i) => {
    for (const [key, cell] of Object.entries(s.cells)) {
      const p = F.parseAddr(key);
      const v = eng.value(i, p.r, p.c), nf = cell.s && cell.s.nf;
      if (typeof cell.v === "string" && cell.v.startsWith("=") && cell.v.length > 1) cell.c = F.isErr(v) ? { e: v.code } : v;
      else delete cell.c;
      // the text a formatted number shows, so a CSV export reads "$1,200.50" or "3/5/2024" like the sheet
      if (nf && typeof v === "number") cell.t = F.format(v, nf).text; else delete cell.t;
    }
    const spilled = {};
    for (const a of eng.anchors.values()) {
      if (a.si !== i) continue;
      for (let r = a.r; r < a.r + a.h; r++) for (let c = a.c; c < a.c + a.w; c++) { if (r === a.r && c === a.c) continue; const v = eng.value(i, r, c); if (v !== null) spilled[A(r, c)] = F.isErr(v) ? { e: v.code } : v; }
    }
    if (Object.keys(spilled).length) s.spilled = spilled; else delete s.spilled;
  });
  return doc;
}
async function save() {
  if (!meta || conflict) return;
  if (saving) { clearTimeout(saveTimer); saveTimer = setTimeout(save, 400); return; }
  saving = true; dirty = false;
  const id = meta.id;
  try {
    const out = await act("save", { id, rev: meta.rev, doc: withCache() });
    if (meta && meta.id === id) { meta.rev = out.rev; if (!dirty) { savedEl.textContent = "All changes saved"; savedEl.className = "saved"; } }
  } catch (e) {
    dirty = true;
    if (/another tab/.test(e.message)) { conflict = true; savedEl.textContent = "Changed elsewhere"; savedEl.className = "saved bad"; openDialog(h("h2", {}, "This spreadsheet changed somewhere else"), h("p", {}, e.message), h("div", { class: "btns" }, h("button", { class: "primary", onclick: () => location.reload() }, "Reload"))); }
    else { savedEl.textContent = "Not saved, retrying…"; savedEl.className = "saved bad"; clearTimeout(saveTimer); saveTimer = setTimeout(save, 4000); }
  } finally { saving = false; }
}
async function flushSave() { clearTimeout(saveTimer); while (saving) await new Promise((r) => setTimeout(r, 50)); if (dirty && meta && !conflict) await save(); }
window.addEventListener("beforeunload", (e) => { if (dirty && meta && !conflict) { save(); e.preventDefault(); e.returnValue = ""; } });
document.addEventListener("visibilitychange", () => { if (document.hidden && dirty) flushSave(); });


// ================================================================ dropdowns, filters, charts, merged cells
// Kept on each sheet (and saved with it):
//   lists:  [{r1, c1, r2, c2, options: [{v, color}], strict}]  a pick list for a range, each choice with a colour
//   filter: {on, hide: {"col": [shown text, …]}}              rows whose value is hidden don't show (height 0)
//   charts: [{id, type, title, src: {r1, c1, r2, c2}, at: {r, c}, w, h}]
//   merges: [{r1, c1, r2, c2}]
const CHIP = { green: ["#ceead6", "#0d652d"], red: ["#fad2cf", "#a50e0e"], yellow: ["#feefc3", "#6d4c00"], blue: ["#d2e3fc", "#174ea6"],
  purple: ["#e9d2fd", "#681da8"], orange: ["#ffdcc2", "#8a3c00"], teal: ["#c2eef0", "#00585c"], gray: ["#e8eaed", "#3c4043"] };
const CHIP_ORDER = Object.keys(CHIP);
const BIG = 1048575;
const inRg = (g, r, c) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;
const rgText = (g) => g.r2 >= BIG ? `${F.colName(g.c1)}:${F.colName(g.c2)}` : rangeText({ r: g.r1, c: g.c1 }, { r: g.r2, c: g.c2 });
function parseRg(text) {
  const t = String(text || "").trim().toUpperCase();
  let m = t.match(/^([A-Z]{1,3}):([A-Z]{1,3})$/);
  if (m) { const a = F.parseAddr(m[1] + "1"), b = F.parseAddr(m[2] + "1"); return { r1: 0, c1: Math.min(a.c, b.c), r2: BIG, c2: Math.max(a.c, b.c) }; }
  m = t.match(/^([A-Z]{1,3}\d+)(?::([A-Z]{1,3}\d+))?$/);
  if (!m) return null;
  const a = F.parseAddr(m[1]), b = F.parseAddr(m[2] || m[1]);
  return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
}
// The selection as a range to apply something to: whole columns stay whole columns.
const selRg = () => { const R = range(); return selKind === "cols" || selKind === "all" ? { r1: 0, c1: R.c1, r2: BIG, c2: R.c2 } : { ...R }; };
// Rows, columns inserted or deleted: ranges move with them (null when everything in it was deleted).
function shiftRg(g, axis, at, n) {
  const k1 = axis === "r" ? "r1" : "c1", k2 = axis === "r" ? "r2" : "c2", out = { ...g };
  if (n > 0) { if (out[k1] >= at) out[k1] += n; if (out[k2] >= at && out[k2] < BIG) out[k2] += n; return out; }
  const d0 = at, d1 = at - n - 1;
  const move = (v) => v < d0 ? v : v > d1 ? v + n : d0;
  if (out[k1] >= d0 && out[k2] <= d1) return null;
  out[k1] = out[k1] > d1 ? out[k1] + n : Math.min(out[k1], d0);
  if (out[k2] < BIG) out[k2] = out[k2] >= d0 && out[k2] <= d1 ? d0 - 1 : move(out[k2]);
  return out[k2] < out[k1] ? null : out;
}
function shiftExtras(s, axis, at, n) {
  if (s.lists) s.lists = s.lists.map((l) => { const g = shiftRg(l, axis, at, n); return g && { ...l, ...g }; }).filter(Boolean);
  if (s.merges) s.merges = s.merges.map((m) => shiftRg(m, axis, at, n)).filter((m) => m && (m.r2 > m.r1 || m.c2 > m.c1));
  if (s.charts) s.charts = s.charts.map((ch) => { const src = shiftRg(ch.src, axis, at, n); if (!src) return null; const p = shiftRg({ r1: ch.at.r, c1: ch.at.c, r2: ch.at.r, c2: ch.at.c }, axis, at, n) || { r1: at, c1: at }; return { ...ch, src, at: { r: axis === "r" ? p.r1 : ch.at.r, c: axis === "c" ? p.c1 : ch.at.c } }; }).filter(Boolean);
  if (s.filter && s.filter.hide && axis === "c") {
    const hide = {};
    for (const [k, v] of Object.entries(s.filter.hide)) { const i = Number(k); if (n < 0 && i >= at && i < at - n) continue; hide[i >= at ? i + n : i] = v; }
    s.filter.hide = hide;
  }
}

// ---------------------------------------------------------------- dropdown lists
function listAt(r, c) { const L = sheet().lists; if (!L || !L.length) return null; for (let i = L.length - 1; i >= 0; i--) if (inRg(L[i], r, c)) return L[i]; return null; }
const optionFor = (rule, text) => rule.options.find((o) => String(o.v).toLowerCase() === String(text).trim().toLowerCase());
function drawChip(rule, d, x, y, w, hgt, active) {
  const opt = d.text ? optionFor(rule, d.text) : null;
  const pad = 3 * zoom, size = Math.round(10 * 1.33 * zoom);
  ctx.save(); ctx.beginPath(); ctx.rect(x, y, w, hgt); ctx.clip();
  if (opt && opt.color && CHIP[opt.color]) {
    const [bg, fc] = CHIP[opt.color], ch = Math.min(hgt - 4 * zoom, size + 8 * zoom);
    ctx.fillStyle = bg; ctx.beginPath(); ctx.roundRect(x + 2 * zoom, y + (hgt - ch) / 2, w - 4 * zoom, ch, ch / 2); ctx.fill();
    ctx.fillStyle = fc;
  } else ctx.fillStyle = d.text && rule.strict && !opt ? "#c5221f" : theme.ink;
  ctx.font = font(d.s);
  if (d.text) ctx.fillText(d.text, x + pad + 5 * zoom, y + hgt / 2 + size * 0.36);
  if (d.text || active) {   // the arrow that opens the list
    const ax = x + w - 11 * zoom, ay = y + hgt / 2;
    ctx.beginPath(); ctx.moveTo(ax - 3.5 * zoom, ay - 2 * zoom); ctx.lineTo(ax + 3.5 * zoom, ay - 2 * zoom); ctx.lineTo(ax, ay + 2.5 * zoom); ctx.closePath();
    ctx.globalAlpha = 0.7; ctx.fill();
  }
  ctx.restore();
}
function openListMenu(r, c) {
  const rule = listAt(r, c); if (!rule) return false;
  const now = shownText(r, c), b = boxRect(r, c, r, c), rect = canvas.getBoundingClientRect();
  const items = rule.options.map((o) => ({ label: String(o.v), color: o.color ? CHIP[o.color][0] : null, check: String(o.v).toLowerCase() === String(now).toLowerCase(),
    action: () => { select(r, c); enterTexts([[String(o.v)]], r, c); } }));
  items.push({ sep: true }, { label: "Clear", action: () => { select(r, c); clearValues(); } });
  openMenu(null, items, false, { clientX: rect.left + b.x, clientY: rect.top + b.y + b.h, preventDefault() {} });
  return true;
}
// A value typed into a strict list must be one of its choices (the same letters in any case become the choice).
function checkList(rows, r0, c0) {
  for (let i = 0; i < rows.length; i++) for (let j = 0; j < rows[i].length; j++) {
    const rule = listAt(r0 + i, c0 + j), t = rows[i][j];
    if (!rule || t === "" || t == null || String(t).startsWith("=")) continue;
    const opt = optionFor(rule, t);
    if (opt) rows[i][j] = String(opt.v);
    else if (rule.strict) { toast(`${A(r0 + i, c0 + j)} only takes: ${rule.options.map((o) => o.v).join(", ")}`, true); return false; }
  }
  return true;
}
function listDialog() {
  const target = listAt(active.r, active.c), base = target ? { ...target } : selRg();
  const rangeIn = h("input", { value: rgText(base), "aria-label": "Range" });
  const rowsEl = h("div", { class: "optlist" });
  let opts = target ? target.options.map((o) => ({ ...o })) : [];
  const strict = h("input", { type: "checkbox", checked: target ? !!target.strict : true });
  const draw2 = () => rowsEl.replaceChildren(...opts.map((o, i) => h("div", { class: "optrow" },
    h("input", { value: o.v, "aria-label": `Choice ${i + 1}`, oninput: (e) => { o.v = e.target.value; } }),
    h("select", { "aria-label": "Colour", onchange: (e) => { o.color = e.target.value || null; draw2(); }, style: o.color ? `background:${CHIP[o.color][0]};color:${CHIP[o.color][1]}` : "" },
      h("option", { value: "" }, "No colour"), ...CHIP_ORDER.map((k) => h("option", { value: k, selected: o.color === k }, k[0].toUpperCase() + k.slice(1)))),
    h("button", { class: "ghost", type: "button", title: "Remove", "aria-label": "Remove choice", onclick: () => { opts.splice(i, 1); draw2(); } }, "✕"))));
  const fromData = () => {
    const g = parseRg(rangeIn.value); if (!g) return;
    const seen = new Map(), b = eng.bounds(si);
    for (let r = g.r1; r <= Math.min(g.r2, b.r); r++) for (let c = g.c1; c <= g.c2; c++) { const t = String(shownText(r, c) || "").trim(); if (t && !seen.has(t.toLowerCase())) seen.set(t.toLowerCase(), t); }
    for (const t of seen.values()) if (!opts.some((o) => String(o.v).toLowerCase() === t.toLowerCase())) opts.push({ v: t, color: CHIP_ORDER[opts.length % CHIP_ORDER.length] });
    draw2();
  };
  draw2();
  const saveIt = (remove) => {
    const g = parseRg(rangeIn.value);
    if (!g) { toast("Use a range like F:F or F2:F50", true); return; }
    const clean2 = opts.map((o) => ({ v: String(o.v).trim(), color: o.color || null })).filter((o) => o.v);
    if (!remove && !clean2.length) { toast("Add at least one choice", true); return; }
    dialog.close();
    structural(remove ? "Remove dropdown" : "Dropdown", () => {
      const s = sheet();
      s.lists = (s.lists || []).filter((l) => l !== target && !(l.r1 === g.r1 && l.c1 === g.c1 && l.r2 === g.r2 && l.c2 === g.c2));
      if (!remove) s.lists.push({ ...g, options: clean2, strict: strict.checked });
    });
  };
  openDialog(h("h2", {}, target ? "Edit dropdown" : "Dropdown list"),
    h("label", { class: "field" }, "Apply to", rangeIn),
    h("div", { class: "hint" }, "Choices, each with an optional colour. Cells show a coloured chip, and an arrow opens the list (or press Alt+↓)."),
    rowsEl,
    h("div", { class: "btns", style: "justify-content:flex-start" },
      h("button", { class: "ghost", type: "button", onclick: () => { opts.push({ v: "", color: CHIP_ORDER[opts.length % CHIP_ORDER.length] }); draw2(); rowsEl.lastChild.querySelector("input").focus(); } }, "+ Add a choice"),
      h("button", { class: "ghost", type: "button", onclick: fromData }, "Use the values already there")),
    h("label", { class: "check" }, strict, " Only allow these choices"),
    h("div", { class: "btns" }, target ? h("button", { class: "ghost bad", onclick: (e) => { e.preventDefault(); saveIt(true); } }, "Remove dropdown") : null, h("span", { class: "grow" }),
      h("button", { class: "ghost", value: "cancel" }, "Cancel"), h("button", { class: "primary", onclick: (e) => { e.preventDefault(); saveIt(false); } }, "Save")));
  if (!opts.length) fromData();
}

// ---------------------------------------------------------------- filters
var hiddenRows = new Set();
const filterOn = () => !!(sheet().filter && sheet().filter.on);
const filterStart = () => (Object.keys(sheet().names || {}).length ? 0 : 1);   // a header in row 1 never hides
function refilter() {
  hiddenRows = new Set();
  const s = sheet(), f = s.filter;
  if (f && f.on && f.hide) {
    const cols = Object.entries(f.hide).filter(([, v]) => v && v.length).map(([k, v]) => [Number(k), new Set(v)]);
    if (cols.length) { const b = eng.bounds(si); for (let r = filterStart(); r <= b.r; r++) if (cols.some(([c, set]) => set.has(String(shownText(r, c) ?? "")))) hiddenRows.add(r); }
  }
  relayout(); sizeSpacer(); draw(); updateStatus();
}
function toggleFilter() {
  const on = !filterOn();
  structural(on ? "Turn on filter" : "Turn off filter", () => { const s = sheet(); if (on) s.filter = { on: true, hide: {} }; else delete s.filter; });
  refilter();
  if (on) toast("Filter is on. Click the arrow in a column's header to sort or filter it.");
}
function sortColumn(c, asc) {
  const b = eng.bounds(si), start = filterStart();
  if (b.r <= start) return toast("Nothing to sort");
  active = { r: start, c }; cursor = { r: start, c };
  sortBy(asc, { r1: start, c1: 0, r2: b.r, c2: Math.max(b.c, c) });
  refilter();
}
function filterMenu(c, ev) {
  const f = sheet().filter || {}, hidden = (f.hide || {})[c] || [];
  openMenu(null, [
    { label: "Sort A to Z", action: () => sortColumn(c, true) }, { label: "Sort Z to A", action: () => sortColumn(c, false) }, { sep: true },
    { label: "Filter by values…", action: () => filterValuesDialog(c) },
    { label: "Clear this column's filter", disabled: !hidden.length, action: () => { structural("Clear filter", () => { delete sheet().filter.hide[c]; }); refilter(); } },
    { label: "Turn off filter", action: toggleFilter },
  ], false, ev);
}
function filterValuesDialog(c) {
  const b = eng.bounds(si), counts = new Map();
  for (let r = filterStart(); r <= b.r; r++) { const t = String(shownText(r, c) ?? ""); counts.set(t, (counts.get(t) || 0) + 1); }
  const hidden = new Set(((sheet().filter || {}).hide || {})[c] || []);
  const values = [...counts.keys()].sort((a, b2) => (a === "") - (b2 === "") || a.localeCompare(b2, undefined, { numeric: true }));
  const boxes = values.map((v) => h("input", { type: "checkbox", checked: !hidden.has(v), value: v }));
  const search = h("input", { placeholder: "Search", "aria-label": "Search values", oninput: () => boxes.forEach((bx) => { bx.parentElement.hidden = !bx.value.toLowerCase().includes(search.value.toLowerCase()); }) });
  const name = (sheet().names || {})[c] || `column ${F.colName(c)}`;
  openDialog(h("h2", {}, `Show rows where ${name} is…`), search,
    h("div", { class: "btns", style: "justify-content:flex-start" },
      h("button", { class: "ghost", type: "button", onclick: () => boxes.forEach((bx) => { if (!bx.parentElement.hidden) bx.checked = true; }) }, "Select all"),
      h("button", { class: "ghost", type: "button", onclick: () => boxes.forEach((bx) => { if (!bx.parentElement.hidden) bx.checked = false; }) }, "Clear")),
    h("div", { class: "vallist" }, boxes.map((bx, i) => h("label", { class: "check" }, bx, ` ${values[i] === "" ? "(Blanks)" : values[i]} `, h("span", { class: "hint" }, String(counts.get(values[i])))))),
    h("div", { class: "btns" }, h("button", { class: "ghost", value: "cancel" }, "Cancel"),
      h("button", { class: "primary", onclick: (e) => {
        e.preventDefault(); dialog.close();
        const hide = boxes.filter((bx) => !bx.checked).map((bx) => bx.value);
        structural("Filter", () => { const s = sheet(); s.filter = s.filter || { on: true, hide: {} }; s.filter.on = true; s.filter.hide = s.filter.hide || {}; if (hide.length) s.filter.hide[c] = hide; else delete s.filter.hide[c]; });
        refilter();
        toast(hiddenRows.size ? `${hiddenRows.size} row${hiddenRows.size === 1 ? "" : "s"} hidden by the filter` : "Showing every row");
      } }, "OK")));
  search.focus();
}
// The filter arrow sits at the right of each header from A to the last column with data.
function filterIconAt(c) { if (!filterOn()) return null; const b = eng.bounds(si); if (c > Math.max(b.c, 0)) return null; const x = colX(c), w = colW(c), s = 14 * Math.min(zoom, 1.25); return { x: x + w - s - 4, y: (hh() - s) / 2, s }; }
function drawFilterIcons() {
  if (!filterOn()) return;
  const b = eng.bounds(si), hide = (sheet().filter || {}).hide || {};
  for (let c = 0; c <= Math.max(b.c, 0); c++) {
    const ic2 = filterIconAt(c); if (!ic2 || ic2.x + ic2.s < hw() || ic2.x > viewW) continue;
    const on = hide[c] && hide[c].length;
    ctx.save(); ctx.fillStyle = on ? theme.accent : theme.head; ctx.strokeStyle = on ? theme.accent : theme.headInk;
    ctx.beginPath(); ctx.roundRect(ic2.x, ic2.y, ic2.s, ic2.s, 3); if (on) ctx.fill(); else ctx.stroke();
    ctx.fillStyle = on ? "#fff" : theme.headInk;   // a funnel
    const cx = ic2.x + ic2.s / 2, cy = ic2.y + ic2.s / 2, u = ic2.s / 14;
    ctx.beginPath(); ctx.moveTo(cx - 4.5 * u, cy - 3.5 * u); ctx.lineTo(cx + 4.5 * u, cy - 3.5 * u); ctx.lineTo(cx + 1 * u, cy + 0.5 * u); ctx.lineTo(cx + 1 * u, cy + 4 * u); ctx.lineTo(cx - 1 * u, cy + 3 * u); ctx.lineTo(cx - 1 * u, cy + 0.5 * u); ctx.closePath(); ctx.fill();
    ctx.restore();
  }
}

// ---------------------------------------------------------------- merged cells
function mergeAt(r, c) { const M = sheet().merges; if (!M || !M.length) return null; for (const m of M) if (inRg(m, r, c)) return m; return null; }
function mergeCells(how) {
  const R = range();
  if (R.r1 === R.r2 && R.c1 === R.c2) return toast("Select two or more cells to merge");
  const groups = how === "rows" ? range2(R.r1, R.r2).map((r) => ({ r1: r, c1: R.c1, r2: r, c2: R.c2 })) : how === "cols" ? range2(R.c1, R.c2).map((c) => ({ r1: R.r1, c1: c, r2: R.r2, c2: c })) : [{ ...R }];
  let lost = 0;
  structural("Merge cells", () => {
    const s = sheet();
    s.merges = (s.merges || []).filter((m) => !(m.r1 <= R.r2 && m.r2 >= R.r1 && m.c1 <= R.c2 && m.c2 >= R.c1));
    for (const g of groups) {
      if (g.r1 === g.r2 && g.c1 === g.c2) continue;
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if ((r !== g.r1 || c !== g.c1) && s.cells[A(r, c)]) { if (s.cells[A(r, c)].v != null && s.cells[A(r, c)].v !== "") lost++; delete s.cells[A(r, c)]; }
      s.merges.push(g);
    }
    eng.touched(si);
  });
  select(R.r1, R.c1);
  if (lost) toast(`Merged. Only the top-left value is kept (${lost} other value${lost === 1 ? "" : "s"} removed; Ctrl+Z brings them back).`);
}
function unmerge() {
  const R = range();
  structural("Unmerge", () => { const s = sheet(); s.merges = (s.merges || []).filter((m) => !(m.r1 <= R.r2 && m.r2 >= R.r1 && m.c1 <= R.c2 && m.c2 >= R.c1)); });
}
// Merged blocks: one background over the whole block (hiding the grid lines inside), and the top-left text across it.
function drawMerges(ra, rb, ca, cb) {
  const M = sheet().merges; if (!M || !M.length) return;
  for (const m of M) {
    if (m.r2 < ra || m.r1 > rb || m.c2 < ca || m.c1 > cb) continue;
    const box = boxRect(m.r1, m.c1, m.r2, m.c2), d = display(m.r1, m.c1), st = d.s || {};
    ctx.fillStyle = st.bg || theme.bg; ctx.fillRect(box.x + 0.5, box.y + 0.5, box.w - 1, box.h - 1);
    if (gridlines) { ctx.strokeStyle = theme.grid; ctx.lineWidth = 1; ctx.strokeRect(Math.round(box.x) - 0.5, Math.round(box.y) - 0.5, Math.round(box.w), Math.round(box.h)); }
    if (!d.text) continue;
    ctx.font = font(st); ctx.fillStyle = d.color ? (COLORS[d.color] || d.color) : st.fc || theme.ink;
    const size = Math.round((st.fs || 10) * 1.33 * zoom), pad = 3 * zoom, align = st.ha || (typeof d.value === "number" ? "right" : "center");
    if (st.wr) { drawWrapped(d.text, box.x, box.y, box.w, box.h, align, { ...st, va: st.va || "middle" }, size); continue; }
    const tw = ctx.measureText(d.text).width, tx = align === "right" ? box.x + box.w - pad - tw : align === "left" ? box.x + pad : box.x + (box.w - tw) / 2;
    const ty = st.va === "top" ? box.y + size + 2 * zoom : st.va === "bottom" ? box.y + box.h - 5 * zoom : box.y + box.h / 2 + size * 0.36;
    ctx.save(); ctx.beginPath(); ctx.rect(box.x, box.y, box.w, box.h); ctx.clip(); ctx.fillText(d.text, tx, ty); ctx.restore();
  }
}

// ---------------------------------------------------------------- charts
const CHART_COLORS = ["#4285f4", "#ea4335", "#fbbc04", "#34a853", "#ff6d01", "#46bdc6", "#7baaf7", "#f07b72", "#a142f4", "#24c1e0"];
var chartEls = new Map();
// What a chart shows: text in one column is counted (how many of each); otherwise the first text column names the
// points and each number column is a series. Rows hidden by a filter are left out, like both apps.
function chartData(ch) {
  const g = { ...ch.src }, b = eng.bounds(si);
  g.r2 = Math.min(g.r2, b.r);
  const names = sheet().names || {};
  let rows = range2(g.r1, g.r2).filter((r) => !hiddenRows.has(r));
  const val = (r, c) => eng.value(si, r, c);
  // a header row inside the range (text over numbers) names the series
  let heads = {};
  if (rows.length > 1 && range2(g.c1, g.c2).every((c) => { const v = val(rows[0], c); return v === null || typeof v === "string"; }) && range2(g.c1, g.c2).some((c) => typeof val(rows[1], c) === "number")) {
    for (const c of range2(g.c1, g.c2)) heads[c] = String(val(rows[0], c) ?? ""); rows = rows.slice(1);
  }
  const colName = (c) => heads[c] || names[c] || `Column ${F.colName(c)}`;
  const numeric = range2(g.c1, g.c2).filter((c) => rows.some((r) => typeof val(r, c) === "number"));
  if (!numeric.length) {   // count the values in the first column
    const counts = new Map();
    for (const r of rows) { const t = String(shownText(r, g.c1) ?? "").trim(); if (t) counts.set(t, (counts.get(t) || 0) + 1); }
    const list = [...counts.entries()].sort((a, b2) => b2[1] - a[1]);
    return { labels: list.map((x) => x[0]), series: [{ name: `Count of ${colName(g.c1)}`, values: list.map((x) => x[1]) }], counted: true };
  }
  const labelCol = range2(g.c1, g.c2).find((c) => !numeric.includes(c));
  const labels = rows.map((r) => labelCol !== undefined ? String(shownText(r, labelCol) ?? "") : String(r + 1));
  return { labels, series: numeric.map((c) => ({ name: colName(c), values: rows.map((r) => { const v = val(r, c); return typeof v === "number" ? v : 0; }) })) };
}
function niceMax(v) { if (v <= 0) return 1; const p = Math.pow(10, Math.floor(Math.log10(v))); for (const m of [1, 2, 2.5, 5, 10]) if (m * p >= v) return m * p; return 10 * p; }
const short = (n) => Math.abs(n) >= 1e6 ? (n / 1e6).toFixed(1).replace(/\.0$/, "") + "M" : Math.abs(n) >= 1e3 ? (n / 1e3).toFixed(1).replace(/\.0$/, "") + "k" : String(Math.round(n * 100) / 100);
function renderChart(ch, cv) {
  const dpr = window.devicePixelRatio || 1, W = cv.clientWidth, H = cv.clientHeight;
  cv.width = W * dpr; cv.height = H * dpr;
  const c2 = cv.getContext("2d"); c2.setTransform(dpr, 0, 0, dpr, 0, 0);
  const ink = theme.ink, soft = theme.headInk, grid = theme.grid;
  c2.clearRect(0, 0, W, H); c2.font = `12px ${FONT}`; c2.textBaseline = "middle";
  const data = chartData(ch), S = data.series, L = data.labels;
  if (!L.length || !S.length) { c2.fillStyle = soft; c2.textAlign = "center"; c2.fillText("No data in " + rgText(ch.src), W / 2, H / 2); return; }
  if (ch.type === "pie") {
    const vals = S[0].values.map((v) => Math.max(0, v)), total = vals.reduce((a, b) => a + b, 0) || 1;
    const legendW = Math.min(170, W * 0.42), R2 = Math.max(10, Math.min(W - legendW - 24, H - 16) / 2), cx = 12 + R2, cy = H / 2;
    let a0 = -Math.PI / 2;
    vals.forEach((v, i) => { const a1 = a0 + (v / total) * Math.PI * 2; c2.fillStyle = CHART_COLORS[i % CHART_COLORS.length]; c2.beginPath(); c2.moveTo(cx, cy); c2.arc(cx, cy, R2, a0, a1); c2.closePath(); c2.fill();
      if (v / total > 0.06) { const m = (a0 + a1) / 2; c2.fillStyle = "#fff"; c2.textAlign = "center"; c2.fillText(Math.round((v / total) * 100) + "%", cx + Math.cos(m) * R2 * 0.62, cy + Math.sin(m) * R2 * 0.62); }
      a0 = a1; });
    c2.textAlign = "left";
    const lx = cx + R2 + 16, step = Math.min(20, (H - 10) / L.length);
    L.forEach((lab, i) => { const y = H / 2 - (L.length * step) / 2 + i * step + step / 2; if (y < 0 || y > H) return; c2.fillStyle = CHART_COLORS[i % CHART_COLORS.length]; c2.fillRect(lx, y - 5, 10, 10); c2.fillStyle = ink;
      let t = `${lab} (${short(vals[i])})`; while (t.length > 3 && c2.measureText(t).width > W - lx - 18) t = t.slice(0, -2) + "…"; c2.fillText(t, lx + 15, y); });
    return;
  }
  const horiz = ch.type === "bar", legend = S.length > 1 ? 18 : 0;
  const all = S.flatMap((s) => s.values), max = niceMax(Math.max(0, ...all)), min = Math.min(0, ...all) < 0 ? -niceMax(-Math.min(...all)) : 0;
  const labW = horiz ? Math.min(W * 0.35, Math.max(...L.map((l) => c2.measureText(l).width)) + 10) : 0;
  const left = horiz ? labW + 6 : Math.max(...[max, min].map((v) => c2.measureText(short(v)).width)) + 12, right = 12, top = 8 + legend, bottom = horiz ? 22 : 30;
  const pw = W - left - right, ph = H - top - bottom;
  if (pw < 20 || ph < 20) return;
  const scale = (v) => (v - min) / (max - min || 1);
  // value gridlines
  c2.strokeStyle = grid; c2.fillStyle = soft; c2.lineWidth = 1;
  for (let i = 0; i <= 4; i++) {
    const v = min + ((max - min) * i) / 4;
    if (horiz) { const x = left + scale(v) * pw; c2.beginPath(); c2.moveTo(Math.round(x) + 0.5, top); c2.lineTo(Math.round(x) + 0.5, top + ph); c2.stroke(); c2.textAlign = "center"; c2.fillText(short(v), x, top + ph + 12); }
    else { const y = top + ph - scale(v) * ph; c2.beginPath(); c2.moveTo(left, Math.round(y) + 0.5); c2.lineTo(left + pw, Math.round(y) + 0.5); c2.stroke(); c2.textAlign = "right"; c2.fillText(short(v), left - 6, y); }
  }
  const n = L.length, slot = (horiz ? ph : pw) / n;
  if (ch.type === "line") {
    S.forEach((s, si2) => { c2.strokeStyle = CHART_COLORS[si2 % CHART_COLORS.length]; c2.lineWidth = 2; c2.beginPath();
      s.values.forEach((v, i) => { const x = left + slot * (i + 0.5), y = top + ph - scale(v) * ph; if (i) c2.lineTo(x, y); else c2.moveTo(x, y); }); c2.stroke();
      c2.fillStyle = c2.strokeStyle; s.values.forEach((v, i) => { if (n <= 40) { c2.beginPath(); c2.arc(left + slot * (i + 0.5), top + ph - scale(v) * ph, 2.5, 0, 7); c2.fill(); } }); });
  } else {
    const bw = Math.max(1, (slot * 0.72) / S.length), zero = scale(0);
    S.forEach((s, si2) => { c2.fillStyle = CHART_COLORS[si2 % CHART_COLORS.length];
      s.values.forEach((v, i) => { const o = slot * 0.14 + si2 * bw + i * slot, a = scale(v);
        if (horiz) { const x0 = left + Math.min(zero, a) * pw; c2.fillRect(x0, top + o, Math.abs(a - zero) * pw, bw - 1); }
        else { const y0 = top + ph - Math.max(zero, a) * ph; c2.fillRect(left + o, y0, bw - 1, Math.abs(a - zero) * ph); } }); });
    if (S.length === 1 && n <= 24) {   // the numbers on the bars
      c2.fillStyle = ink; c2.font = `11px ${FONT}`;
      S[0].values.forEach((v, i) => { const a = scale(v); if (horiz) { c2.textAlign = "left"; c2.fillText(short(v), left + a * pw + 4, top + slot * (i + 0.5)); } else { c2.textAlign = "center"; c2.fillText(short(v), left + slot * (i + 0.5), top + ph - a * ph - 8); } });
      c2.font = `12px ${FONT}`;
    }
  }
  // category labels
  c2.fillStyle = soft;
  const every = Math.ceil(n / Math.max(1, Math.floor((horiz ? ph : pw) / (horiz ? 16 : 60))));
  L.forEach((lab, i) => { if (i % every) return; let t = lab; const room = horiz ? labW - 4 : slot * every - 4; while (t.length > 2 && c2.measureText(t).width > room) t = t.slice(0, -2) + "…";
    if (horiz) { c2.textAlign = "right"; c2.fillText(t, left - 6, top + slot * (i + 0.5)); } else { c2.textAlign = "center"; c2.fillText(t, left + slot * (i + 0.5), top + ph + 14); } });
  if (legend) { c2.textAlign = "left"; let x = left; S.forEach((s, i) => { c2.fillStyle = CHART_COLORS[i % CHART_COLORS.length]; c2.fillRect(x, 4, 10, 10); c2.fillStyle = ink; c2.fillText(s.name, x + 14, 9); x += c2.measureText(s.name).width + 32; }); }
}
function positionCharts() {
  const list = (doc && sheet().charts) || [], keep = new Set();
  for (const ch of list) {
    keep.add(ch.id);
    let el = chartEls.get(ch.id);
    if (!el) {
      const cv = h("canvas", { class: "chartcv" });
      const title = h("div", { class: "charttitle" });
      const menuBtn = h("button", { class: "ghost chartmenu", title: "Chart options", "aria-label": "Chart options", onmousedown: (e) => e.stopPropagation(),
        onclick: (e) => openMenu(e.currentTarget, [{ label: "Edit chart…", action: () => chartDialog(ch.id) }, { label: "Delete chart", action: () => structural("Delete chart", () => { sheet().charts = sheet().charts.filter((x) => x.id !== ch.id); }) }]) }, "⋮");
      const grip = h("div", { class: "chartgrip", title: "Resize" });
      el = h("div", { class: "chart" }, h("div", { class: "charthead" }, title, menuBtn), cv, grip);
      el._cv = cv; el._title = title;
      el.addEventListener("mousedown", (e) => chartDrag(e, ch.id, e.target === grip ? "size" : "move"));
      el.addEventListener("dblclick", () => chartDialog(ch.id));
      wrapEl.append(el); chartEls.set(ch.id, el);
    }
    const x = colX(ch.at.c), y = rowY(ch.at.r), w = ch.w * Math.min(zoom, 1.5), hgt = ch.h * Math.min(zoom, 1.5);
    el.style.left = x + "px"; el.style.top = y + "px"; el.style.width = w + "px"; el.style.height = hgt + "px";
    const cutL = Math.max(0, hw() - x), cutT = Math.max(0, hh() - y);
    el.style.clipPath = cutL || cutT ? `inset(${cutT}px 0 0 ${cutL}px)` : "";
    el.hidden = x > viewW || y > viewH || x + w < hw() || y + hgt < hh();
    el._title.textContent = ch.title || chartData(ch).series.map((s) => s.name).join(", ");
    const key = JSON.stringify([ch, w, hgt, chartData(ch), theme.ink]);
    if (!el.hidden && el._key !== key) { el._key = key; renderChart(ch, el._cv); }
  }
  for (const [id, el] of chartEls) if (!keep.has(id)) { el.remove(); chartEls.delete(id); }
}
function clearCharts() { for (const el of chartEls.values()) el.remove(); chartEls.clear(); }
function chartDrag(e, id, kind) {
  if (e.button !== 0 || e.target.closest(".chartmenu")) return;
  e.preventDefault();
  const ch = sheet().charts.find((x) => x.id === id), el = chartEls.get(id), x0 = e.clientX, y0 = e.clientY, start = { ...ch, at: { ...ch.at } };
  const rect = canvas.getBoundingClientRect(), zz = Math.min(zoom, 1.5);
  const move = (ev) => {
    if (kind === "size") { ch.w = Math.max(220, start.w + (ev.clientX - x0) / zz); ch.h = Math.max(150, start.h + (ev.clientY - y0) / zz); }
    else { el.style.transform = `translate(${ev.clientX - x0}px, ${ev.clientY - y0}px)`; }
    if (kind === "size") positionCharts();
  };
  const up = (ev) => {
    window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up);
    el.style.transform = "";
    const moved = Math.abs(ev.clientX - x0) + Math.abs(ev.clientY - y0) > 3;
    if (!moved) { if (kind === "size") { ch.w = start.w; ch.h = start.h; } return; }
    const after = { w: ch.w, h: ch.h };
    ch.w = start.w; ch.h = start.h;
    structural(kind === "size" ? "Resize chart" : "Move chart", () => {
      if (kind === "size") { ch.w = Math.round(after.w); ch.h = Math.round(after.h); }
      else { const px = colX(start.at.c) + ev.clientX - x0, py = rowY(start.at.r) + ev.clientY - y0; ch.at = { r: rowFromY(Math.max(hh(), py)), c: colFromX(Math.max(hw(), px)) }; }
    });
    void rect;
  };
  window.addEventListener("mousemove", move); window.addEventListener("mouseup", up);
}
function chartDialog(id) {
  const old = id && (sheet().charts || []).find((x) => x.id === id);
  let src = old ? old.src : null;
  if (!src) {
    const R = range(), b = eng.bounds(si);
    src = R.r1 !== R.r2 || R.c1 !== R.c2 ? { ...R, r2: Math.min(R.r2, Math.max(b.r, R.r1)) } : currentRegion();
    if (selKind === "cols") src = { r1: 0, c1: R.c1, r2: b.r, c2: R.c2 };
  }
  const type = h("select", { "aria-label": "Chart type" }, [["column", "Column"], ["bar", "Bar"], ["line", "Line"], ["pie", "Pie"]].map(([v, l]) => h("option", { value: v, selected: (old ? old.type : "column") === v }, l)));
  const title = h("input", { value: old ? old.title || "" : "", placeholder: "Leave empty for the column names", "aria-label": "Chart title" });
  const rangeIn = h("input", { value: rangeText({ r: src.r1, c: src.c1 }, { r: src.r2, c: src.c2 }), "aria-label": "Data range" });
  openDialog(h("h2", {}, old ? "Edit chart" : "Insert chart"),
    h("label", { class: "field" }, "Data", rangeIn),
    h("div", { class: "hint" }, "One column of text counts each value (like how many jobs are in each Progress). A text column next to number columns names the points, and each number column is drawn."),
    h("label", { class: "field" }, "Type", type), h("label", { class: "field" }, "Title", title),
    h("div", { class: "btns" }, h("button", { class: "ghost", value: "cancel" }, "Cancel"),
      h("button", { class: "primary", onclick: (e) => {
        e.preventDefault();
        const g = parseRg(rangeIn.value); if (!g) { toast("Use a range like F1:F40", true); return; }
        dialog.close();
        structural(old ? "Edit chart" : "Insert chart", () => {
          const s = sheet(); s.charts = s.charts || [];
          if (old) Object.assign(s.charts.find((x) => x.id === id), { src: g, type: type.value, title: title.value.trim() });
          else s.charts.push({ id: "c" + Date.now().toString(36), type: type.value, title: title.value.trim(), src: g, at: { r: g.r1, c: Math.min(g.c2 + 2, sheet().cols - 1) }, w: 440, h: 270 });
        });
      } }, old ? "Save" : "Insert")));
}

route().catch((e) => { app.replaceChildren(h("div", { class: "home" }, h("h1", {}, "Can't load Sheets"), h("p", {}, e.message))); });
