"""The SQLite file behind the Sheets tab. Every rule about what can be saved lives here.

    books      one row per workbook; the workbook itself is JSON (see below), with a revision number so two open tabs
               can't silently overwrite each other
    versions   snapshots for "version history": one at most every 10 minutes while you edit, and one before anything
               that replaces a workbook wholesale (restore). The newest 60 per workbook are kept.

A workbook: {"sheets": [{"name", "rows", "cols", "cells": {"A1": {"v": value, "s": {style}, "c": cached result}},
"colW": {"0": px}, "rowH": {"0": px}, "freeze": {"r": n, "c": n}, "names": {"0": "Company"}}], "active": 0}.
names are column names shown in the header bar instead of the letters; formulas still use the letters. v is a number, true/false, text, or
text starting with "=" for a formula; c is the last calculated result of a formula, and t the text a formatted cell shows
(like "$1,200.50" or "3/5/2024"), both kept for exports and the data hub.
"""
import datetime as dt
import json
import os
import re
import sqlite3
from pathlib import Path

MAX_DOC = 25 * 1024 * 1024
MAX_SHEETS = 200
MAX_NAME = 100
MAX_CELLS = 2_000_000
KEEP_VERSIONS = 60
SNAPSHOT_EVERY = 600  # seconds
ADDR = re.compile(r"^[A-Z]{1,3}[1-9]\d{0,6}$")
BAD_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")

SCHEMA = """
CREATE TABLE IF NOT EXISTS books (
    id       INTEGER PRIMARY KEY,
    name     TEXT NOT NULL,
    doc      TEXT NOT NULL,
    rev      INTEGER NOT NULL DEFAULT 1,
    size     INTEGER NOT NULL DEFAULT 0,
    created  TEXT NOT NULL,
    updated  TEXT NOT NULL,
    opened   TEXT
);
CREATE TABLE IF NOT EXISTS versions (
    id       INTEGER PRIMARY KEY,
    book_id  INTEGER NOT NULL REFERENCES books(id) ON DELETE CASCADE,
    doc      TEXT NOT NULL,
    note     TEXT NOT NULL DEFAULT '',
    created  TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS versions_book ON versions(book_id, id);
"""


class SheetsError(ValueError):
    """Something you can fix. The message is meant for you."""


def app_dir():
    return Path(__file__).resolve().parent.parent


def db_path():
    return Path(os.environ.get("SHEETS_DB") or app_dir() / "data" / "sheets.db")


def now():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def connect(path=None):
    path = Path(path or db_path())
    path.parent.mkdir(parents=True, exist_ok=True)
    fresh = not path.exists()
    conn = sqlite3.connect(path, timeout=10, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.executescript(SCHEMA)
    if fresh:
        os.chmod(path, 0o600)
    return conn


def blank_sheet(name="Sheet1"):
    return {"name": name, "rows": 1000, "cols": 26, "cells": {}, "colW": {}, "rowH": {}, "freeze": {"r": 0, "c": 0}, "names": {}}


def blank_doc():
    return {"sheets": [blank_sheet()], "active": 0}


def clean_name(name, what="The workbook name"):
    text = " ".join(str(name or "").split())
    if not text:
        raise SheetsError(f"{what} can't be empty")
    if len(text) > MAX_NAME:
        raise SheetsError(f"{what} is too long (keep it under {MAX_NAME} characters)")
    return text


def clean_doc(doc):
    """Checks the shape of a workbook before it's saved, so a bad save can't leave something that won't open."""
    if not isinstance(doc, dict) or not isinstance(doc.get("sheets"), list) or not doc["sheets"]:
        raise SheetsError("A workbook needs at least one sheet")
    if len(doc["sheets"]) > MAX_SHEETS:
        raise SheetsError(f"A workbook can have up to {MAX_SHEETS} sheets")
    names, total = set(), 0
    out = {"sheets": [], "active": 0}
    for sheet in doc["sheets"]:
        if not isinstance(sheet, dict) or not isinstance(sheet.get("cells"), dict):
            raise SheetsError("Every sheet needs its cells")
        name = clean_name(sheet.get("name"), "A sheet name")
        if BAD_SHEET_CHARS.search(name) or name.startswith("'") or name.endswith("'"):
            raise SheetsError(f"“{name}” can't be a sheet name (no [ ] : * ? / \\ and no ' at the ends)")
        if name.lower() in names:
            raise SheetsError(f"Two sheets are called “{name}”")
        names.add(name.lower())
        cells = {}
        for key, cell in sheet["cells"].items():
            if not ADDR.match(key) or not isinstance(cell, dict):
                raise SheetsError(f"Bad cell {key[:20]}")
            v = cell.get("v")
            if v is not None and not isinstance(v, (str, int, float, bool)):
                raise SheetsError(f"Bad value in {key}")
            if isinstance(v, str) and len(v) > 50000:
                raise SheetsError(f"{key} holds more than 50,000 characters")
            kept = {k: cell[k] for k in ("v", "s", "c", "n", "t") if k in cell and cell[k] is not None}
            if kept:
                cells[key] = kept
        total += len(cells)
        if total > MAX_CELLS:
            raise SheetsError(f"A workbook can hold up to {MAX_CELLS:,} filled cells")
        clean = {"name": name, "cells": cells}
        for key, low, high in (("rows", 1, 1_048_576), ("cols", 1, 16_384)):
            try:
                clean[key] = min(high, max(low, int(sheet.get(key) or (1000 if key == "rows" else 26))))
            except (TypeError, ValueError):
                raise SheetsError(f"Bad {key} count") from None
        for key in ("colW", "rowH"):
            sizes = sheet.get(key) or {}
            if not isinstance(sizes, dict):
                raise SheetsError("Bad column widths")
            clean[key] = {str(int(k)): max(2, min(2000, int(v))) for k, v in sizes.items() if str(k).isdigit() and isinstance(v, (int, float))}
        # Column names: shown in the header bar instead of A, B, C, so a header doesn't take up row 1.
        col_names = sheet.get("names") or {}
        if not isinstance(col_names, dict):
            raise SheetsError("Bad column names")
        clean["names"] = {str(int(k)): " ".join(v.split())[:100] for k, v in col_names.items()
                          if str(k).isdigit() and isinstance(v, str) and v.strip()}
        freeze = sheet.get("freeze") or {}
        clean["freeze"] = {"r": max(0, min(50, int(freeze.get("r") or 0))), "c": max(0, min(26, int(freeze.get("c") or 0)))}
        for key in ("spilled", "hidden", "filter", "tab"):
            if key in sheet:
                clean[key] = sheet[key]
        clean.update(_clean_extras(sheet))
        out["sheets"].append(clean)
    try:
        out["active"] = max(0, min(len(out["sheets"]) - 1, int(doc.get("active") or 0)))
    except (TypeError, ValueError):
        out["active"] = 0
    return out


def _rg(g, what):
    try:
        out = {k: int(g[k]) for k in ("r1", "c1", "r2", "c2")}
    except (KeyError, TypeError, ValueError):
        raise SheetsError(f"Bad {what}") from None
    if not (0 <= out["r1"] <= out["r2"] <= 1_048_575 and 0 <= out["c1"] <= out["c2"] <= 16_383):
        raise SheetsError(f"Bad {what}")
    return out


def _clean_extras(sheet):
    """Dropdown lists, merged cells and charts, checked so a bad save can't break the sheet."""
    out = {}
    lists = sheet.get("lists")
    if lists:
        if not isinstance(lists, list) or len(lists) > 500:
            raise SheetsError("Bad dropdowns")
        out["lists"] = []
        for rule in lists:
            g = _rg(rule, "dropdown")
            opts = [{"v": str(o.get("v"))[:200], "color": o.get("color") if isinstance(o.get("color"), str) and len(o.get("color")) <= 20 else None}
                    for o in (rule.get("options") or [])[:500] if isinstance(o, dict) and str(o.get("v", "")).strip()]
            if opts:
                out["lists"].append({**g, "options": opts, "strict": bool(rule.get("strict"))})
    merges = sheet.get("merges")
    if merges:
        if not isinstance(merges, list) or len(merges) > 20_000:
            raise SheetsError("Bad merged cells")
        out["merges"] = [g for g in (_rg(m, "merged cells") for m in merges) if g["r2"] > g["r1"] or g["c2"] > g["c1"]]
    charts = sheet.get("charts")
    if charts:
        if not isinstance(charts, list) or len(charts) > 100:
            raise SheetsError("Bad charts")
        out["charts"] = []
        for ch in charts:
            if not isinstance(ch, dict) or ch.get("type") not in ("column", "bar", "line", "pie"):
                raise SheetsError("Bad chart")
            at = ch.get("at") or {}
            out["charts"].append({"id": str(ch.get("id") or "")[:40], "type": ch["type"], "title": str(ch.get("title") or "")[:200],
                                  "src": _rg(ch.get("src") or {}, "chart range"),
                                  "at": {"r": max(0, int(at.get("r") or 0)), "c": max(0, int(at.get("c") or 0))},
                                  "w": max(150, min(3000, int(ch.get("w") or 440))), "h": max(100, min(3000, int(ch.get("h") or 270)))})
    return out


def encode(doc):
    text = json.dumps(doc, separators=(",", ":"), ensure_ascii=False)
    if len(text.encode()) > MAX_DOC:
        raise SheetsError("That workbook is too big to save (over 25 MB)")
    return text


# ---------------------------------------------------------------- books

def books(conn):
    out = []
    for r in conn.execute("SELECT id, name, size, created, updated, opened, doc FROM books ORDER BY coalesce(opened, updated) DESC, id DESC"):
        try:
            sheets = [s["name"] for s in json.loads(r["doc"])["sheets"]]
        except (ValueError, KeyError, TypeError):
            sheets = []
        out.append({"id": r["id"], "name": r["name"], "size": r["size"], "created": r["created"], "updated": r["updated"], "opened": r["opened"], "sheets": sheets})
    return out


def book_row(conn, book_id):
    row = conn.execute("SELECT * FROM books WHERE id = ?", (book_id,)).fetchone()
    if not row:
        raise SheetsError("That workbook isn't here anymore")
    return row


def add_book(conn, name, doc=None):
    name = clean_name(name)
    text = encode(clean_doc(doc or blank_doc()))
    stamp = now()
    with conn:
        cur = conn.execute("INSERT INTO books (name, doc, size, created, updated) VALUES (?, ?, ?, ?, ?)", (name, text, len(text), stamp, stamp))
    return cur.lastrowid


def save(conn, book_id, rev, doc):
    """Saves if nobody else saved since you opened it (rev). Returns the new revision."""
    row = book_row(conn, book_id)
    if rev != row["rev"]:
        raise SheetsError("This workbook was changed in another tab or window. Reload to get the latest, then redo your last change.")
    text = encode(clean_doc(doc))
    if text == row["doc"]:
        return row["rev"]
    last = conn.execute("SELECT created FROM versions WHERE book_id = ? ORDER BY id DESC LIMIT 1", (book_id,)).fetchone()
    stamp = now()
    with conn:
        if not last or (dt.datetime.now() - dt.datetime.strptime(last[0], "%Y-%m-%d %H:%M:%S")).total_seconds() >= SNAPSHOT_EVERY:
            snapshot(conn, book_id, row["doc"], "Autosave")
        conn.execute("UPDATE books SET doc = ?, size = ?, rev = rev + 1, updated = ? WHERE id = ?", (text, len(text), stamp, book_id))
    return row["rev"] + 1


def snapshot(conn, book_id, text, note):
    conn.execute("INSERT INTO versions (book_id, doc, note, created) VALUES (?, ?, ?, ?)", (book_id, text, note, now()))
    conn.execute("""DELETE FROM versions WHERE book_id = ? AND id NOT IN
                    (SELECT id FROM versions WHERE book_id = ? ORDER BY id DESC LIMIT ?)""", (book_id, book_id, KEEP_VERSIONS))


def rename(conn, book_id, name):
    book_row(conn, book_id)
    with conn:
        conn.execute("UPDATE books SET name = ?, updated = ? WHERE id = ?", (clean_name(name), now(), book_id))


def delete(conn, book_id):
    book_row(conn, book_id)
    with conn:
        conn.execute("DELETE FROM books WHERE id = ?", (book_id,))


def opened(conn, book_id):
    with conn:
        conn.execute("UPDATE books SET opened = ? WHERE id = ?", (now(), book_id))


def versions(conn, book_id):
    book_row(conn, book_id)
    return [dict(r) for r in conn.execute("SELECT id, note, created, length(doc) size FROM versions WHERE book_id = ? ORDER BY id DESC", (book_id,))]


def restore(conn, book_id, version_id):
    row = book_row(conn, book_id)
    v = conn.execute("SELECT doc FROM versions WHERE id = ? AND book_id = ?", (version_id, book_id)).fetchone()
    if not v:
        raise SheetsError("That version isn't here anymore")
    with conn:
        snapshot(conn, book_id, row["doc"], "Before restoring an older version")
        conn.execute("UPDATE books SET doc = ?, size = ?, rev = rev + 1, updated = ? WHERE id = ?", (v[0], len(v[0]), now(), book_id))
    return row["rev"] + 1
