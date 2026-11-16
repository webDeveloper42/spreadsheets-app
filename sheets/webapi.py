"""What the dashboard's Sheets tab can do. Storage rules come from store.py, files from files.py.

Formulas are worked out in the page (static/sheets/formula.js), which saves each formula's last result alongside it so
exports and the data hub can see values without running the engine here.
"""
import json
import re
import threading
from pathlib import Path

from . import files, store
from .store import SheetsError

MAX_IMPORT = 50 * 1024 * 1024


def _t(cells):
    """{"A1": value or (value, style)} -> cells with parsed values"""
    out = {}
    for key, v in cells.items():
        if isinstance(v, tuple):
            out[key] = {"v": v[0], "s": v[1]}
        else:
            out[key] = {"v": v}
    return out


HEAD = {"b": 1, "bg": "#e8eaed"}
MONEY = {"nf": "$#,##0.00"}
TEMPLATES = {
    "budget": ("Monthly budget", lambda: {"sheets": [{**store.blank_sheet("Budget"), "colW": {"0": 180, "1": 110, "2": 110, "3": 110}, "freeze": {"r": 1, "c": 0},
        "cells": _t({"A1": ("Category", HEAD), "B1": ("Planned", HEAD), "C1": ("Spent", HEAD), "D1": ("Left", HEAD),
                     "A2": "Rent", "B2": (1200, MONEY), "C2": (1200, MONEY), "A3": "Groceries", "B3": (400, MONEY), "C3": (312.45, MONEY),
                     "A4": "Transport", "B4": (150, MONEY), "C4": (96, MONEY), "A5": "Fun", "B5": (100, MONEY), "C5": (130, MONEY),
                     "D2": ("=B2-C2", MONEY), "D3": ("=B3-C3", MONEY), "D4": ("=B4-C4", MONEY), "D5": ("=B5-C5", MONEY),
                     "A7": ("Total", {"b": 1}), "B7": ("=SUM(B2:B5)", {**MONEY, "b": 1}), "C7": ("=SUM(C2:C5)", {**MONEY, "b": 1}), "D7": ("=B7-C7", {**MONEY, "b": 1}),
                     "A9": "Over budget in", "B9": '=TEXTJOIN(", ",TRUE,FILTER(A2:A5,D2:D5<0,""))'})}], "active": 0}),
    "todo": ("To-do list", lambda: {"sheets": [{**store.blank_sheet("Tasks"), "colW": {"0": 60, "1": 260, "2": 110, "3": 90}, "freeze": {"r": 1, "c": 0},
        "cells": _t({"A1": ("Done", HEAD), "B1": ("Task", HEAD), "C1": ("Due", HEAD), "D1": ("Days left", HEAD),
                     "A2": False, "B2": "Write the report", "C2": ("=TODAY()+3", {"nf": "mmm d"}), "D2": '=IF(A2,"done",C2-TODAY())',
                     "A3": True, "B3": "Pay the electric bill", "C3": ("=TODAY()-1", {"nf": "mmm d"}), "D3": '=IF(A3,"done",C3-TODAY())',
                     "F1": ("Progress", {"b": 1}), "F2": ("=COUNTIF(A2:A50,TRUE)/MAX(1,COUNTA(B2:B50))", {"nf": "0%"})})}], "active": 0}),
    "loan": ("Loan calculator", lambda: {"sheets": [{**store.blank_sheet("Loan"), "colW": {"0": 200, "1": 130},
        "cells": _t({"A1": ("Loan amount", {"b": 1}), "B1": (20000, MONEY), "A2": ("Yearly interest rate", {"b": 1}), "B2": (0.065, {"nf": "0.00%"}),
                     "A3": ("Years", {"b": 1}), "B3": 5, "A5": ("Monthly payment", {"b": 1}), "B5": ("=-PMT(B2/12,B3*12,B1)", {**MONEY, "b": 1}),
                     "A6": "Total paid", "B6": ("=B5*B3*12", MONEY), "A7": "Total interest", "B7": ("=B6-B1", MONEY),
                     "D1": ("Month", HEAD), "E1": ("Interest", HEAD), "F1": ("Principal", HEAD)})}], "active": 0}),
    "workout": ("Workout log", lambda: {"sheets": [{**store.blank_sheet("Log"), "colW": {"0": 110, "1": 160}, "freeze": {"r": 1, "c": 0},
        "cells": _t({"A1": ("Date", HEAD), "B1": ("Exercise", HEAD), "C1": ("Sets", HEAD), "D1": ("Reps", HEAD), "E1": ("Weight", HEAD), "F1": ("Volume", HEAD),
                     "A2": ("=TODAY()", {"nf": "m/d/yyyy"}), "B2": "Squat", "C2": 3, "D2": 8, "E2": 135, "F2": "=C2*D2*E2",
                     "H1": ("Total volume", {"b": 1}), "H2": "=SUM(F:F)"})}], "active": 0}),
}
# The loan's month-by-month table: plain formulas, so it also works after exporting to Excel or Google Sheets.
TEMPLATES["loan"] = (TEMPLATES["loan"][0], (lambda make: lambda: _loan_table(make()))(TEMPLATES["loan"][1]))


def _loan_table(doc):
    cells = doc["sheets"][0]["cells"]
    cells["D2"] = {"v": 1}
    for r in range(2, 62):
        if r > 2:
            cells[f"D{r}"] = {"v": f"=IF(D{r - 1}<$B$3*12,D{r - 1}+1,\"\")"}
        cells[f"E{r}"] = {"v": f'=IF(D{r}="","",-IPMT($B$2/12,D{r},$B$3*12,$B$1))', "s": MONEY}
        cells[f"F{r}"] = {"v": f'=IF(D{r}="","",-PPMT($B$2/12,D{r},$B$3*12,$B$1))', "s": MONEY}
    return doc


class SheetsApi:
    def __init__(self, path=None):
        self.conn = store.connect(path)
        self.lock = threading.Lock()

    def close(self):
        self.conn.close()

    # ------------------------------------------------------------ reads
    def state(self):
        with self.lock:
            return {"books": store.books(self.conn), "templates": [{"key": k, "name": v[0]} for k, v in TEMPLATES.items()]}

    def book(self, book_id):
        if book_id is None:
            raise SheetsError("Which workbook?")
        with self.lock:
            row = store.book_row(self.conn, book_id)
            store.opened(self.conn, book_id)
            return {"id": row["id"], "name": row["name"], "rev": row["rev"], "updated": row["updated"], "doc": json.loads(row["doc"])}

    def export(self, book_id, fmt, sheet_index=0):
        """(bytes, file name, content type) for a download."""
        with self.lock:
            row = store.book_row(self.conn, book_id)
        doc = files.names_as_rows(json.loads(row["doc"]))   # column names become a real header row in the file
        safe = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", row["name"]).strip() or "workbook"
        if fmt == "xlsx":
            return files.write_xlsx(doc), f"{safe}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        if fmt in ("csv", "tsv"):
            try:
                sheet = doc["sheets"][int(sheet_index)]
            except (ValueError, IndexError, TypeError):
                raise SheetsError("That sheet isn't in the workbook") from None
            text = files.write_csv(sheet, "," if fmt == "csv" else "\t")
            name = safe if len(doc["sheets"]) == 1 else f"{safe} - {sheet['name']}"
            return ("﻿" + text).encode("utf-8"), f"{name}.{fmt}", f"text/{'csv' if fmt == 'csv' else 'tab-separated-values'}; charset=utf-8"
        if fmt == "json":
            return json.dumps(doc, indent=1).encode(), f"{safe}.json", "application/json"
        raise SheetsError("Export as xlsx, csv or tsv")

    # ------------------------------------------------------------ writes
    def do(self, action, data):
        handler = getattr(self, f"_do_{action}", None) if isinstance(action, str) and action.isidentifier() else None
        if not handler:
            raise SheetsError("Unknown action")
        return handler(data)

    @staticmethod
    def _id(data, key="id"):
        value = data.get(key)
        if isinstance(value, int) and not isinstance(value, bool):
            return value
        if isinstance(value, str) and value.isdigit():
            return int(value)
        raise SheetsError("Bad request")

    def _do_new(self, data):
        template = data.get("template")
        if template and template not in TEMPLATES:
            raise SheetsError("Unknown template")
        name = data.get("name") or (TEMPLATES[template][0] if template else self._untitled())
        with self.lock:
            return {"id": store.add_book(self.conn, name, TEMPLATES[template][1]() if template else None)}

    def _untitled(self):
        with self.lock:
            names = {b["name"] for b in store.books(self.conn)}
        n = 1
        while f"Untitled spreadsheet{'' if n == 1 else f' {n}'}" in names:
            n += 1
        return f"Untitled spreadsheet{'' if n == 1 else f' {n}'}"

    def _do_save(self, data):
        try:
            rev = int(data.get("rev"))
        except (TypeError, ValueError):
            raise SheetsError("Bad request") from None
        with self.lock:
            return {"rev": store.save(self.conn, self._id(data), rev, data.get("doc"))}

    def _do_rename(self, data):
        with self.lock:
            store.rename(self.conn, self._id(data), data.get("name"))
        return {"ok": True}

    def _do_delete(self, data):
        with self.lock:
            store.delete(self.conn, self._id(data))
        return {"ok": True}

    def _do_duplicate(self, data):
        with self.lock:
            row = store.book_row(self.conn, self._id(data))
            return {"id": store.add_book(self.conn, f"Copy of {row['name']}"[:100], json.loads(row["doc"]))}

    def _do_versions(self, data):
        with self.lock:
            return {"versions": store.versions(self.conn, self._id(data))}

    def _do_restore(self, data):
        with self.lock:
            return {"rev": store.restore(self.conn, self._id(data), self._id(data, "version"))}

    def import_file(self, path, name=None):
        """Makes a workbook from an uploaded .xlsx, .csv or .tsv file."""
        path = Path(path)
        if path.stat().st_size > MAX_IMPORT:
            raise SheetsError("That file is over 50 MB")
        data = path.read_bytes()
        stem = (name or path.stem).rsplit(".", 1)[0] if name else path.stem
        suffix = Path(name or path.name).suffix.lower()
        if suffix in (".xlsx", ".xlsm") or data[:2] == b"PK":
            doc = files.read_xlsx(data)
        elif suffix in (".csv", ".tsv", ".txt"):
            doc = files.read_csv(data, stem[:100] or "Sheet1")
        elif suffix in (".xls", ".ods", ".numbers"):
            raise SheetsError(f"{suffix} files can't be read here. Save it as .xlsx or .csv first (File, Save As or Download).")
        else:
            raise SheetsError("Choose an .xlsx, .csv or .tsv file")
        with self.lock:
            return {"id": store.add_book(self.conn, stem[:100] or "Imported", doc)}
