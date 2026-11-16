import io
import json
import os
import shutil
import stat
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
from sheets import files, store  # noqa: E402
from sheets.store import SheetsError  # noqa: E402
from sheets.webapi import SheetsApi  # noqa: E402

FORMULA_JS = Path(os.environ.get("CAPTURE_DIR") or Path.home() / "capture-app") / "capture_app" / "static" / "sheets" / "formula.js"
NODE = shutil.which("node") or str(Path.home() / ".local/node/bin/node")


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.api = SheetsApi(os.path.join(self.tmp.name, "sheets.db"))
        self.addCleanup(self.api.close)


class FormulaEngine(unittest.TestCase):
    @unittest.skipUnless(os.path.exists(NODE) and FORMULA_JS.exists(), "node or formula.js missing")
    def test_formula_engine(self):
        """The 500+ checks of formulas, functions, formats and typed input in tests/formula_test.js."""
        out = subprocess.run([NODE, str(HERE / "formula_test.js"), str(FORMULA_JS)], capture_output=True, text=True, timeout=120)
        self.assertEqual(out.returncode, 0, out.stdout[-3000:] + out.stderr[-2000:])
        self.assertIn("0 failed", out.stdout)


class Books(Base):
    def test_new_save_conflict_and_versions(self):
        bid = self.api.do("new", {})["id"]
        book = self.api.book(bid)
        self.assertEqual((book["name"], book["rev"], book["doc"]["sheets"][0]["name"]), ("Untitled spreadsheet", 1, "Sheet1"))
        self.assertEqual(self.api.do("new", {})["id"], bid + 1)
        self.assertEqual(self.api.state()["books"][0]["name"], "Untitled spreadsheet 2")
        doc = book["doc"]
        doc["sheets"][0]["cells"]["A1"] = {"v": 5}
        doc["sheets"][0]["cells"]["B1"] = {"v": "=A1*2", "c": 10}
        rev = self.api.do("save", {"id": bid, "rev": 1, "doc": doc})["rev"]
        self.assertEqual(rev, 2)
        with self.assertRaisesRegex(SheetsError, "another tab"):
            self.api.do("save", {"id": bid, "rev": 1, "doc": doc})
        self.assertEqual(len(self.api.do("versions", {"id": bid})["versions"]), 1)
        doc["sheets"][0]["cells"]["A1"] = {"v": 6}
        self.api.do("save", {"id": bid, "rev": 2, "doc": doc})  # within 10 minutes: no new snapshot
        versions = self.api.do("versions", {"id": bid})["versions"]
        self.assertEqual(len(versions), 1)
        self.api.do("restore", {"id": bid, "version": versions[0]["id"]})
        self.assertEqual(self.api.book(bid)["doc"]["sheets"][0]["cells"], {})
        self.assertEqual(len(self.api.do("versions", {"id": bid})["versions"]), 2)  # what was there before the restore
        self.assertEqual(stat.S_IMODE(os.stat(os.path.join(self.tmp.name, "sheets.db")).st_mode), 0o600)

    def test_bad_workbooks_are_refused(self):
        bid = self.api.do("new", {})["id"]
        good = store.blank_doc()
        for bad in ({}, {"sheets": []}, {"sheets": [{"name": "A", "cells": {"A0": {"v": 1}}}]}, {"sheets": [{"name": "A", "cells": {"A1": {"v": [1]}}}]},
                    {"sheets": [{"name": "A", "cells": {}}, {"name": "a", "cells": {}}]}, {"sheets": [{"name": "A/B", "cells": {}}]}, {"sheets": [{"name": " ", "cells": {}}]}):
            with self.assertRaises(SheetsError, msg=bad):
                self.api.do("save", {"id": bid, "rev": 1, "doc": bad})
        self.assertEqual(self.api.do("save", {"id": bid, "rev": 1, "doc": good})["rev"], 1)  # unchanged: same revision

    def test_rename_duplicate_delete_templates(self):
        for t in ("budget", "todo", "loan", "workout"):
            bid = self.api.do("new", {"template": t})["id"]
            self.assertTrue(self.api.book(bid)["doc"]["sheets"][0]["cells"])
        self.api.do("rename", {"id": bid, "name": "  My   lifts "})
        self.assertEqual(self.api.book(bid)["name"], "My lifts")
        copy = self.api.do("duplicate", {"id": bid})["id"]
        self.assertEqual(self.api.book(copy)["name"], "Copy of My lifts")
        self.api.do("delete", {"id": copy})
        with self.assertRaises(SheetsError):
            self.api.book(copy)
        with self.assertRaises(SheetsError):
            self.api.do("new", {"template": "nope"})


class Files(Base):
    def doc(self):
        sheet = store.blank_sheet("Money")
        sheet["cells"] = {"A1": {"v": "Item", "s": {"b": 1, "bg": "#e8eaed"}}, "B1": {"v": "Cost", "s": {"b": 1}},
                          "A2": {"v": "Rent & <bills>"}, "B2": {"v": 1200.5, "s": {"nf": "$#,##0.00", "fc": "#ff0000"}},
                          "A3": {"v": "'=not a formula"}, "B3": {"v": 0.25, "s": {"nf": "0%"}}, "C1": {"v": 45366, "s": {"nf": "yyyy-mm-dd", "ha": "center", "wr": 1}},
                          "B4": {"v": "=SUM(B2:B3)", "c": 1200.75, "s": {"nf": "$#,##0.00", "bd": {"t": 1, "b": 1}}},
                          "C4": {"v": '=XLOOKUP("Rent",A2:A3,B2:B3,"none")', "c": "none"}, "D1": {"v": True}, "D2": {"v": "=1/0", "c": {"e": "#DIV/0!"}}}
        sheet["colW"] = {"0": 180}
        sheet["rowH"] = {"0": 30}
        sheet["freeze"] = {"r": 1, "c": 0}
        second = store.blank_sheet("Other sheet")
        second["cells"] = {"A1": {"v": "='Money'!B4*2", "c": 2401.5}}
        return {"sheets": [sheet, second], "active": 0}

    def test_xlsx_round_trip(self):
        data = files.write_xlsx(self.doc())
        z = zipfile.ZipFile(io.BytesIO(data))
        self.assertIn("xl/worksheets/sheet2.xml", z.namelist())
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode()
        self.assertIn("<f>_xlfn.XLOOKUP(", sheet_xml)  # newer functions get Excel's prefix
        self.assertIn('t="e"', sheet_xml)
        self.assertIn('fullCalcOnLoad="1"', z.read("xl/workbook.xml").decode())
        back = files.read_xlsx(data)
        a, b = back["sheets"]
        self.assertEqual((a["name"], b["name"]), ("Money", "Other sheet"))
        self.assertEqual(a["cells"]["A2"]["v"], "Rent & <bills>")
        self.assertEqual(a["cells"]["A3"]["v"], "'=not a formula")
        self.assertEqual(a["cells"]["B2"]["v"], 1200.5)
        self.assertEqual(a["cells"]["B2"]["s"], {"nf": "$#,##0.00", "fc": "#ff0000"})
        self.assertEqual(a["cells"]["A1"]["s"], {"b": 1, "bg": "#e8eaed"})
        self.assertEqual(a["cells"]["B4"]["v"], "=SUM(B2:B3)")
        self.assertEqual(a["cells"]["B4"]["c"], 1200.75)
        self.assertEqual(a["cells"]["B4"]["s"]["bd"], {"t": 1, "b": 1})
        self.assertEqual(a["cells"]["C4"]["v"], '=XLOOKUP("Rent",A2:A3,B2:B3,"none")')
        self.assertEqual(a["cells"]["C1"]["s"], {"nf": "yyyy-mm-dd", "ha": "center", "wr": 1})
        self.assertIs(a["cells"]["D1"]["v"], True)
        self.assertEqual(a["cells"]["D2"]["c"], {"e": "#DIV/0!"})
        self.assertEqual(a["colW"]["0"], 180)
        self.assertEqual(a["rowH"]["0"], 30)
        self.assertEqual(a["freeze"], {"r": 1, "c": 0})
        self.assertEqual(b["cells"]["A1"]["v"], "='Money'!B4*2")
        store.clean_doc(back)  # what comes back is always savable

    def test_shared_formulas_from_excel(self):
        sheet = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
                 '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2+$C$1</f><v>2</v></c></row>'
                 '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>'
                 '<row r="3"><c r="A3" t="s"><v>0</v></c><c r="B3"><f t="shared" si="0"/></c></row></sheetData></worksheet>')
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            z.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>')
            z.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="x"/></Relationships>')
            z.writestr("xl/worksheets/sheet1.xml", sheet)
            z.writestr("xl/sharedStrings.xml", '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t>Hello </t></r><r><t>world</t></r></si></sst>')
        doc = files.read_xlsx(buf.getvalue())
        cells = doc["sheets"][0]["cells"]
        self.assertEqual(cells["B2"]["v"], "=A2*2+$C$1")
        self.assertEqual(cells["B3"]["v"], "=A3*2+$C$1")
        self.assertEqual(cells["A3"]["v"], "Hello world")
        self.assertEqual(files.shift_formula('=A1&"A1"&Sheet2!B$2', 1, 1), '=B2&"A1"&Sheet2!C$2')

    def test_bad_files(self):
        with self.assertRaisesRegex(SheetsError, "valid .xlsx"):
            files.read_xlsx(b"PK nothing")
        path = Path(self.tmp.name) / "old.xls"
        path.write_bytes(b"\xd0\xcf\x11\xe0")
        with self.assertRaisesRegex(SheetsError, "Save it as .xlsx"):
            self.api.import_file(path, "old.xls")

    def test_csv_in_and_out(self):
        path = Path(self.tmp.name) / "bank.csv"
        path.write_bytes("Date;Amount;Note\n2024-03-05;-12.50;Café\n3/6/2024;1,000;\"has; semicolon\"\n;;=HYPERLINK(\"x\")\n".encode("cp1252"))
        bid = self.api.import_file(path, "bank.csv")["id"]
        book = self.api.book(bid)
        cells = book["doc"]["sheets"][0]["cells"]
        self.assertEqual(book["name"], "bank")
        self.assertEqual(cells["A2"], {"v": 45356, "s": {"nf": "yyyy-mm-dd"}})
        self.assertEqual(cells["B2"]["v"], -12.5)
        self.assertEqual(cells["C2"]["v"], "Café")
        self.assertEqual(cells["B3"], {"v": 1000, "s": {"nf": "#,##0"}})
        self.assertEqual(cells["C3"]["v"], "has; semicolon")
        self.assertEqual(cells["C4"]["v"], "'=HYPERLINK(\"x\")")  # never run as a formula
        data, name, kind = self.api.export(bid, "csv", 0)
        text = data.decode("utf-8-sig")
        book = self.api.book(bid)
        self.assertTrue(text.startswith("Date,Amount,Note\r\n"))
        self.assertIn("has; semicolon", text)
        # cells the page has shown keep their formatted text, so dates and money export the way they look
        doc = book["doc"]
        doc["sheets"][0]["cells"]["A2"]["t"] = "2024-03-05"
        self.api.do("save", {"id": bid, "rev": book["rev"], "doc": doc})
        self.assertIn("2024-03-05,-12.5", self.api.export(bid, "csv", 0)[0].decode("utf-8-sig"))
        self.assertEqual(name, "bank.csv")
        sheet = {"cells": {"A1": {"v": "=SEQUENCE(3)", "c": 1}, "B1": {"v": "=1/0", "c": {"e": "#DIV/0!"}}}, "spilled": {"A2": 2, "A3": 3}}
        self.assertEqual(files.write_csv(sheet), "1,#DIV/0!\r\n2\r\n3\r\n")
        with self.assertRaises(SheetsError):
            self.api.export(bid, "pdf")


if __name__ == "__main__":
    unittest.main()
