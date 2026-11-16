"""Excel (.xlsx) and CSV in and out, with only Python's standard library.

.xlsx is a zip of XML files (the Office Open XML format). Reading keeps values, formulas (shared formulas are expanded),
bold/italic/underline/strike, font size and colour, fill colour, alignment, wrap, number formats, column widths, row
heights and frozen panes. Writing does the same, with each formula's last result so Excel and Google Sheets show values
straight away, and asks Excel to recalculate when it opens the file.
"""
import csv
import datetime as dt
import io
import posixpath
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from .store import SheetsError, blank_sheet

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
MAX_UNZIPPED = 300 * 1024 * 1024
MAX_IMPORT_CELLS = 2_000_000
BUILTIN_FORMATS = {0: None, 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 5: "$#,##0;($#,##0)", 6: "$#,##0;[Red]($#,##0)", 7: "$#,##0.00;($#,##0.00)",
                   8: "$#,##0.00;[Red]($#,##0.00)", 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??", 14: "m/d/yyyy", 15: "d-mmm-yy",
                   16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm",
                   37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss",
                   46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@"}
# Functions that newer Excel only understands with this prefix in a file (without it they show #NAME?).
XLFN = ["XLOOKUP", "XMATCH", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "LET", "IFS", "SWITCH", "MAXIFS", "MINIFS", "CONCAT",
        "TEXTJOIN", "XOR", "IFNA", "DAYS", "ISOWEEKNUM", "STDEV.S", "STDEV.P", "VAR.S", "VAR.P", "MODE.SNGL", "RANK.EQ", "RANK.AVG", "PERCENTILE.INC",
        "PERCENTILE.EXC", "QUARTILE.INC", "NORM.DIST", "NORM.S.DIST", "NORM.INV", "NORM.S.INV", "CEILING.MATH", "FLOOR.MATH", "COVARIANCE.P",
        "COVARIANCE.S", "FORECAST.LINEAR", "VSTACK", "HSTACK", "TOCOL", "TOROW", "CHOOSECOLS", "CHOOSEROWS", "TAKE", "DROP", "TEXTSPLIT",
        "TEXTBEFORE", "TEXTAFTER", "SHEET", "SHEETS", "FORMULATEXT", "ISFORMULA", "NUMBERVALUE", "UNICHAR", "UNICODE", "RRI", "BASE", "DECIMAL",
        "ARRAYTOTEXT", "COMBINA", "ERROR.TYPE"]
XLFN_RE = re.compile(r'("(?:[^"]|"")*")|\b(' + "|".join(re.escape(n) for n in sorted(XLFN, key=len, reverse=True)) + r')(?=\()', re.I)


def col_name(c):
    s = ""
    c += 1
    while c:
        c, m = divmod(c - 1, 26)
        s = chr(65 + m) + s
    return s


def col_index(name):
    n = 0
    for ch in name.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def split_addr(a):
    m = re.match(r"^\$?([A-Za-z]{1,3})\$?(\d+)$", a)
    return (int(m[2]) - 1, col_index(m[1])) if m else None


# ---------------------------------------------------------------- reading typed text (CSV cells)
NUM = re.compile(r"^([+-])?(\$)?((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(%)?$")


def parse_text(text):
    """A CSV value as (value, number format): numbers, money, percents and ISO or US dates, like typing it in."""
    t = text.strip()
    if not t:
        return None, None
    m = NUM.match(t)
    if m:
        v = float(m[3].replace(",", ""))
        v = -v if m[1] == "-" else v
        v = int(v) if v.is_integer() and "." not in m[3] and not m[4] else v
        if m[4]:
            return v / 100, "0%"
        if m[2]:
            return v, "$#,##0.00" if "." in m[3] else "$#,##0"
        return v, ("#,##0.00" if "." in m[3] else "#,##0") if "," in m[3] else None
    for pattern, fmt in (("%Y-%m-%d", "yyyy-mm-dd"), ("%m/%d/%Y", "m/d/yyyy"), ("%Y-%m-%d %H:%M:%S", "yyyy-mm-dd h:mm"), ("%Y-%m-%d %H:%M", "yyyy-mm-dd h:mm")):
        try:
            d = dt.datetime.strptime(t, pattern)
        except ValueError:
            continue
        return round((d - dt.datetime(1899, 12, 30)).total_seconds() / 86400, 10), fmt
    if t.upper() in ("TRUE", "FALSE"):
        return t.upper() == "TRUE", None
    if t.startswith("="):  # a CSV value is never run as a formula unless you type it in yourself
        return "'" + t, None
    return text, None


def read_csv(data, name="Sheet1"):
    text = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    sample = text[:20000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel_tab if sample.count("\t") > sample.count(",") else csv.excel
    sheet = blank_sheet(name[:100] or "Sheet1")
    rows = list(csv.reader(io.StringIO(text), dialect))
    count = 0
    for r, row in enumerate(rows):
        for c, raw in enumerate(row):
            if c >= 16384:
                break
            v, fmt = parse_text(raw)
            if v is None:
                continue
            cell = {"v": v}
            if fmt:
                cell["s"] = {"nf": fmt}
            sheet["cells"][col_name(c) + str(r + 1)] = cell
            count += 1
            if count > MAX_IMPORT_CELLS:
                raise SheetsError("That file has more than 2 million filled cells")
    sheet["rows"] = max(1000, len(rows) + 100)
    sheet["cols"] = max(26, max((len(r) for r in rows), default=0) + 2)
    return {"sheets": [sheet], "active": 0}


def shown_value(cell):
    if isinstance(cell.get("t"), str):
        return cell["t"]
    v = cell.get("c", cell.get("v")) if isinstance(cell.get("v"), str) and cell["v"].startswith("=") else cell.get("v")
    if isinstance(v, dict):  # a cached error
        return v.get("e", "")
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, str) and v.startswith("'"):
        return v[1:]
    return "" if v is None else str(v)


def write_csv(sheet, delimiter=","):
    """A sheet's shown values (formulas as their results), as CSV text. Spilled results are included."""
    cells = dict(sheet.get("cells", {}))
    spilled = sheet.get("spilled") or {}
    grid = {}
    for key, cell in cells.items():
        p = split_addr(key)
        if p:
            grid[p] = shown_value(cell)
    for key, v in spilled.items():
        p = split_addr(key)
        if p and p not in grid:
            grid[p] = shown_value({"v": v})
    if not grid:
        return ""
    rows = max(p[0] for p in grid) + 1
    cols = max(p[1] for p in grid) + 1
    out = io.StringIO()
    w = csv.writer(out, delimiter=delimiter, lineterminator="\r\n")
    for r in range(rows):
        row = [grid.get((r, c), "") for c in range(cols)]
        while row and row[-1] == "":
            row.pop()
        w.writerow(row)
    return out.getvalue()


# ---------------------------------------------------------------- reading .xlsx
def _text(el):
    """All the text inside a shared-string or inline-string item, rich-text runs included."""
    if el is None:
        return ""
    # phonetic runs (rPh), which Excel adds for Japanese text, aren't part of the value
    rph = {id(t) for p in el.iter(f"{{{NS['m']}}}rPh") for t in p.iter(f"{{{NS['m']}}}t")}
    return "".join(t.text or "" for t in el.iter(f"{{{NS['m']}}}t") if id(t) not in rph)


def _color(el):
    if el is None:
        return None
    rgb = el.get("rgb")
    if rgb and len(rgb) == 8:
        return "#" + rgb[2:].lower()
    if rgb and len(rgb) == 6:
        return "#" + rgb.lower()
    return None  # theme and indexed colours are left as default


def shift_formula(text, dr, dc):
    """Moves the relative references in a formula, for Excel's shared formulas."""
    def move(m):
        if m.group(1):
            return m.group(1)
        sheet, ac, col, ar, row = m.group(2) or "", m.group(3), m.group(4), m.group(5), m.group(6)
        c = col_index(col) + (0 if ac else dc)
        r = int(row) - 1 + (0 if ar else dr)
        if c < 0 or r < 0:
            return "#REF!"
        return f"{sheet}{ac}{col_name(c)}{ar}{r + 1}"
    return re.sub(r"(\"(?:[^\"]|\"\")*\")|((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(])", move, text)


def _sheet_key(name):
    return name.strip("'").replace("''", "'").lower()


def _move_down(text, home, named):
    """A formula with every reference into a sheet that has column names moved one row down (absolute ones too),
    because on export those names become row 1. `home` is the formula's own sheet."""
    def move(m):
        if m.group(1):
            return m.group(1)
        sheet, ac, col, ar, row = m.group(2) or "", m.group(3), m.group(4), m.group(5), m.group(6)
        target = _sheet_key(sheet[:-1]) if sheet else home.lower()
        if target not in named:
            return m.group(0)
        return f"{sheet}{ac}{col}{ar}{int(row) + 1}"
    return re.sub(r"(\"(?:[^\"]|\"\")*\")|((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(])", move, text)


def names_as_rows(doc):
    """A copy of the workbook where each sheet's column names are a real first row (its cells one row down, and
    formulas pointing at them moved to match), so exported files open with headers in Excel and Google Sheets."""
    named = {s["name"].lower() for s in doc["sheets"] if s.get("names")}
    if not named:
        return doc
    out = dict(doc, sheets=[])
    for sheet in doc["sheets"]:
        has = sheet["name"].lower() in named
        shift = 1 if has else 0
        cells = {}
        for key, cell in (sheet.get("cells") or {}).items():
            p = split_addr(key)
            if not p:
                continue
            cell = dict(cell)
            if isinstance(cell.get("v"), str) and cell["v"].startswith("="):
                cell["v"] = _move_down(cell["v"], sheet["name"], named)
            cells[col_name(p[1]) + str(p[0] + 1 + shift)] = cell
        new = dict(sheet, cells=cells)
        new.pop("names", None)
        if has:
            for c, name in sheet["names"].items():
                cells[col_name(int(c)) + "1"] = {"v": name, "s": {"b": 1}}
            new["spilled"] = {col_name(p[1]) + str(p[0] + 2): v for k, v in (sheet.get("spilled") or {}).items() if (p := split_addr(k))}
            new["rowH"] = {str(int(k) + 1): v for k, v in (sheet.get("rowH") or {}).items()}
            new["freeze"] = dict(sheet.get("freeze") or {}, r=((sheet.get("freeze") or {}).get("r") or 0) + 1)   # the header stays in view
            new["rows"] = (sheet.get("rows") or 1000) + 1
            down = lambda g: {**g, "r1": g["r1"] + 1, "r2": min(1_048_575, g["r2"] + 1)}  # noqa: E731
            new["merges"] = [down(m) for m in sheet.get("merges") or []]
            new["lists"] = [down(rule) for rule in sheet.get("lists") or []]
        out["sheets"].append(new)
    return out


def read_xlsx(data):
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        raise SheetsError("That isn't a valid .xlsx file") from None
    if sum(i.file_size for i in z.infolist()) > MAX_UNZIPPED:
        raise SheetsError("That .xlsx file is too big once unpacked")
    names = set(z.namelist())

    def xml(path):
        return ET.fromstring(z.read(path)) if path in names else None

    wb = xml("xl/workbook.xml")
    if wb is None:
        raise SheetsError("That .xlsx file has no workbook inside")
    rels = {}
    rel_root = xml("xl/_rels/workbook.xml.rels")
    if rel_root is not None:
        for rel in rel_root.findall("rel:Relationship", NS):
            target = rel.get("Target", "")
            rels[rel.get("Id")] = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
    shared = []
    ss = xml("xl/sharedStrings.xml")
    if ss is not None:
        shared = [_text(si) for si in ss.findall("m:si", NS)]
    styles = _read_styles(xml("xl/styles.xml"))
    doc = {"sheets": [], "active": 0}
    total = 0
    for sh in wb.findall("m:sheets/m:sheet", NS):
        path = rels.get(sh.get(f"{{{NS['r']}}}id"))
        root = xml(path) if path else None
        if root is None:
            continue
        sheet = blank_sheet((sh.get("name") or f"Sheet{len(doc['sheets']) + 1}")[:100])
        if sh.get("state") in ("hidden", "veryHidden"):
            sheet["hidden"] = True
        masters = {}
        max_r = max_c = 0
        for row in root.iterfind("m:sheetData/m:row", NS):
            ht = row.get("ht")
            if ht and row.get("customHeight") in ("1", "true") and row.get("r"):
                sheet["rowH"][str(int(row.get("r")) - 1)] = round(float(ht) * 4 / 3)
            for c in row.iterfind("m:c", NS):
                ref = c.get("r")
                pos = split_addr(ref) if ref else None
                if not pos:
                    continue
                t = c.get("t", "n")
                f = c.find("m:f", NS)
                v = c.find("m:v", NS)
                value = None
                if t == "s" and v is not None:
                    i = int(v.text or 0)
                    value = shared[i] if i < len(shared) else ""
                elif t == "inlineStr":
                    value = _text(c.find("m:is", NS))
                elif t == "b" and v is not None:
                    value = v.text == "1"
                elif t in ("str", "e") and v is not None:
                    value = v.text or ""
                elif v is not None and v.text is not None:
                    try:
                        n = float(v.text)
                        value = int(n) if n.is_integer() and abs(n) < 1e15 else n
                    except ValueError:
                        value = v.text
                cell = {}
                formula = None
                if f is not None:
                    ftext = f.text
                    if f.get("t") == "shared":
                        si = f.get("si")
                        if ftext:
                            masters[si] = (ftext, pos)
                        elif si in masters:
                            base, (mr, mc) = masters[si]
                            ftext = shift_formula(base, pos[0] - mr, pos[1] - mc)
                    if ftext:
                        formula = "=" + re.sub(r"_xlfn\.|_xlws\.|_xlpm\.", "", ftext)
                if formula:
                    cell["v"] = formula
                    if value is not None:
                        cell["c"] = {"e": value} if t == "e" else value
                elif value is not None and value != "":
                    cell["v"] = ("'" + value) if isinstance(value, str) and (value.startswith("=") or value.startswith("'")) else value
                style = styles[int(c.get("s", 0))] if styles and int(c.get("s", 0)) < len(styles) else None
                if style:
                    cell["s"] = dict(style)
                if cell.get("v") is None and not cell.get("s"):
                    continue
                if cell.get("v") is None and cell.get("s"):
                    cell["v"] = None
                sheet["cells"][col_name(pos[1]) + str(pos[0] + 1)] = {k: v2 for k, v2 in cell.items() if v2 is not None}
                max_r, max_c = max(max_r, pos[0]), max(max_c, pos[1])
                total += 1
                if total > MAX_IMPORT_CELLS:
                    raise SheetsError("That file has more than 2 million filled cells")
        for col in root.iterfind("m:cols/m:col", NS):
            width = col.get("width")
            if width:
                lo, hi = int(col.get("min", 1)), min(int(col.get("max", 1)), 300)
                for ci in range(lo - 1, hi):
                    sheet["colW"][str(ci)] = max(16, round(float(width) * 7 + 5))
                if col.get("hidden") in ("1", "true"):
                    for ci in range(lo - 1, hi):
                        sheet["colW"][str(ci)] = 2
        pane = root.find("m:sheetViews/m:sheetView/m:pane", NS)
        if pane is not None and pane.get("state") in ("frozen", "frozenSplit"):
            sheet["freeze"] = {"r": min(50, int(float(pane.get("ySplit", 0)))), "c": min(26, int(float(pane.get("xSplit", 0))))}
        sheet["rows"] = max(1000, max_r + 100)
        sheet["cols"] = max(26, max_c + 3)
        for mc in root.iterfind("m:mergeCells/m:mergeCell", NS):
            parts = [split_addr(x) for x in (mc.get("ref") or "").split(":")]
            if len(parts) == 2 and all(parts):
                sheet.setdefault("merges", []).append({"r1": parts[0][0], "c1": parts[0][1], "r2": parts[1][0], "c2": parts[1][1]})
        for dvel in root.iterfind("m:dataValidations/m:dataValidation", NS):
            f1 = dvel.find("m:formula1", NS)
            if dvel.get("type") != "list" or f1 is None or not (f1.text or "").startswith('"'):
                continue
            choices = [x.strip() for x in f1.text.strip('"').split(",") if x.strip()]
            for ref in (dvel.get("sqref") or "").split():
                parts = [split_addr(x) for x in ref.split(":")]
                parts = parts * 2 if len(parts) == 1 else parts
                if len(parts) == 2 and all(parts) and choices:
                    sheet.setdefault("lists", []).append({"r1": parts[0][0], "c1": parts[0][1], "r2": parts[1][0], "c2": parts[1][1],
                                                          "options": [{"v": x, "color": None} for x in choices[:500]],
                                                          "strict": dvel.get("showErrorMessage") in ("1", "true")})
        doc["sheets"].append(sheet)
    if not doc["sheets"]:
        raise SheetsError("That .xlsx file has no sheets")
    return doc


def _read_styles(root):
    if root is None:
        return []
    custom = {int(n.get("numFmtId")): n.get("formatCode") for n in root.iterfind("m:numFmts/m:numFmt", NS)}
    fonts = []
    for f in root.iterfind("m:fonts/m:font", NS):
        s = {}
        if f.find("m:b", NS) is not None and f.find("m:b", NS).get("val", "1") not in ("0", "false"):
            s["b"] = 1
        if f.find("m:i", NS) is not None and f.find("m:i", NS).get("val", "1") not in ("0", "false"):
            s["i"] = 1
        if f.find("m:u", NS) is not None and f.find("m:u", NS).get("val", "single") != "none":
            s["u"] = 1
        if f.find("m:strike", NS) is not None and f.find("m:strike", NS).get("val", "1") not in ("0", "false"):
            s["st"] = 1
        sz = f.find("m:sz", NS)
        if sz is not None and sz.get("val") and abs(float(sz.get("val")) - 11) > 0.1:
            s["fs"] = round(float(sz.get("val")))
        color = _color(f.find("m:color", NS))
        if color and color not in ("#000000",):
            s["fc"] = color
        fonts.append(s)
    fills = []
    for fill in root.iterfind("m:fills/m:fill", NS):
        p = fill.find("m:patternFill", NS)
        color = _color(p.find("m:fgColor", NS)) if p is not None and p.get("patternType") not in (None, "none", "gray125") else None
        fills.append({"bg": color} if color and color != "#ffffff" else {})
    borders = []
    for b in root.iterfind("m:borders/m:border", NS):
        s = {}
        for side, key in (("top", "t"), ("right", "r"), ("bottom", "b"), ("left", "l")):
            e = b.find(f"m:{side}", NS)
            if e is not None and e.get("style"):
                s[key] = 1
        borders.append({"bd": s} if s else {})
    out = []
    for xf in root.iterfind("m:cellXfs/m:xf", NS):
        s = {}
        fid = int(xf.get("numFmtId", 0))
        code = custom.get(fid, BUILTIN_FORMATS.get(fid))
        if code and code.lower() != "general":
            s["nf"] = code
        for lst, key in ((fonts, "fontId"), (fills, "fillId"), (borders, "borderId")):
            i = int(xf.get(key, 0))
            if i < len(lst):
                s.update(lst[i])
        al = xf.find("m:alignment", NS)
        if al is not None:
            if al.get("horizontal") in ("left", "center", "right"):
                s["ha"] = al.get("horizontal")
            if al.get("vertical") in ("top", "center"):
                s["va"] = "middle" if al.get("vertical") == "center" else "top"
            if al.get("wrapText") in ("1", "true"):
                s["wr"] = 1
        out.append(s or None)
    return out


# ---------------------------------------------------------------- writing .xlsx
def _xlsx_formula(text):
    body = text[1:]
    return XLFN_RE.sub(lambda m: m.group(1) or "_xlfn." + m.group(2).upper(), body)


def _argb(color):
    c = (color or "").lstrip("#")
    if len(c) == 3:
        c = "".join(ch * 2 for ch in c)
    return "FF" + c.upper() if re.fullmatch(r"[0-9a-fA-F]{6}", c) else None


def write_xlsx(doc):
    fonts, fills, borders, numfmts, xfs = [("", None, 11)], ["none", "gray125"], [()], {}, [(0, 0, 0, 0, None)]
    font_ix, fill_ix, border_ix, xf_ix = {fonts[0]: 0}, {}, {(): 0}, {xfs[0]: 0}

    def style_index(s):
        if not s:
            return 0
        flags = "".join(k for k in ("b", "i", "u", "st") if s.get(k))
        font = (flags, _argb(s.get("fc")), int(s.get("fs") or 11))
        if font not in font_ix:
            font_ix[font] = len(fonts)
            fonts.append(font)
        fill = _argb(s.get("bg"))
        if fill and fill not in fill_ix:
            fill_ix[fill] = len(fills)
            fills.append(fill)
        bd = tuple(k for k in ("l", "r", "t", "b") if (s.get("bd") or {}).get(k))
        if bd not in border_ix:
            border_ix[bd] = len(borders)
            borders.append(bd)
        nf = s.get("nf")
        num_id = 0
        if nf:
            builtin = next((k for k, v in BUILTIN_FORMATS.items() if v == nf), None)
            if builtin is not None:
                num_id = builtin
            else:
                if nf not in numfmts:
                    numfmts[nf] = 164 + len(numfmts)
                num_id = numfmts[nf]
        align = (s.get("ha"), s.get("va"), bool(s.get("wr")))
        key = (num_id, font_ix[font], fill_ix.get(fill, 0), border_ix[bd], align if any(align) else None)
        if key not in xf_ix:
            xf_ix[key] = len(xfs)
            xfs.append(key)
        return xf_ix[key]

    sheets_xml = []
    for sheet in doc["sheets"]:
        rows = {}
        for key, cell in sheet.get("cells", {}).items():
            p = split_addr(key)
            if p:
                rows.setdefault(p[0], []).append((p[1], key, cell))
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
        fr = sheet.get("freeze") or {}
        if fr.get("r") or fr.get("c"):
            top = col_name(fr.get("c") or 0) + str((fr.get("r") or 0) + 1)
            split = (f'xSplit="{fr["c"]}" ' if fr.get("c") else "") + (f'ySplit="{fr["r"]}" ' if fr.get("r") else "")
            out.append(f'<sheetViews><sheetView workbookViewId="0"><pane {split}topLeftCell="{top}" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>')
        widths = sheet.get("colW") or {}
        if widths:
            out.append("<cols>" + "".join(f'<col min="{int(c) + 1}" max="{int(c) + 1}" width="{max(0.5, (w - 5) / 7):.2f}" customWidth="1"/>'
                                         for c, w in sorted(widths.items(), key=lambda kv: int(kv[0]))) + "</cols>")
        out.append("<sheetData>")
        heights = sheet.get("rowH") or {}
        for r in sorted(set(rows) | {int(k) for k in heights}):
            ht = heights.get(str(r))
            attrs = f' ht="{ht * 0.75:.2f}" customHeight="1"' if ht else ""
            out.append(f'<row r="{r + 1}"{attrs}>')
            for c, key, cell in sorted(rows.get(r, []), key=lambda x: x[0]):
                s = style_index(cell.get("s"))
                sattr = f' s="{s}"' if s else ""
                v = cell.get("v")
                if isinstance(v, str) and v.startswith("=") and len(v) > 1:
                    cached = cell.get("c")
                    f = f"<f>{escape(_xlsx_formula(v))}</f>"
                    if isinstance(cached, bool):
                        out.append(f'<c r="{key}"{sattr} t="b">{f}<v>{int(cached)}</v></c>')
                    elif isinstance(cached, (int, float)):
                        out.append(f'<c r="{key}"{sattr}>{f}<v>{repr(float(cached)) if isinstance(cached, float) else cached}</v></c>')
                    elif isinstance(cached, dict):
                        out.append(f'<c r="{key}"{sattr} t="e">{f}<v>{escape(str(cached.get("e", "#VALUE!")))}</v></c>')
                    elif isinstance(cached, str):
                        out.append(f'<c r="{key}"{sattr} t="str">{f}<v>{escape(cached)}</v></c>')
                    else:
                        out.append(f'<c r="{key}"{sattr}>{f}</c>')
                elif isinstance(v, bool):
                    out.append(f'<c r="{key}"{sattr} t="b"><v>{int(v)}</v></c>')
                elif isinstance(v, (int, float)):
                    out.append(f'<c r="{key}"{sattr}><v>{repr(float(v)) if isinstance(v, float) else v}</v></c>')
                elif isinstance(v, str) and v != "":
                    text = v[1:] if v.startswith("'") else v
                    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
                    out.append(f'<c r="{key}"{sattr} t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is></c>')
                elif s:
                    out.append(f'<c r="{key}"{sattr}/>')
            out.append("</row>")
        out.append("</sheetData>")
        merges = sheet.get("merges") or []
        if merges:
            out.append(f'<mergeCells count="{len(merges)}">' + "".join(
                f'<mergeCell ref="{col_name(m["c1"])}{m["r1"] + 1}:{col_name(m["c2"])}{m["r2"] + 1}"/>' for m in merges) + "</mergeCells>")
        rules = [rule for rule in sheet.get("lists") or [] if rule.get("options")]
        if rules:
            # Excel keeps a list's choices in one quoted, comma-separated text of up to 255 characters
            dv = []
            for rule in rules:
                choices = ",".join(str(o["v"]).replace('"', "'").replace(",", " ") for o in rule["options"])[:255]
                ref = f'{col_name(rule["c1"])}{rule["r1"] + 1}:{col_name(rule["c2"])}{min(rule["r2"], 1_048_575) + 1}'
                dv.append(f'<dataValidation type="list" allowBlank="1" showErrorMessage="{1 if rule.get("strict") else 0}" sqref="{ref}">'
                          f'<formula1>{escape(chr(34) + choices + chr(34))}</formula1></dataValidation>')
            out.append(f'<dataValidations count="{len(dv)}">' + "".join(dv) + "</dataValidations>")
        out.append("</worksheet>")
        sheets_xml.append("".join(out))

    def font_xml(f):
        flags, color, size = f  # flags is made of b, i, u and st
        inner = ("<b/>" if "b" in flags else "") + ("<i/>" if "i" in flags else "") + ("<strike/>" if "st" in flags else "") + ("<u/>" if "u" in flags else "")
        return f'<font>{inner}<sz val="{size}"/>{f"<color rgb={chr(34)}{color}{chr(34)}/>" if color else ""}<name val="Calibri"/></font>'

    def fill_xml(f):
        if f in ("none", "gray125"):
            return f'<fill><patternFill patternType="{f}"/></fill>'
        return f'<fill><patternFill patternType="solid"><fgColor rgb="{f}"/><bgColor indexed="64"/></patternFill></fill>'

    def border_xml(b):
        side = lambda name, key: f'<{name} style="thin"><color auto="1"/></{name}>' if key in b else f"<{name}/>"  # noqa: E731
        return "<border>" + side("left", "l") + side("right", "r") + side("top", "t") + side("bottom", "b") + "<diagonal/></border>"

    def xf_xml(x):
        num, font, fill, border, align = x
        attrs = f'numFmtId="{num}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0"'
        attrs += (' applyNumberFormat="1"' if num else "") + (' applyFont="1"' if font else "") + (' applyFill="1"' if fill else "") + (' applyBorder="1"' if border else "")
        if align:
            ha, va, wr = align
            a = (f' horizontal="{ha}"' if ha else "") + (f' vertical="{"center" if va == "middle" else va}"' if va else "") + (' wrapText="1"' if wr else "")
            return f'<xf {attrs} applyAlignment="1"><alignment{a}/></xf>'
        return f"<xf {attrs}/>"

    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              + (f'<numFmts count="{len(numfmts)}">' + "".join(f'<numFmt numFmtId="{i}" formatCode="{escape(code, {chr(34): "&quot;"})}"/>' for code, i in numfmts.items()) + "</numFmts>" if numfmts else "")
              + f'<fonts count="{len(fonts)}">' + "".join(font_xml(f) for f in fonts) + "</fonts>"
              + f'<fills count="{len(fills)}">' + "".join(fill_xml(f) for f in fills) + "</fills>"
              + f'<borders count="{len(borders)}">' + "".join(border_xml(b) for b in borders) + "</borders>"
              + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              + f'<cellXfs count="{len(xfs)}">' + "".join(xf_xml(x) for x in xfs) + "</cellXfs>"
              + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    n = len(sheets_xml)
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                + "".join(f'<sheet name="{escape(s["name"], {chr(34): "&quot;"})}" sheetId="{i + 1}" r:id="rId{i + 1}"'
                          f'{" state=" + chr(34) + "hidden" + chr(34) if s.get("hidden") else ""}/>' for i, s in enumerate(doc["sheets"]))
                + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               + "".join(f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(n))
               + f'<Relationship Id="rId{n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
             '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
             '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
             + "".join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(n))
             + "</Types>")
    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles)
        for i, text in enumerate(sheets_xml):
            z.writestr(f"xl/worksheets/sheet{i + 1}.xml", text)
    return buf.getvalue()
