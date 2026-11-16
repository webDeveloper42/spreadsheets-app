# Sheets

A spreadsheet app that works like Excel and Google Sheets. It runs in your browser as part of a personal dashboard.

It's built from scratch: no code or artwork from Excel or Google Sheets is used.

---

## Features

### Formulas
- **230+ functions**: math, statistics, logic, text, dates, money and lookups (VLOOKUP, INDEX/MATCH, XLOOKUP and others).
- **Spilling formulas**: FILTER, SORT, UNIQUE and SEQUENCE fill the cells below them automatically.
- **References work like Excel**: `$A$1`, other sheets (`Sheet2!A1`), whole columns (`A:A`) and `LET`.
- **Same error messages**: `#N/A` and `#DIV/0!` look the same as in Excel, and the status bar explains each one.
- **A full function list**: **Help → Every function** shows every function and what it does.

### Typing and editing
- **Smart typing**: `$1,200` becomes money, `15%` a percent, `3/5/2024` a date and `2:30 pm` a time.
- **The usual shortcuts**: Enter, Tab, F2, F4, Ctrl+D, Ctrl+Enter, Ctrl+arrow keys and others.
- **Formula help while typing**: suggestions and argument hints appear as you type a formula.
- **Fill handle**: drag the corner of a cell to copy a formula or continue a series (1, 2, 3… or Jan, Feb…).
- **Copy and paste with Excel and Google Sheets**, keeping bold, colors and formulas.
- **Undo and redo** for everything.

### Formatting
- **Number formats**: currency, percent, dates, times and custom formats.
- **Text and cell styles**: bold, italic, underline, colors, borders, alignment and text wrapping.
- **Cell extras**: checkboxes, notes and merged cells.
- **Shrink to fit**: long text shrinks to fit its cell.

### Organizing data
- **Column names**: show names like "Company" in the header instead of A, B, C.
- **Dropdown lists** with colored chips.
- **Filters** on the header row.
- **Cleanup tools**: sort, remove duplicates, trim spaces, split text into columns, and find and replace.
- **Live totals**: the status bar shows the sum, average, min, max and count of the selected cells.

### Charts
- **Chart types**: column, bar, line and pie charts that float over the sheet.

### Rows, columns and sheets
- **Rows and columns**: insert, delete, resize, freeze and hide them.
- **Sheets**: add, rename, copy, color, hide and reorder them. Formulas update to match.

### Saving and files
- **Autosave**: every change is saved automatically.
- **Version history**: you can go back to an earlier copy.
- **Two open tabs**: you get a warning instead of one tab overwriting the other.
- **Import** `.xlsx`, `.csv` and `.tsv` files.
- **Export** to `.xlsx`, `.csv` or `.tsv`.

### Templates
- **Ready-made sheets**: monthly budget, to-do list, loan calculator and workout log.

### Not built yet
- Conditional formatting (beyond dropdown colors)
- Pivot tables
- Comment threads
- Sharing with other people

---

## How it's organized

| File | What it does |
|---|---|
| `static/sheets/formula.js` | The formula engine: all the functions, number formats and smart typing |
| `static/sheets/app.js` | The page itself: the grid, menus, editing and saving |
| `static/sheets/index.html`, `app.css` | The page layout and styles |
| `sheets/store.py` | Saves workbooks and version history in a SQLite database |
| `sheets/files.py` | Reads and writes `.xlsx` and `.csv` files (Python standard library only, no add-ons) |
| `sheets/webapi.py` | Connects the page to the server, and holds the templates |

Formulas are calculated in the browser. Each save also stores the results, so exports don't have to recalculate anything.

---

## Running the tests

```sh
# Storage, files, templates and the formula engine
python3 -m unittest discover -s tests

# The page: start a test server on throwaway data, then drive the page in Node
python3 tests/serve_isolated.py 8891 /tmp/sheets-test &
node tests/ui_test.js http://127.0.0.1:8891
```

The page tests need the dashboard (`capture-app`) installed next to this app.
