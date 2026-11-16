# Sheets

The dashboard's **Sheets** tab (Plan group, `/sheets/` on port 8800): spreadsheets built from how Excel and Google Sheets
work. None of their code or artwork is used.

## What it does

* **Formulas**: 230+ functions (`Help → Every function` lists them with what each does): maths, statistics, logic,
  lookup (VLOOKUP, HLOOKUP, INDEX/MATCH, XLOOKUP, XMATCH, INDIRECT, OFFSET), text (TEXT, TEXTJOIN, SPLIT, REGEX…),
  dates (DATEDIF, NETWORKDAYS, EOMONTH…), finance (PMT, IPMT, FV, NPV, IRR, RATE…), information, and the dynamic-array
  ones that **spill** into the cells below (FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, TRANSPOSE, VSTACK/HSTACK, TAKE/DROP…).
  Any function works over a range (=LEN(A1:A5) gives five answers), so ARRAYFORMULA isn't needed.
  `$A$1` absolute references, other sheets (`Sheet2!A1`, `'My sheet'!A1`), whole columns/rows (`A:A`, `1:1`), `LET`,
  `{1,2;3,4}` arrays, the same error values (#N/A, #DIV/0!, #REF!, #SPILL!…, with the reason in the status bar),
  circular references caught, and `=0.1+0.2=0.3` is TRUE like both apps.
* **Typing**: `$1,200` is money, `15%` a percent, `3/5/2024` or `2024-03-05` a date, `2:30 pm` a time, `'007` text.
* **Editing like the apps**: Enter/Tab move, F2 edits, typing a formula suggests functions and shows the arguments,
  clicking or arrowing while typing a formula puts that cell in, F4 cycles `$`, unclosed brackets close themselves,
  Ctrl+Enter fills a selection, Alt+Enter is a new line, Ctrl+arrow jumps to the edge of the data, Ctrl+Space and
  Shift+Space select columns and rows.
* **Fill handle**: drag the corner square to copy formulas (references move) or continue a series (1, 2, 3…, dates,
  Jan/Feb…, Monday…, Item 1, Item 2…); double-click it to fill down as far as the next column goes. Ctrl+D / Ctrl+R.
* **Copy and paste** with Excel and Google Sheets (tab-separated text, plus bold/italic/colours from their tables),
  references moving on paste, cut to move, and Ctrl+Shift+V for values only. Pasting into a selection that's a multiple
  of what was copied repeats it.
* **Formatting**: number formats (currency, percent, dates, times, durations, financial, scientific, custom codes with
  up to four sections and [Red]), decimals +/-, bold, italic, underline, strikethrough, font size, text and fill colours,
  borders, alignment, wrap (rows grow to fit), checkboxes, notes.
* **Rows, columns and sheets**: insert and delete (formulas everywhere follow), resize by dragging, double-click to fit,
  freeze rows and columns, zoom, more rows appear as you scroll; sheets can be added, renamed (formulas follow),
  duplicated, coloured, hidden, moved and deleted.
* **Data**: sort (keeping a header row), remove duplicate rows, trim spaces, split text into columns, find and replace
  (case, whole cell, inside formulas, every sheet). The status bar shows the sum, average, min, max and count.
* **Undo/redo** for everything, **autosave** a moment after each change, **version history** (one copy at most every
  10 minutes, and before every restore; the newest 60 kept), and a warning instead of overwriting when two tabs edit
  the same workbook.
* **Files**: import `.xlsx` (values, formulas including shared ones, styles, number formats, widths, heights, frozen
  panes, hidden sheets) and `.csv`/`.tsv` (numbers, money, percents and dates recognised; formula-looking text is kept
  as text). Download as `.xlsx` (formulas with their results, styles, widths, frozen panes) or a sheet as `.csv`/`.tsv`
  with values shown the way the sheet shows them. `.xls`, `.ods` and `.numbers` need saving as `.xlsx` first.
* **Templates**: monthly budget, to-do list, loan calculator (with a month-by-month table), workout log.

Not built (yet): charts, merged cells, conditional formatting, data validation dropdowns, pivot tables, filter views,
comments with replies, and sharing with other people (the server is yours alone). Excel's spilled formulas reopen as
formulas Excel calculates itself.

## How it fits together

```
~/capture-app/capture_app/static/sheets/formula.js   the formula engine, number formats and typed input (no page code;
                                                      Node runs it for the tests)
~/capture-app/capture_app/static/sheets/app.js       the page: canvas grid, editing, menus, files, saving
sheets/store.py     SQLite: books (each workbook as JSON with a revision) and versions. data/sheets.db, chmod 600
sheets/files.py     .xlsx and .csv in and out, standard library only
sheets/webapi.py    what the tab can do, plus the templates
```

Formulas run in the browser. Each save carries every formula's latest result (`c`) and each formatted number's shown
text (`t`), so exports and the data hub see values without running the engine on the server. Routes:
`/api/sheets/{state,book,do,import,export}`. The data hub knows it as `sheets` (optional).

## This repository

`static/sheets/` is a copy of the page files from `~/capture-app/capture_app/static/sheets/`; `./sync_for_git.sh`
refreshes it before a commit. Workbooks are not in the repository (it is public); the script also writes a local copy
to `data/snapshot/sheets.db`, which git ignores.

## Tests

```
python3 -m unittest discover -s tests          # storage, files, templates, and the formula engine (tests/formula_test.js, 578 checks, via node)
python3 tests/serve_isolated.py 8891 /tmp/x &  # the dashboard on throwaway data
node tests/ui_test.js http://127.0.0.1:8891     # the page itself, driven in Node with a stand-in browser
```

`backups/before-sheets-tab-2026-11-10/` has the dashboard's server.py, tools.py and tests from before this tab.
