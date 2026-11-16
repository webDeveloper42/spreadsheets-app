#!/bin/sh
# Copies the page files from the dashboard into this folder, ready to commit.
set -e
cd "$(dirname "$0")"
mkdir -p static/sheets data/snapshot
cp ~/capture-app/capture_app/static/sheets/index.html ~/capture-app/capture_app/static/sheets/app.css \
   ~/capture-app/capture_app/static/sheets/app.js ~/capture-app/capture_app/static/sheets/formula.js static/sheets/
python3 - <<'PY'
import sqlite3
src = sqlite3.connect("file:data/sheets.db?mode=ro", uri=True)
dst = sqlite3.connect("data/snapshot/sheets.db")
src.backup(dst)
dst.execute("PRAGMA journal_mode=DELETE"); dst.close(); src.close()
PY
echo "synced"
