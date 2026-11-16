"""Runs the dashboard on throwaway data, for driving the Sheets page in tests.  python3 serve_isolated.py PORT FOLDER

Every app the dashboard opens points into FOLDER, so nothing real is read or written.
"""
import os
import sys

port, base = int(sys.argv[1]), os.path.abspath(sys.argv[2])
for d in ("todo", "motion", "xdg", "cal", "hub", "notes", "disc", "incoming", "jobs", "media-files", "vc", "received", "home"):
    os.makedirs(os.path.join(base, d), exist_ok=True)
os.environ.update({"CAPTURE_DATA_DIR": base, "TICKDO_HOME": base + "/todo", "MOTION_DATA_DIR": base + "/motion", "XDG_DATA_HOME": base + "/xdg",
                   "CALENDAR_DATA_DIR": base + "/cal", "HUB_DIR": base + "/hub", "NOTES_DB": base + "/notes/n.db", "DISCOVER_DB": base + "/disc/d.db",
                   "CAPTURE_INCOMING": base + "/incoming", "DISCOVER_NO_EMAIL": "1", "JOBRIGHT_DATA_DIR": base + "/jobs", "LIMINALS_DB": base + "/liminals.db",
                   "OLLAMA_HOST": "http://127.0.0.1:9", "MEDIA_GRAB_DB": base + "/media.db", "MEDIA_GRAB_DIR": base + "/media-files", "MEDIA_GRAB_YTDLP": "/nonexistent",
                   "MEDIA_GRAB_HOME": base + "/home", "FTP_BRIDGE_DB": base + "/ftp.db", "FTP_BRIDGE_RECEIVED": base + "/received", "FTP_BRIDGE_HOME": base + "/home",
                   "SHEETS_DB": base + "/sheets.db", "DOCS_DB": base + "/docs.db", "DOCS_IMAGES": base + "/doc-images", "DRIVE_ROOT": base + "/Drive", "DRIVE_DATA": base + "/drive-data", "WRITEFIX_PERSONAL": base + "/writefix-personal.txt", "LIMINALS_VOICE_CACHE": base + "/vc"})
sys.path.insert(0, os.path.expanduser("~/capture-app"))
from http.server import ThreadingHTTPServer  # noqa: E402
from capture_app import server  # noqa: E402

print("serving", flush=True)
ThreadingHTTPServer(("127.0.0.1", port), server.Handler).serve_forever()
