import os, sys, time, threading, webbrowser
from pathlib import Path

BASE = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
os.chdir(BASE)

def open_browser():
    time.sleep(1.8)
    webbrowser.open("http://127.0.0.1:5000")

threading.Thread(target=open_browser, daemon=True).start()

from app import app
app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)
