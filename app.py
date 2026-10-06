"""Replay Scout - enter Showdown username(s), get a scouting workbook.
Run:  pip install -r requirements.txt && python app.py   ->  http://localhost:5000
"""
import io, itertools, json, os, re, threading, time, uuid
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor

import requests
from flask import Flask, jsonify, render_template_string, request, send_file
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = "https://replay.pokemonshowdown.com"
HDR = {"User-Agent": "replay-scout/1.0 (personal scouting tool)"}
HERE = os.path.dirname(os.path.abspath(__file__))
CACHE, OUT = os.path.join(HERE, "cache"), os.path.join(HERE, "output")
os.makedirs(CACHE, exist_ok=True); os.makedirs(OUT, exist_ok=True)
norm = lambda s: re.sub(r"[^a-z0-9]", "", (s or "").lower())

# ---------------------------------------------------------------- fetching
def search(name, fmt, max_pages=40):
    rows = []
    for page in range(1, max_pages + 1):
        p = {"user": name, "page": page}
        if fmt: p["format"] = fmt
        r = requests.get(f"{BASE}/search.json", params=p, headers=HDR, timeout=20)
        r.raise_for_status()
        batch = r.json()
        rows += batch
        if len(batch) < 50: break
        time.sleep(0.3)
    return rows
