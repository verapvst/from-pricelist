#!/usr/bin/env python3
"""
sync_data.py — Sync Pice List Tracker.xlsx → ADMIN_UNITS in index.html

Run after updating the Excel:
    python3 sync_data.py

Then commit both files via GitHub Desktop.
"""

import json, math, re
import pandas as pd
from pathlib import Path

REPO  = Path(__file__).parent
EXCEL = REPO / "Pice List Tracker.xlsx"
HTML  = REPO / "index.html"

STATUS_MAP = {
    'AVAILABLE':     'Available',
    'RESERVED':      'Reserved',
    'PRE-RESERVED':  'Reserved',
    'PSPA':          'PSPA',
    'DEED':          'Deeded',
    'DEEDED':        'Deeded',
    'UPON REQUEST':  'Upon Request',
    'UPON  REQUEST': 'Upon Request',
    'NOT FOR SALE':  'Not for Sale',
}

def _nan(v):
    return v is None or (isinstance(v, float) and math.isnan(v))

def s(v):
    """Clean string — strip whitespace, return None if empty."""
    if _nan(v): return None
    r = str(v).strip()
    return r if r else None

def n(v):
    """Clean number — return int if whole, float if not, None if missing."""
    if _nan(v): return None
    try:
        f = float(v)
        return int(f) if f == int(f) else round(f, 2)
    except: return None

def d(v):
    """Clean date → 'YYYY-MM-DD' string or None."""
    if _nan(v): return None
    try:
        return pd.Timestamp(v).strftime('%Y-%m-%d')
    except: return None

def ext(row):
    """Sum all exterior sub-areas (NaN treated as 0)."""
    cols = ['Balconies [sqm]', 'Garden [sqm]', 'Patio [sqm]', 'Terrace [sqm]', 'Rooftop [sqm]']
    total = sum(float(row[c]) for c in cols if not _nan(row.get(c)))
    return round(total, 2) if total else 0

# ── Read Excel ──
print(f"Reading {EXCEL.name}...")
df = pd.read_excel(EXCEL, sheet_name='DATABASE')
print(f"  {len(df)} rows found.")

# ── Build units array ──
units = []
skipped = 0
for _, row in df.iterrows():
    status_raw = s(row.get('Selling Status')) or ''
    status = STATUS_MAP.get(status_raw.upper(), status_raw or 'Available')

    disc_raw = n(row.get('Total Discunt'))
    disc = round(disc_raw * 100, 2) if disc_raw else 0

    unit = {
        'p':    s(row.get('Development')),
        'r':    s(row.get('Local')),
        'u':    s(row.get('Unit PH')),
        't':    s(row.get('Typology')),
        'a':    n(row.get('Area Total  [sqm]')),
        'ai':   n(row.get('Interior')),
        'ae':   ext(row),
        'pool': n(row.get('Pool [#]')) or 0,
        's':    status,
        'pi':   n(row.get('Inicial Asking Price (€) (Total)')),
        'pc':   n(row.get('Current Asking price( Total)')),
        'ps':   n(row.get('Sell price (Total)')),
        'disc': disc,
        'ld':   None,
        'rd':   d(row.get('PSPA Date')),
        'cd':   d(row.get('Delivery Expected')),
        'co':   s(row.get('Owner Comercial')),
    }

    if not unit['p']:
        skipped += 1
        continue

    units.append(unit)

print(f"  {len(units)} units built, {skipped} skipped (no project name).")

# ── Replace ADMIN_UNITS in index.html ──
html = HTML.read_text(encoding='utf-8')

# Find start of const ADMIN_UNITS=
start = html.find('const ADMIN_UNITS=')
if start == -1:
    print("ERROR: 'const ADMIN_UNITS=' not found in index.html")
    exit(1)

# Find the semicolon that ends it
end = html.find(';', start)
if end == -1:
    print("ERROR: Could not find end of ADMIN_UNITS")
    exit(1)

json_str = json.dumps(units, ensure_ascii=False, separators=(',', ':'))
html = html[:start] + f'const ADMIN_UNITS={json_str}' + html[end:]

HTML.write_text(html, encoding='utf-8')
print(f"Done. index.html updated with {len(units)} units.")
print("Next: commit both files via GitHub Desktop.")
