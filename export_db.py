#!/usr/bin/env python3
"""
Norton Abrasives — Database Export Script
=========================================
Run this script to convert your Excel file to bonded.json

Usage:
    python export_db.py                          # uses Bonded_Equivalent.xlsx by default
    python export_db.py my_file.xlsx             # use a specific Excel file
    python export_db.py my_file.xlsx Sheet2      # use a specific sheet name

Output:
    data/bonded.json   (created automatically)

Required columns in Excel (row 1 = headers):
    Competition  | Grain | Grit | Grade | Bond  | Speed | 
    Grain (Norton) | Grit (Norton) | Grade (Norton) | Bond (Norton) | 
    Application | Norton Equivalent Spec

Install dependencies if needed:
    pip install pandas openpyxl
"""

import sys
import json
import os
import pandas as pd
from datetime import datetime

# ── CONFIG ──────────────────────────────────────────────
DEFAULT_FILE  = 'Bonded_Equivlent_Test.xlsx'   # change to your filename
DEFAULT_SHEET = 0                               # 0 = first sheet, or use sheet name

# Column mapping — adjust if your Excel headers differ
COL_MAP = {
    'competitor':   'Competition',
    'comp_grain':   'Grain',
    'comp_grit':    'Grit',
    'comp_grade':   'Grade',
    'comp_bond':    'Bond ',       # note trailing space in source file
    'speed':        'Speed',
    'norton_grain': 'Grain ',      # note trailing space in source file
    'norton_grit':  'Grit2',
    'norton_grade': 'Grade3',
    'norton_bond':  'Bond',
    'application':  'Application',
    'norton_spec':  'Norton Equivalent Spec',
}
# ────────────────────────────────────────────────────────


def export(xlsx_path, sheet=DEFAULT_SHEET):
    print(f"\n📂  Reading: {xlsx_path}")
    if not os.path.exists(xlsx_path):
        print(f"❌  File not found: {xlsx_path}")
        sys.exit(1)

    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=0)
    print(f"    Rows loaded: {len(df)}")
    print(f"    Columns:     {df.columns.tolist()}")

    # Auto-detect column mapping if headers differ
    actual_map = {}
    for key, expected_col in COL_MAP.items():
        # Try exact match first
        if expected_col in df.columns:
            actual_map[key] = expected_col
        else:
            # Try stripped match
            stripped = {c.strip(): c for c in df.columns}
            if expected_col.strip() in stripped:
                actual_map[key] = stripped[expected_col.strip()]
            else:
                print(f"⚠️   Could not find column '{expected_col}' for field '{key}'")

    missing = [k for k in COL_MAP if k not in actual_map]
    if missing:
        print(f"\n❌  Missing columns: {missing}")
        print(    "    Check COL_MAP at the top of this script and adjust headers.")
        sys.exit(1)

    # Strip string columns
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()

    # Drop rows with no competitor
    df = df[df[actual_map['competitor']].notna()]
    df = df[df[actual_map['competitor']] != 'nan']

    records = []
    errors  = 0
    for idx, row in df.iterrows():
        try:
            rec = {
                'competitor':   row[actual_map['competitor']],
                'comp_grain':   row[actual_map['comp_grain']],
                'comp_grit':    int(float(row[actual_map['comp_grit']])),
                'comp_grade':   row[actual_map['comp_grade']],
                'comp_bond':    row[actual_map['comp_bond']],
                'application':  row[actual_map['application']],
                'norton_spec':  row[actual_map['norton_spec']],
                'norton_grain': row[actual_map['norton_grain']],
                'norton_grit':  int(float(row[actual_map['norton_grit']])),
                'norton_grade': row[actual_map['norton_grade']],
                'norton_bond':  row[actual_map['norton_bond']],
                'speed':        row[actual_map['speed']],
            }
            records.append(rec)
        except Exception as e:
            errors += 1
            if errors <= 5:
                print(f"⚠️   Row {idx+2} skipped: {e}")

    print(f"\n✅  Records exported:  {len(records)}")
    if errors:
        print(f"⚠️   Rows skipped:      {errors}")

    # Summary
    competitors = sorted(set(r['competitor'] for r in records))
    apps        = sorted(set(r['application'] for r in records))
    print(f"    Competitors:       {competitors}")
    print(f"    Applications:      {apps}")

    # Write output
    os.makedirs('data', exist_ok=True)
    out_path = os.path.join('data', 'bonded.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(records, f, indent=2, ensure_ascii=False)

    size_kb = os.path.getsize(out_path) // 1024
    print(f"\n📄  Written to:  {out_path}  ({size_kb} KB)")
    print(f"    Timestamp:   {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("\n🚀  Done! Upload data/bonded.json to GitHub to update the live app.\n")


if __name__ == '__main__':
    xlsx = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_FILE
    sht  = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_SHEET
    export(xlsx, sht)
