#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Minimal test of _accumuler_valeurs_tous_mois to find the crash"""
import sys
import pandas as pd
from pulse_v2.data.cache import _lister_fichiers_mensuels, charger_noms_feuilles_depuis_cells

# Get file list
files = _lister_fichiers_mensuels()
print(f"Found {len(files)} files", file=sys.stderr, flush=True)

if not files:
    print("No files found!", file=sys.stderr)
    sys.exit(1)

# Try to read just the first file
first_file = files[0]
path = first_file[2]
print(f"Testing first file: {path.split(chr(92))[-1]}", file=sys.stderr, flush=True)

sheets = list(set(charger_noms_feuilles_depuis_cells().values()))
print(f"Sheets: {sheets}", file=sys.stderr, flush=True)

try:
    data = pd.read_excel(path, sheet_name=sheets, header=None)
    print(f"✓ Read succeeded, type: {type(data)}", file=sys.stderr, flush=True)
    
    if isinstance(data, pd.DataFrame):
        print(f"  Single sheet: {data.shape}", file=sys.stderr, flush=True)
    else:
        for sheet_name, df in data.items():
            print(f"  Sheet {sheet_name}: {df.shape}", file=sys.stderr, flush=True)
            
except Exception as e:
    print(f"✗ Read failed: {e}", file=sys.stderr, flush=True)
    import traceback
    traceback.print_exc(file=sys.stderr)
    sys.exit(1)

print("✓ File reading works", file=sys.stderr, flush=True)
