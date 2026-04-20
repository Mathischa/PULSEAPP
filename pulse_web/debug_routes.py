#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Debug script to list all Flask routes."""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app

print("=" * 80)
print("All registered Flask routes:")
print("=" * 80)

for rule in app.url_map.iter_rules():
    print(f"{rule.rule:50} {str(rule.methods):30} {rule.endpoint}")

print("=" * 80)
print(f"\nLooking for browse_folder route...")
browse_routes = [r for r in app.url_map.iter_rules() if 'browse_folder' in r.rule]
if browse_routes:
    print("✓ Found browse_folder routes:")
    for r in browse_routes:
        print(f"  {r.rule} -> {r.endpoint}")
else:
    print("✗ No browse_folder route found!")

print("\nLooking for any /api/import_profils routes...")
api_routes = [r for r in app.url_map.iter_rules() if '/api/import_profils' in r.rule]
if api_routes:
    print("✓ Found api/import_profils routes:")
    for r in api_routes:
        print(f"  {r.rule} -> {r.endpoint}")
else:
    print("✗ No api/import_profils routes found!")
