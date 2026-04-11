#!/usr/bin/env python
# -*- coding: utf-8 -*-
from pulse_v2.data.cache import CACHE, TOKENS, sections

print("=" * 70)
print("CACHE CONTENT CHECK")
print("=" * 70)

print(f"\nSections ({len(sections)}): {list(sections.keys())}")
print(f"Total cache entries: {len(CACHE)}")

print("\nFirst 5 cache keys and their data sizes:")
for key in list(CACHE.keys())[:5]:
    v = CACHE[key]
    dates_count = len(v.get('dates', []))
    reel_count = len(v.get('reel', []))
    prev_count = len(v.get('prev_vals', []))
    print(f"  {key}: dates={dates_count}, reel={reel_count}, prev_series={prev_count}")

print("\nTokens (flux by section):")
for section, flux_list in TOKENS.items():
    print(f"  {section}: {len(flux_list)} flux")
    if flux_list:
        print(f"    - {flux_list[0][0]} (col {flux_list[0][1]})")

print("\n" + "=" * 70)
