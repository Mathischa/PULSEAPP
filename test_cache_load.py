#!/usr/bin/env python
# -*- coding: utf-8 -*-
from pulse_v2.data.cache import init_full_load, CACHE, TOKENS, sections

print("=" * 70)
print("AVANT init_full_load:")
print(f"  CACHE: {len(CACHE)} items")
print(f"  TOKENS: {sum(len(v) for v in TOKENS.values())} flux")
print(f" Sections: {list(sections.keys())}")

try:
    init_full_load()
    print("\n" + "=" * 70)
    print("APRES init_full_load (SUCCÈS):")
    print(f"  CACHE: {len(CACHE)} items")
    print(f"  TOKENS total: {sum(len(v) for v in TOKENS.values())} flux")
    print(f"  Sections: {list(sections.keys())}")
    if CACHE:
        print(f"\n  Premier item du cache:")
        for key in list(CACHE.keys())[:1]:
            v = CACHE[key]
            print(f"    {key}: dates={len(v.get('dates', []))}, reel={len(v.get('reel', []))}")
except Exception as e:
    print("\n" + "=" * 70)
    print(f"ERREUR during init_full_load: {e}")
    import traceback
    traceback.print_exc()
