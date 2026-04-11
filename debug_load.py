#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import traceback

try:
    from pulse_v2.data.cache import init_full_load, CACHE
    print("Loading data...", file=sys.stderr, flush=True)
    init_full_load()
    print(f"✓ SUCCESS: {len(CACHE)} cache entries", file=sys.stderr, flush=True)
except SystemExit as e:
    print(f"SystemExit: {e}", file=sys.stderr)
    sys.exit(0)
except Exception as e:
    print(f"✗ ERROR: {type(e).__name__}", file=sys.stderr)
    print(f"  Message: {str(e)}", file=sys.stderr)
    traceback.print_exc(file=sys.stderr)
    sys.exit(1)
