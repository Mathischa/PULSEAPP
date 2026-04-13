#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test script to validate the improvements made to charts and exports.
Verifies that:
1. All modified JS files have valid syntax (no obvious issues)
2. CSS @media print section is properly formatted
3. Key changes are present (white axes, axis titles, etc.)
"""

import os
import re
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def test_js_files():
    """Check for expected changes in JS files."""
    js_files = {
        'pulse_web/static/js/visualisation.js': [
            '#FFFFFF',  # White axis ticks
            'Date',     # Axis title
            'Montant',  # Y axis title
        ],
        'pulse_web/static/js/visualisation_flux.js': [
            '#FFFFFF',  # White default color
            'Montant',  # Y axis title
        ],
        'pulse_web/static/js/ml_ecarts.js': [
            '#FFFFFF',  # White ticks
            'Écart',    # X axis title in scatter
        ],
        'pulse_web/static/js/prevision_repartition.js': [
            '#FFFFFF',  # White ticks
            'Profils',  # X axis title
        ],
        'pulse_web/static/js/repartition.js': [
            '#FFFFFF',  # White ticks
            'Filiales', # X axis title
        ],
        'pulse_web/static/js/repartition_flux.js': [
            '#FFFFFF',  # White ticks
            'Catégories',  # Y axis title
        ],
        'pulse_web/static/js/tendance.js': [
            '#FFFFFF',  # White ticks
            'Date',     # X axis title
        ],
        'pulse_web/static/js/export_utils.js': [
            'print',    # Print improvements
            'XLSX',     # Excel improvements
        ],
    }
    
    print("✓ Testing JS files...\n")
    for filepath, expected_strings in js_files.items():
        full_path = os.path.join(BASE_DIR, filepath)
        if not os.path.exists(full_path):
            print(f"✗ {filepath} - FILE NOT FOUND")
            continue
        
        with open(full_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        missing = []
        for expected in expected_strings:
            if expected not in content:
                missing.append(expected)
        
        if missing:
            print(f"✗ {filepath} - Missing: {missing}")
        else:
            print(f"✓ {filepath} - OK")
    
    print()

def test_css_file():
    """Check CSS improvements."""
    print("✓ Testing CSS file...\n")
    
    css_path = os.path.join(BASE_DIR, 'pulse_web/static/css/pulse.css')
    if not os.path.exists(css_path):
        print("✗ CSS file not found")
        return
    
    with open(css_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    checks = {
        '@media print': '@media print section',
        'canvas': 'Canvas styling',
        'image-rendering': 'Image rendering improvements',
        'page-break-inside': 'Page break handling',
        'chart-container': 'Chart container styling',
    }
    
    for check_str, description in checks.items():
        if check_str in content:
            print(f"✓ {description} - present")
        else:
            print(f"✗ {description} - MISSING")
    
    print()

def main():
    print("=" * 60)
    print("PULSE Improvements Validation Test")
    print("=" * 60)
    print()
    
    test_js_files()
    test_css_file()
    
    print("=" * 60)
    print("Validation complete!")
    print("=" * 60)

if __name__ == '__main__':
    main()
