"""Regenerate QCR form for FTI Submittal #4 with correct reopen_count"""
import sys
sys.path.insert(0, '.')

from app import generate_qcr_form_html

# Regenerate QCR form for item 36 (FTI Submittal #4)
item_id = 36
result = generate_qcr_form_html(item_id)

if result.get('success'):
    print(f"QCR form regenerated successfully!")
    print(f"Path: {result.get('form_path')}")
else:
    print(f"Error: {result.get('error')}")
