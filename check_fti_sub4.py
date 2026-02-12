"""Check FTI Submittal #4 - Compare Friday response with current DB"""
import sqlite3
import json
from pathlib import Path

# Friday's response
old_path = Path(r"\\sac-filsrv1\Projects\Structural-028\Projects\LEB\9.0_Const_Svcs\Fti\Submittals\Submittal - 4 - 260 - MES (Mechanical Equipment Skid)\Responses\_old_iteration__qcr_response.json")
if old_path.exists():
    with open(old_path, 'r') as f:
        friday = json.load(f)
    print("=== FRIDAY (Feb 6) Response - REJECTED ===")
    print(f"  Submitted: {friday.get('_submitted_at')}")
    print(f"  Action: {friday.get('qc_action')}")
    print(f"  Category: {friday.get('response_category')}")
    print(f"  Response Mode: {friday.get('response_mode')}")
    print(f"  Status would be: Ready for Response")
else:
    print("Friday response file not found")

# Current DB state
conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()
c.execute('SELECT qcr_action, qcr_response_at, qcr_response_status, qcr_notes, status, qcr_response_category FROM item WHERE id=36')
db = c.fetchone()

print()
print("=== CURRENT DB State ===")
print(f"  QCR Action: {db['qcr_action']}")
print(f"  QCR Response At: {db['qcr_response_at']}")
print(f"  QCR Response Status: {db['qcr_response_status']}")
print(f"  QCR Response Category: {db['qcr_response_category']}")
print(f"  Item Status: {db['status']}")

conn.close()
