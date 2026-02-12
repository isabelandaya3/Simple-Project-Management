"""Recover the incorrectly rejected QCR response from Friday for FTI Submittal #4"""
import sqlite3
import json
from pathlib import Path
from datetime import datetime

# The response file that was incorrectly rejected
old_response_path = Path(r"\\sac-filsrv1\Projects\Structural-028\Projects\LEB\9.0_Const_Svcs\Fti\Submittals\Submittal - 4 - 260 - MES (Mechanical Equipment Skid)\Responses\_old_iteration__qcr_response.json")

# Read the response
with open(old_response_path, 'r') as f:
    data = json.load(f)

print("=== Response from Friday ===")
print(f"Submitted at: {data.get('_submitted_at')}")
print(f"QCR Name: {data.get('qcr_name')}")
print(f"Action: {data.get('qc_action')}")
print(f"Response Mode: {data.get('response_mode')}")
print(f"Response Category: {data.get('response_category')}")
print(f"Response Text: {data.get('response_text', '')[:100]}...")
print(f"QCR Notes: {data.get('qcr_notes', '')[:100]}...")
print(f"Selected Files: {data.get('selected_files')}")

# Ask for confirmation
confirm = input("\nRecover this response? (y/n): ")
if confirm.lower() != 'y':
    print("Aborted.")
    exit()

# Connect to database
conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

item_id = data.get('item_id')
qc_action = data.get('qc_action')

# Update the item with the response
selected_files_json = json.dumps(data.get('selected_files', []))
final_response_text = data.get('response_text', '')

if qc_action == 'Send Back':
    cursor.execute('''
        UPDATE item SET
            qcr_action = 'Send Back',
            qcr_notes = ?,
            qcr_internal_notes = ?,
            qcr_response_at = ?,
            qcr_response_status = 'Waiting for Revision',
            reviewer_response_status = 'Revision Requested',
            status = 'In Review'
        WHERE id = ?
    ''', (
        data.get('qcr_notes'),
        data.get('qcr_internal_notes'),
        data.get('_submitted_at'),
        item_id
    ))
else:
    # Approve or Modify
    cursor.execute('''
        UPDATE item SET
            qcr_action = ?,
            qcr_notes = ?,
            qcr_internal_notes = ?,
            qcr_response_at = ?,
            qcr_response_status = 'Responded',
            qcr_response_mode = ?,
            qcr_response_text = ?,
            qcr_response_category = ?,
            final_response_category = ?,
            final_response_text = ?,
            final_response_files = ?,
            status = 'Ready for Response'
        WHERE id = ?
    ''', (
        qc_action,
        data.get('qcr_notes'),
        data.get('qcr_internal_notes'),
        data.get('_submitted_at'),
        data.get('response_mode'),
        final_response_text,
        data.get('response_category'),
        data.get('response_category'),
        final_response_text,
        selected_files_json,
        item_id
    ))

conn.commit()

# Verify the update
cursor.execute('SELECT qcr_action, qcr_response_status, qcr_response_at, status FROM item WHERE id = ?', (item_id,))
result = cursor.fetchone()
print(f"\n=== Updated Item ===")
print(f"QCR Action: {result['qcr_action']}")
print(f"QCR Response Status: {result['qcr_response_status']}")
print(f"QCR Response At: {result['qcr_response_at']}")
print(f"Status: {result['status']}")

conn.close()

# Rename the file to indicate it's been processed
processed_path = old_response_path.parent / f"_qcr_response_recovered_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
old_response_path.rename(processed_path)
print(f"\nRenamed file to: {processed_path.name}")
print("\nRecovery complete!")
