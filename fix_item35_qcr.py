import sqlite3
import json
from datetime import datetime
from app import send_qcr_assignment_email

# Connect to database
conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Get item 35 details
cursor.execute('SELECT * FROM item WHERE id = 35')
item = cursor.fetchone()

if item:
    print("=== Item 35 Current Status ===")
    print(f"ID: {item['id']}")
    print(f"Identifier: {item['identifier']}")
    print(f"Status: {item['status']}")
    
    # Check if there are reviewer columns
    columns = [desc[0] for desc in cursor.description]
    if 'reviewer_response_at' in columns:
        print(f"Reviewer Response At: {item['reviewer_response_at']}")
    if 'email_token_reviewer' in columns:
        print(f"Expected Token: {item['email_token_reviewer']}")
    
    print("\n=== Manually Processing Response ===")
    
    # Read the response file data
    response_data = {
        "response_category": "Approved as Noted",
        "selected_files": ["13 34 19 ShpDwg_Reactions-AnchorBoldDwg_Rev01 _ReviewedHDR - Revised.pdf"],
        "notes": "See attached engineer comments.",
        "internal_notes": ""
    }
    
    # Update the database directly
    selected_files_json = json.dumps(response_data["selected_files"])
    submitted_at = "2026-01-23T18:03:46.679Z"
    
    cursor.execute('''
        UPDATE item SET
            reviewer_response_category = ?,
            reviewer_notes = ?,
            reviewer_internal_notes = ?,
            reviewer_selected_files = ?,
            reviewer_response_at = ?,
            reviewer_response_status = 'Responded',
            reviewer_response_version = 0,
            status = 'In QC'
        WHERE id = ?
    ''', (
        response_data["response_category"],
        response_data["notes"],
        response_data["internal_notes"],
        selected_files_json,
        submitted_at,
        35
    ))
    
    conn.commit()
    print("✓ Database updated with reviewer response")
    
    # Now send the QCR email
    print("\n=== Sending QCR Email ===")
    try:
        result = send_qcr_assignment_email(35, is_revision=False, version=0)
        if result.get('success'):
            print(f"✓ QCR email sent successfully!")
            print(f"  Message: {result.get('message', 'Email sent')}")
        else:
            print(f"✗ Failed to send QCR email")
            print(f"  Error: {result.get('error', 'Unknown error')}")
    except Exception as e:
        print(f"✗ Exception occurred: {e}")
        import traceback
        traceback.print_exc()
else:
    print("Item 35 not found")

conn.close()
print("\n=== Done ===")
