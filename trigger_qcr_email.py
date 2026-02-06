"""Script to manually trigger QCR email for an item.

WARNING: This script should only be used when:
1. The QCR email genuinely failed to send
2. The item is in 'In QC' status with all reviewers responded
3. qcr_email_sent_at should be checked BEFORE running this

If qcr_email_sent_at is already set, the email will be skipped (atomic check).
Do NOT manually reset qcr_email_sent_at unless the email was never actually sent.
"""
import sys
import sqlite3
import pythoncom
import win32com.client
from datetime import datetime
import json

# Import from app.py
sys.path.insert(0, '.')
from app import send_multi_reviewer_qcr_email, get_db

item_id = 18

# First check current state
conn = get_db()
cursor = conn.cursor()
cursor.execute('SELECT identifier, status, qcr_email_sent_at FROM item WHERE id = ?', (item_id,))
item = cursor.fetchone()
conn.close()

if item:
    print(f"Item: {item['identifier']}")
    print(f"Status: {item['status']}")
    print(f"QCR Email Sent At: {item['qcr_email_sent_at']}")
    print()
    
    if item['qcr_email_sent_at']:
        print("WARNING: QCR email was already sent!")
        print("The atomic check will skip this to prevent duplicates.")
        print("If you're SURE the email was never received, manually set qcr_email_sent_at to NULL first.")
        print()

print(f"Triggering QCR email for item {item_id}...")
result = send_multi_reviewer_qcr_email(item_id)
print(f"Result: {result}")

# Verify
conn = get_db()
cursor = conn.cursor()
cursor.execute('SELECT qcr_email_sent_at FROM item WHERE id = ?', (item_id,))
item = cursor.fetchone()
print(f"qcr_email_sent_at: {item['qcr_email_sent_at']}")
conn.close()
