"""Script to manually trigger QCR email for an item."""
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
