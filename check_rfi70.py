import sqlite3
import json
from pathlib import Path

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# First, list all tables
cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
tables = cursor.fetchall()
print("=== Tables in database ===")
for t in tables:
    print(t['name'])
print()

# Check schema of item table
cursor.execute("PRAGMA table_info(item)")
cols = cursor.fetchall()
print("=== item table schema ===")
for c in cols:
    print(f"  {c['name']} ({c['type']})")
print()

# List all RFIs to see identifier format
print("=== Sample RFIs to see identifier format ===")
cursor.execute("SELECT id, type, identifier, title, status FROM item WHERE type = 'RFI' ORDER BY id")
rfis = cursor.fetchall()
for r in rfis:
    print(f"  ID={r['id']}, identifier='{r['identifier']}', status={r['status']}, title={r['title'][:50] if r['title'] else 'N/A'}")
print()

# Find RFI #70 - using identifier column, try LIKE search
cursor.execute("SELECT id, type, identifier, title, status, reviewer_response_status, qcr_response_status, reviewer_email_sent_at, qcr_email_sent_at, folder_link FROM item WHERE identifier LIKE '%70%' OR title LIKE '%70%'")
rows = cursor.fetchall()

print("=== RFI #70 from items table ===")
if rows:
    for row in rows:
        print(json.dumps(dict(row), indent=2, default=str))
else:
    print("No RFI #70 found")

# Check reviewers for this item
if rows:
    item_id = rows[0]['id']
    print(f"\n=== Reviewers for RFI #70 (item_id={item_id}) ===")
    cursor.execute("SELECT * FROM item_reviewers WHERE item_id = ?", (item_id,))
    reviewers = cursor.fetchall()
    for r in reviewers:
        print(json.dumps(dict(r), indent=2, default=str))

# Check reminder_log
if rows:
    print(f"\n=== Reminder log for RFI #70 ===")
    cursor.execute("SELECT * FROM reminder_log WHERE item_id = ?", (item_id,))
    reminders = cursor.fetchall()
    for r in reminders:
        print(json.dumps(dict(r), indent=2, default=str))

conn.close()
