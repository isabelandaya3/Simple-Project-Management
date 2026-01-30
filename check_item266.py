#!/usr/bin/env python3
"""Check item 266 in database."""
import sqlite3

conn = sqlite3.connect('c:/Users/IANDAYA/Documents/Project Management -Simple/tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Search for submittal 266
print("=== Searching for Submittal 266 ===")
cursor.execute("SELECT id, identifier, title, status, reviewer_email_sent_at, reviewer_response_at, qcr_email_sent_at, qcr_response_at, initial_reviewer_id, qcr_id FROM item WHERE identifier LIKE '%266%'")
rows = cursor.fetchall()
for row in rows:
    print(dict(row))

if rows:
    item_id = rows[0]['id']
    print(f"\n=== Full details for item {item_id} ===")
    cursor.execute("SELECT * FROM item WHERE id = ?", (item_id,))
    item = cursor.fetchone()
    if item:
        for key in item.keys():
            print(f"{key}: {item[key]}")
    
    print(f"\n=== Reviewers for item {item_id} ===")
    cursor.execute("SELECT * FROM item_reviewers WHERE item_id = ?", (item_id,))
    reviewers = cursor.fetchall()
    for r in reviewers:
        print(dict(r))

conn.close()
