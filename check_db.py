#!/usr/bin/env python3
"""Check reviewer info in database."""
import sqlite3

conn = sqlite3.connect('c:/Users/IANDAYA/Documents/Project Management -Simple/tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

print("=== Reviewers for Item 17 ===")
cursor.execute('SELECT * FROM item_reviewers WHERE item_id = 17')
rows = cursor.fetchall()
for row in rows:
    print(dict(row))

print("\n=== Item 17 Details ===")
cursor.execute('SELECT * FROM item WHERE id = 17')
item = cursor.fetchone()
if item:
    print(dict(item))

conn.close()
