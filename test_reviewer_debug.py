import sqlite3
from pathlib import Path

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

item_id = 53

cursor.execute('''
    SELECT i.*, 
           ir.display_name as reviewer_name, ir.email as reviewer_email,
           qcr.display_name as qcr_name, qcr.email as qcr_email
    FROM item i
    LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
    LEFT JOIN user qcr ON i.qcr_id = qcr.id
    WHERE i.id = ?
''', (item_id,))
item = cursor.fetchone()

# Convert to dict for easier handling
item = dict(item)

print(f"reviewer_name value: {repr(item.get('reviewer_name'))}")
print(f"bool test: {bool(item.get('reviewer_name'))}")
print(f"is None: {item.get('reviewer_name') is None}")
print(f"== None: {item.get('reviewer_name') == None}")
print(f"not test: {not item.get('reviewer_name')}")

conn.close()
