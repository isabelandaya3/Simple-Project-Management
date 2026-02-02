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

print(f"Before fallback: reviewer_name = '{item.get('reviewer_name')}'")

# Check for reviewer names in item_reviewers table if reviewer_name is not set
if not item.get('reviewer_name'):
    print("reviewer_name is empty, checking item_reviewers...")
    cursor.execute('''
        SELECT GROUP_CONCAT(reviewer_name, ', ') as reviewer_names
        FROM item_reviewers
        WHERE item_id = ?
    ''', (item_id,))
    result = cursor.fetchone()
    print(f"Query result: {result}")
    print(f"reviewer_names: {result['reviewer_names'] if result else 'None'}")
    if result and result['reviewer_names']:
        item['reviewer_name'] = result['reviewer_names']

print(f"After fallback: reviewer_name = '{item.get('reviewer_name')}'")

conn.close()
