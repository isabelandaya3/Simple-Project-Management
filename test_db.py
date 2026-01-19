import sqlite3
from pathlib import Path

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

# Check item status
c.execute('SELECT id, identifier, status, qcr_id FROM item WHERE id = 17')
item = c.fetchone()
print(f"Item 17: identifier={item['identifier']}, status={item['status']}, qcr_id={item['qcr_id']}")

# Check reviewers
c.execute('SELECT ir.reviewer_name, ir.response_at, ir.response_category FROM item_reviewers ir WHERE item_id = 17')
revs = c.fetchall()
print('\nReviewers:')
for r in revs:
    print(f"  {r['reviewer_name']}: responded={r['response_at'] is not None}, category={r['response_category']}")

conn.close()
