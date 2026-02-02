import sqlite3
conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Check a few items to see their reviewer setup
cursor.execute('''
    SELECT i.id, i.identifier, i.initial_reviewer_id, i.multi_reviewer_mode,
           u.display_name as reviewer_from_user
    FROM item i
    LEFT JOIN user u ON i.initial_reviewer_id = u.id
    WHERE i.closed_at IS NULL
    LIMIT 10
''')
print('Items with reviewer info:')
for row in cursor.fetchall():
    print(f"  ID {row['id']}: {row['identifier']} - initial_reviewer_id={row['initial_reviewer_id']}, multi_mode={row['multi_reviewer_mode']}, user_name={row['reviewer_from_user']}")

print()
print('Item reviewers table:')
cursor.execute('SELECT item_id, reviewer_name, reviewer_email FROM item_reviewers')
for row in cursor.fetchall():
    print(f"  item_id={row['item_id']}: {row['reviewer_name']} ({row['reviewer_email']})")
conn.close()
