import sqlite3
conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Check item 53
cursor.execute('''
    SELECT i.id, i.identifier, i.initial_reviewer_id, i.multi_reviewer_mode, i.folder_link,
           u.display_name as reviewer_from_user
    FROM item i
    LEFT JOIN user u ON i.initial_reviewer_id = u.id
    WHERE i.id = 53
''')
row = cursor.fetchone()
print('Item 53:')
print(f"  identifier: {row['identifier']}")
print(f"  initial_reviewer_id: {row['initial_reviewer_id']}")
print(f"  multi_mode: {row['multi_reviewer_mode']}")
print(f"  reviewer_from_user: {row['reviewer_from_user']}")
print(f"  folder_link: {row['folder_link']}")

print()
cursor.execute('SELECT reviewer_name FROM item_reviewers WHERE item_id = 53')
rows = cursor.fetchall()
print(f"item_reviewers for 53: {[r['reviewer_name'] for r in rows]}")
conn.close()
