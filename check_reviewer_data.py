import sqlite3

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Check item_reviewers table
cursor.execute('SELECT * FROM item_reviewers WHERE item_id = 37')
reviewers = cursor.fetchall()
print('Reviewers in item_reviewers table:')
for r in reviewers:
    print(f'  - {r["reviewer_name"]} ({r["reviewer_email"]})')

# Check item table
cursor.execute('SELECT initial_reviewer_id FROM item WHERE id = 37')
item = cursor.fetchone()
print(f'\nInitial reviewer ID in item table: {item["initial_reviewer_id"]}')

# Check if the name shows up in the query
cursor.execute('''
    SELECT i.id, i.identifier,
           CASE 
               WHEN EXISTS (SELECT 1 FROM item_reviewers WHERE item_id = i.id) THEN (
                   SELECT GROUP_CONCAT(reviewer_name, ', ') 
                   FROM item_reviewers 
                   WHERE item_id = i.id
               )
               ELSE ir.display_name 
           END as initial_reviewer_name
    FROM item i
    LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
    WHERE i.id = 37
''')
result = cursor.fetchone()
print(f'\nQuery result for initial_reviewer_name: "{result["initial_reviewer_name"]}"')

conn.close()
