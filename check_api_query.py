import sqlite3
import json

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Run the exact query from the API
cursor.execute('''
    SELECT i.*, 
           u.display_name as assigned_to_name,
           CASE 
               WHEN EXISTS (SELECT 1 FROM item_reviewers WHERE item_id = i.id) THEN (
                   SELECT GROUP_CONCAT(reviewer_name, ', ') 
                   FROM item_reviewers 
                   WHERE item_id = i.id
               )
               ELSE ir.display_name 
           END as initial_reviewer_name,
           qcr.display_name as qcr_name
    FROM item i
    LEFT JOIN user u ON i.assigned_to_user_id = u.id
    LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
    LEFT JOIN user qcr ON i.qcr_id = qcr.id
    WHERE i.id = 37
''')

item = dict(cursor.fetchone())
print(f"Item ID: {item['id']}")
print(f"Identifier: {item['identifier']}")
print(f"Initial Reviewer Name: '{item['initial_reviewer_name']}'")
print(f"QCR Name: '{item['qcr_name']}'")
print(f"Status: {item['status']}")

conn.close()
