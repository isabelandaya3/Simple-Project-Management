import sqlite3

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

# Check items with assigned status
c.execute('''SELECT i.id, i.status, u.display_name as reviewer 
             FROM item i 
             LEFT JOIN user u ON i.initial_reviewer_id = u.id 
             WHERE i.status IN ("Assigned", "In Review", "In QC")
             LIMIT 5''')
items = [dict(r) for r in c.fetchall()]
print('Items with assigned status:')
for i in items:
    reviewer = i['reviewer'] or 'None'
    # Test last name extraction
    if reviewer and ',' in reviewer:
        last_name = reviewer.split(',')[0].strip()
    elif reviewer:
        last_name = reviewer.split()[0] if ' ' in reviewer else reviewer
    else:
        last_name = 'Reviewer'
    print(f"  {i['id']}: {i['status']} - reviewer: {reviewer} -> ball: {last_name}")

conn.close()
