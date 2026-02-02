import sqlite3
from datetime import datetime, timedelta

def subtract_business_days(from_date, num_days):
    current = from_date
    days_subtracted = 0
    while days_subtracted < num_days:
        current -= timedelta(days=1)
        if current.weekday() < 5:
            days_subtracted += 1
    return current

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

cursor.execute('''
    SELECT id, identifier, type, date_received, due_date, priority,
           initial_reviewer_due_date, qcr_due_date
    FROM item
    WHERE date_received IS NOT NULL AND due_date IS NOT NULL
''')
items = cursor.fetchall()

print(f'Recalculating due dates for {len(items)} items...')
print()

updated = 0
for item in items:
    date_received = datetime.strptime(item['date_received'][:10], '%Y-%m-%d')
    contractor_due = datetime.strptime(item['due_date'][:10], '%Y-%m-%d')
    
    # QCR due = 1 business day before contractor due
    qcr_due = subtract_business_days(contractor_due, 1)
    
    # Reviewer due = 2 business days before QCR due
    reviewer_due = subtract_business_days(qcr_due, 2)
    
    if reviewer_due < date_received:
        reviewer_due = date_received
    
    new_reviewer_due = reviewer_due.strftime('%Y-%m-%d')
    new_qcr_due = qcr_due.strftime('%Y-%m-%d')
    
    old_reviewer = item['initial_reviewer_due_date'][:10] if item['initial_reviewer_due_date'] else 'None'
    old_qcr = item['qcr_due_date'][:10] if item['qcr_due_date'] else 'None'
    
    if old_reviewer != new_reviewer_due or old_qcr != new_qcr_due:
        print(f"{item['type']} {item['identifier']}: Contractor Due={item['due_date'][:10]}")
        print(f"  Reviewer: {old_reviewer} -> {new_reviewer_due}")
        print(f"  QCR:      {old_qcr} -> {new_qcr_due}")
        
        cursor.execute('''
            UPDATE item SET
                initial_reviewer_due_date = ?,
                qcr_due_date = ?
            WHERE id = ?
        ''', (new_reviewer_due, new_qcr_due, item['id']))
        updated += 1

conn.commit()
conn.close()
print()
print(f'Updated {updated} items')
