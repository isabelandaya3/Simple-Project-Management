"""Check reminder status for a specific submittal."""
import sqlite3
from datetime import datetime, timedelta

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

today = datetime.now().date()
yesterday = today - timedelta(days=1)
print(f"Today: {today}")
print(f"Yesterday: {yesterday}")
print()

# Find the submittal
cursor.execute('''
    SELECT i.*, 
           ir.display_name as initial_reviewer_name,
           qcr.display_name as qcr_name
    FROM item i
    LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
    LEFT JOIN user qcr ON i.qcr_id = qcr.id
    WHERE i.identifier LIKE '%03 20 00-1%'
''')
item = cursor.fetchone()

if item:
    print('=== SUBMITTAL DETAILS ===')
    print(f"ID: {item['id']}")
    print(f"Identifier: {item['identifier']}")
    print(f"Status: {item['status']}")
    print(f"Multi-reviewer mode: {item['multi_reviewer_mode']}")
    due_date = item['initial_reviewer_due_date']
    print(f"Initial Reviewer Due: {due_date}")
    print(f"QCR Due: {item['qcr_due_date']}")
    print(f"Reviewer Response At: {item['reviewer_response_at']}")
    print(f"QCR Response At: {item['qcr_response_at']}")
    print(f"Initial Reviewer ID: {item['initial_reviewer_id']}")
    print(f"Initial Reviewer Name: {item['initial_reviewer_name']}")
    print(f"QCR Name: {item['qcr_name']}")
    print()
    
    # Check item_reviewers table
    cursor.execute('''
        SELECT ir.*, u.display_name, u.email
        FROM item_reviewers ir
        LEFT JOIN user u ON ir.user_id = u.id
        WHERE ir.item_id = ?
    ''', (item['id'],))
    reviewers = cursor.fetchall()
    
    print('=== REVIEWERS ===')
    for r in reviewers:
        print(f"  Reviewer: {r['reviewer_name']} ({r['reviewer_email']})")
        print(f"    User ID: {r['user_id']}")
        print(f"    Email Sent At: {r['email_sent_at']}")
        print(f"    Response At: {r['response_at']}")
        print(f"    Response Category: {r['response_category']}")
        print(f"    Needs Response: {r['needs_response']}")
        print()
    
    # Check reminder_log for the CURRENT due date
    print(f'=== REMINDERS FOR CURRENT DUE DATE ({due_date}) ===')
    cursor.execute('''
        SELECT * FROM reminder_log 
        WHERE item_id = ? AND due_date = ?
        ORDER BY sent_at DESC
    ''', (item['id'], due_date))
    current_reminders = cursor.fetchall()
    if current_reminders:
        for rem in current_reminders:
            print(f"  Sent: {rem['sent_at']}, Stage: {rem['reminder_stage']}, To: {rem['recipient_email']}")
    else:
        print('  No reminders logged for current due date')
    
    # Check ALL reminders
    print()
    print('=== ALL REMINDERS (any due date) ===')
    cursor.execute('''
        SELECT * FROM reminder_log 
        WHERE item_id = ?
        ORDER BY sent_at DESC
    ''', (item['id'],))
    all_reminders = cursor.fetchall()
    if all_reminders:
        for rem in all_reminders:
            print(f"  Sent: {rem['sent_at']}, Due: {rem['due_date']}, Stage: {rem['reminder_stage']}, To: {rem['recipient_email']}")
    else:
        print('  No reminders logged')
else:
    print('Submittal not found')

conn.close()
