#!/usr/bin/env python3
"""Check the status of items in the database."""

import sqlite3
from datetime import datetime, timedelta

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

today = datetime.now().date()
tomorrow = today + timedelta(days=1)
yesterday = today - timedelta(days=1)

print(f'Today: {today}')
print(f'Tomorrow: {tomorrow}')
print(f'Yesterday: {yesterday}')
print()

# Check items with due dates around today
print('All items with reviewer/QCR due dates:')
print('=' * 80)
c.execute('''
    SELECT i.id, i.identifier, i.status, i.multi_reviewer_mode,
           i.initial_reviewer_due_date, i.qcr_due_date,
           i.reviewer_email_sent_at, i.reviewer_response_at,
           i.qcr_email_sent_at, i.qcr_response_at,
           ir.display_name as reviewer_name, ir.email as reviewer_email,
           qcr.display_name as qcr_name, qcr.email as qcr_email
    FROM item i
    LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
    LEFT JOIN user qcr ON i.qcr_id = qcr.id
    WHERE i.identifier NOT LIKE 'TEST-REMINDER-%'
    AND (i.initial_reviewer_due_date IS NOT NULL OR i.qcr_due_date IS NOT NULL)
    ORDER BY i.initial_reviewer_due_date
''')

for row in c.fetchall():
    row = dict(row)
    print(f'ID {row["id"]}: {row["identifier"]}')
    print(f'  Status: {row["status"]}, Multi-reviewer: {row["multi_reviewer_mode"]}')
    print(f'  Reviewer Due: {row["initial_reviewer_due_date"]}, QCR Due: {row["qcr_due_date"]}')
    print(f'  Reviewer: {row["reviewer_name"]} ({row["reviewer_email"]})')
    print(f'  QCR: {row["qcr_name"]} ({row["qcr_email"]})')
    print(f'  Reviewer email sent: {row["reviewer_email_sent_at"] is not None}, Reviewer responded: {row["reviewer_response_at"] is not None}')
    print(f'  QCR email sent: {row["qcr_email_sent_at"] is not None}, QCR responded: {row["qcr_response_at"] is not None}')
    print()

# Check multi-reviewer items
print()
print('Multi-reviewer details:')
print('=' * 80)
c.execute('''
    SELECT ir.item_id, ir.reviewer_name, ir.reviewer_email, 
           ir.email_sent_at, ir.response_at, ir.needs_response,
           i.identifier, i.status, i.initial_reviewer_due_date
    FROM item_reviewers ir
    JOIN item i ON ir.item_id = i.id
    WHERE i.identifier NOT LIKE 'TEST-REMINDER-%'
''')

rows = c.fetchall()
if rows:
    for row in rows:
        row = dict(row)
        print(f'Item {row["item_id"]} ({row["identifier"]}): {row["reviewer_name"]} ({row["reviewer_email"]})')
        print(f'  Due: {row["initial_reviewer_due_date"]}, Status: {row["status"]}')
        print(f'  Email sent: {row["email_sent_at"] is not None}, Responded: {row["response_at"] is not None}, Needs response: {row["needs_response"]}')
        print()
else:
    print('No multi-reviewer items found.')

# Summary of items needing action
print()
print('=' * 80)
print('ITEMS THAT COULD RECEIVE REMINDERS (if due date matches):')
print('=' * 80)

# Single reviewer - waiting for reviewer
c.execute('''
    SELECT i.id, i.identifier, i.status, i.initial_reviewer_due_date,
           u.display_name, u.email
    FROM item i
    JOIN user u ON i.initial_reviewer_id = u.id
    WHERE i.multi_reviewer_mode = 0 
    AND i.status IN ('Assigned', 'In Review')
    AND i.reviewer_email_sent_at IS NOT NULL
    AND i.reviewer_response_at IS NULL
    AND i.identifier NOT LIKE 'TEST-%'
''')

print('\nSingle Reviewer - Waiting for reviewer response:')
for row in c.fetchall():
    row = dict(row)
    due = row['initial_reviewer_due_date']
    status = 'DUE TODAY' if str(due) == str(today) else ('OVERDUE' if str(due) == str(yesterday) else f'due {due}')
    print(f'  {row["identifier"]}: {row["display_name"]} ({row["email"]}) - {status}')

# Multi-reviewer items
c.execute('''
    SELECT ir.item_id, i.identifier, i.initial_reviewer_due_date,
           ir.reviewer_name, ir.reviewer_email, ir.response_at
    FROM item_reviewers ir
    JOIN item i ON ir.item_id = i.id
    WHERE i.multi_reviewer_mode = 1
    AND ir.email_sent_at IS NOT NULL
    AND ir.response_at IS NULL
    AND ir.needs_response = 1
    AND i.identifier NOT LIKE 'TEST-%'
''')

print('\nMulti-Reviewer - Waiting for individual reviewer responses:')
for row in c.fetchall():
    row = dict(row)
    due = row['initial_reviewer_due_date']
    status = 'DUE TODAY' if str(due) == str(today) else ('OVERDUE' if str(due) == str(yesterday) else f'due {due}')
    print(f'  {row["identifier"]}: {row["reviewer_name"]} ({row["reviewer_email"]}) - {status}')

conn.close()
