"""Check what reminders are needed."""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import get_items_needing_reminders, init_db

init_db()
items = get_items_needing_reminders()

print('Items needing reminders today:')
print('=' * 60)

print(f'Single Reviewer items: {len(items["single_reviewer"])}')
for item, role, due_date, stage in items['single_reviewer']:
    recipient = item['reviewer_email'] if role == 'reviewer' else item['qcr_email']
    print(f'  - {item["identifier"]} ({item["type"]}): {role} ({recipient}) - {stage} (due {due_date})')

print(f'\nMulti-Reviewer items: {len(items["multi_reviewer"])}')
for item, reviewer, role, due_date, stage in items['multi_reviewer']:
    print(f'  - {item["identifier"]} ({item["type"]}): {reviewer["reviewer_name"]} ({reviewer["reviewer_email"]}) - {stage} (due {due_date})')

print(f'\nMulti-Reviewer QCR items: {len(items["multi_reviewer_qcr"])}')
for item, due_date, stage in items['multi_reviewer_qcr']:
    print(f'  - {item["identifier"]} ({item["type"]}): QCR ({item["qcr_email"]}) - {stage} (due {due_date})')

total = len(items['single_reviewer']) + len(items['multi_reviewer']) + len(items['multi_reviewer_qcr'])
print(f'\nTotal items needing reminders: {total}')
