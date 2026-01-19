#!/usr/bin/env python3
"""Check and process reminders for real items due today."""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import get_items_needing_reminders, process_all_reminders, init_db

init_db()
items = get_items_needing_reminders()

print('Items needing reminders today:')
print('=' * 60)

print(f'\nSingle Reviewer items: {len(items["single_reviewer"])}')
for item, role, due_date, stage in items['single_reviewer']:
    recipient = item['reviewer_email'] if role == 'reviewer' else item['qcr_email']
    print(f'  - {item["identifier"]}: {role} ({recipient}) - {stage} (due {due_date})')

print(f'\nMulti-Reviewer items: {len(items["multi_reviewer"])}')
for item, reviewer, role, due_date, stage in items['multi_reviewer']:
    print(f'  - {item["identifier"]}: {reviewer["reviewer_name"]} ({reviewer["reviewer_email"]}) - {stage} (due {due_date})')

print(f'\nMulti-Reviewer QCR items: {len(items["multi_reviewer_qcr"])}')
for item, due_date, stage in items['multi_reviewer_qcr']:
    print(f'  - {item["identifier"]}: QCR ({item["qcr_email"]}) - {stage} (due {due_date})')

total = len(items['single_reviewer']) + len(items['multi_reviewer']) + len(items['multi_reviewer_qcr'])
print(f'\nTotal items needing reminders: {total}')

if total > 0:
    print('\n' + '=' * 60)
    response = input('Send reminders now? (yes/no): ')
    if response.lower() == 'yes':
        print('\nProcessing reminders...')
        results = process_all_reminders()
        print('\nResults:')
        print(f'  Single reviewer sent: {results["single_reviewer_sent"]}')
        print(f'  Single reviewer skipped (already sent): {results["single_reviewer_skipped"]}')
        print(f'  Multi-reviewer sent: {results["multi_reviewer_sent"]}')
        print(f'  Multi-reviewer skipped: {results["multi_reviewer_skipped"]}')
        print(f'  Multi-reviewer QCR sent: {results["multi_reviewer_qcr_sent"]}')
        print(f'  Multi-reviewer QCR skipped: {results["multi_reviewer_qcr_skipped"]}')
        if results['errors']:
            print(f'\nErrors:')
            for err in results['errors']:
                print(f'  - {err}')
        print('\nDone!')
    else:
        print('Cancelled.')
else:
    print('\nNo reminders needed at this time.')
