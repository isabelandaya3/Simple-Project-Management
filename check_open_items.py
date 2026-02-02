"""Check open items for missing emails or status issues."""
import sqlite3

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

c.execute('''
    SELECT id, identifier, status, 
           reviewer_response_at, qcr_response_at,
           reviewer_email_sent_at, qcr_email_sent_at,
           qcr_action
    FROM item 
    WHERE status NOT IN ('Closed')
    ORDER BY id
''')
items = c.fetchall()

print('Open Items Status Check:')
print('='*100)

issues_found = []

for item in items:
    rev_resp = 'Y' if item['reviewer_response_at'] else 'N'
    qcr_sent = 'Y' if item['qcr_email_sent_at'] else 'N'
    qcr_resp = item['qcr_action'] or '-'
    status = item['status']
    
    # Check for issues
    issues = []
    if item['reviewer_response_at'] and not item['qcr_email_sent_at'] and status not in ('In Review', 'Unassigned', 'Assigned'):
        issues.append('Reviewer responded but QCR email not sent')
    
    if item['qcr_action'] == 'Send Back' and status != 'In Review':
        issues.append(f'QCR sent back but status is "{status}" not "In Review"')
    
    if item['qcr_action'] in ('Approve', 'Modify') and status != 'Ready for Response':
        issues.append(f'QCR approved but status is "{status}" not "Ready for Response"')
    
    print(f'{item["identifier"]:25} | Status: {status:18} | RevResp: {rev_resp} | QCREmail: {qcr_sent} | QCRAction: {qcr_resp}')
    
    if issues:
        for issue in issues:
            print(f'  !! ISSUE: {issue}')
        issues_found.append((item['id'], item['identifier'], issues))

print('\n' + '='*100)
if issues_found:
    print(f'Found {len(issues_found)} items with issues:')
    for item_id, identifier, item_issues in issues_found:
        print(f'  - {identifier} (ID: {item_id})')
        for issue in item_issues:
            print(f'      {issue}')
else:
    print('No issues found - all items look correct!')

conn.close()
