"""Generate response form for Submittal #4 with correct tokens."""
import sqlite3
import json
from pathlib import Path

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

# Get item details
c.execute('''
    SELECT i.*, 
           ir.display_name as reviewer_name, ir.email as reviewer_email,
           qcr.display_name as qcr_name, qcr.email as qcr_email
    FROM item i
    LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
    LEFT JOIN user qcr ON i.qcr_id = qcr.id
    WHERE i.id = 36
''')
item = dict(c.fetchone())

# Get the reviewer from item_reviewers table
c.execute('SELECT * FROM item_reviewers WHERE item_id = 36')
reviewer_row = c.fetchone()
reviewer = dict(reviewer_row) if reviewer_row else None

print("Item tokens:")
print(f"  email_token_reviewer: {item['email_token_reviewer']}")
print(f"  email_token_qcr: {item['email_token_qcr']}")
if reviewer:
    print(f"  item_reviewers token for {reviewer['reviewer_name']}: {reviewer['email_token']}")

# The reviewer form should use the token from item_reviewers for this reviewer
token = reviewer['email_token'] if reviewer else item['email_token_reviewer']
print(f"\nUsing token: {token}")

# Load template
template_path = Path('templates/_RESPONSE_FORM_TEMPLATE_v3.hta')
if not template_path.exists():
    print(f"Template not found: {template_path}")
    exit(1)

with open(template_path, 'r', encoding='utf-8') as f:
    template = f.read()

def js_escape(s):
    if not s:
        return ''
    return s.replace('\\', '\\\\').replace('"', '\\"').replace("'", "\\'").replace('\n', '\\n').replace('\r', '')

def html_escape(s):
    if not s:
        return ''
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

folder_path_url = str(item['folder_link']).replace('\\', '/')

# Replace placeholders
html = template.replace('{{ITEM_ID}}', str(item['id']))
html = html.replace('{{ITEM_TYPE}}', item['type'] or '')
html = html.replace('{{ITEM_IDENTIFIER}}', item['identifier'] or '')
html = html.replace('{{ITEM_TITLE}}', js_escape(item['title']) or 'N/A')
html = html.replace('{{ITEM_TITLE_HTML}}', html_escape(item['title'] or 'N/A'))
html = html.replace('{{DATE_RECEIVED}}', item['date_received'] or 'N/A')
html = html.replace('{{REVIEWER_DUE_DATE}}', item['initial_reviewer_due_date'] or 'N/A')
html = html.replace('{{QCR_DUE_DATE}}', item['qcr_due_date'] or 'N/A')
html = html.replace('{{CONTRACTOR_DUE_DATE}}', item['due_date'] or 'N/A')
html = html.replace('{{REVIEWER_NAME}}', js_escape(reviewer['reviewer_name']) if reviewer else 'N/A')
html = html.replace('{{REVIEWER_EMAIL}}', reviewer['reviewer_email'] if reviewer else '')
html = html.replace('{{TOKEN}}', token)
html = html.replace('{{FOLDER_PATH}}', js_escape(item['folder_link']) or '')
html = html.replace('{{FOLDER_PATH_RAW}}', html_escape(item['folder_link'] or ''))
html = html.replace('{{FOLDER_PATH_URL}}', folder_path_url)
html = html.replace('{{FOLDER_FILES_JSON}}', json.dumps([]))
html = html.replace('{{RFI_QUESTION}}', js_escape(item.get('rfi_question', '') or ''))
html = html.replace('{{IS_RFI}}', 'true' if (item['type'] or '').upper() == 'RFI' else 'false')

# Save to Responses R2 folder
folder_path = Path(item['folder_link'])
reopen_count = item.get('reopen_count', 0) or 0
if reopen_count > 0:
    responses_folder = folder_path / f"Responses R{reopen_count + 1}"
else:
    responses_folder = folder_path / "Responses"

# Create folder if needed
responses_folder.mkdir(exist_ok=True)

form_path = responses_folder / f"[EOR] RESPONSE FORM - {item['identifier']}.hta"
with open(form_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"\nForm created at: {form_path}")
print(f"Form exists: {form_path.exists()}")

conn.close()
