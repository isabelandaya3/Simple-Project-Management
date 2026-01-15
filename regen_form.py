import sqlite3
import json
from pathlib import Path

# Config
DB_PATH = 'tracker.db'
TEMPLATE_PATH = Path('templates/_RESPONSE_FORM_TEMPLATE_v3.html')

# Get item data
conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row
cursor = conn.cursor()
cursor.execute('''
    SELECT i.*, 
           ir.display_name as reviewer_name, ir.email as reviewer_email
    FROM item i
    LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
    WHERE i.id = 7
''')
item = dict(cursor.fetchone())
conn.close()

# Read template
with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
    template = f.read()

print(f"Template has 'Select PDF': {'Select PDF' in template}")
print(f"Template has 'Connect to Item': {'Connect to Item' in template}")

# Replace placeholders
html = template.replace('{{ITEM_ID}}', str(item['id']))
html = html.replace('{{ITEM_TYPE}}', item['type'] or '')
html = html.replace('{{ITEM_IDENTIFIER}}', item['identifier'] or '')
html = html.replace('{{ITEM_TITLE}}', (item['title'] or 'N/A').replace('"', ''))
html = html.replace('{{DATE_RECEIVED}}', item['date_received'] or 'N/A')
html = html.replace('{{REVIEWER_DUE_DATE}}', item['initial_reviewer_due_date'] or 'N/A')
html = html.replace('{{QCR_DUE_DATE}}', item['qcr_due_date'] or 'N/A')
html = html.replace('{{CONTRACTOR_DUE_DATE}}', item['due_date'] or 'N/A')
html = html.replace('{{REVIEWER_NAME}}', (item['reviewer_name'] or 'N/A').replace('"', ''))
html = html.replace('{{REVIEWER_EMAIL}}', item['reviewer_email'] or '')
html = html.replace('{{TOKEN}}', item['email_token_reviewer'] or 'test-token')
html = html.replace('{{FOLDER_PATH}}', item['folder_link'] or '')
html = html.replace('{{FOLDER_PATH_URL}}', (item['folder_link'] or '').replace('\\', '/'))
html = html.replace('{{FOLDER_FILES_JSON}}', '[]')

# Save to folder
folder_path = Path(item['folder_link'])
form_path = folder_path / '_RESPONSE_FORM.html'
with open(form_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'Form saved to: {form_path}')
print(f'File size: {form_path.stat().st_size} bytes')
