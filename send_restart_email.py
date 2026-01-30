"""Send restart workflow email for Submittal #4."""
import sqlite3
import win32com.client
import pythoncom
from datetime import datetime
from pathlib import Path
import sys
sys.path.insert(0, '.')

def fmt(d):
    if not d: return 'Not set'
    try: return datetime.strptime(d[:10], '%Y-%m-%d').strftime('%b %d, %Y')
    except: return d

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()
c.execute('SELECT * FROM item WHERE id = 36')
item = dict(c.fetchone())
c.execute('SELECT * FROM item_reviewers WHERE item_id = 36')
reviewer = dict(c.fetchone())
c.execute('SELECT email FROM user WHERE id = ?', (item['qcr_id'],))
qcr = c.fetchone()
qcr_email = qcr['email'] if qcr else None
conn.close()

folder = item['folder_link']
folder_html = f'<a href="file:///{folder.replace(chr(92), "/")}">{folder}</a>' if folder else 'Not set'
reopen = item.get('reopen_count') or 0
resp_folder = f'Responses R{reopen + 1}' if reopen > 0 else 'Responses'

# Generate the form using app.py function
from app import generate_reviewer_form_html
form_result = generate_reviewer_form_html(36)
if form_result.get('success'):
    form_link = f'file:///{form_result["path"].replace(chr(92), "/")}'
    print(f"Form generated at: {form_result['path']}")
else:
    print(f"Form generation failed: {form_result.get('error')}")
    form_link = None

admin_note = 'Contractor is adding additional drawing to this submittal for review'

html = f'''<div style="font-family:Segoe UI,Arial;color:#333;font-size:14px;line-height:1.5;">
<h2 style="color:#dc2626;">🔄⚠️ ITEM REOPENED: {item['identifier']}</h2>
<p>This item was previously <strong>CLOSED</strong> but has been reopened due to contractor changes.</p>

<div style="margin:15px 0;padding:15px;background:#fef3c7;border:2px solid #f59e0b;border-radius:8px;">
<div style="font-size:15px;color:#92400e;font-weight:bold;">⚠️ ITEM REOPENED - NEW REVIEW REQUIRED</div>
<div style="font-size:13px;color:#92400e;margin-top:8px;">Your previous response has been cleared. Please review the updated materials and submit a new response.</div>
</div>

<div style="margin:15px 0;padding:15px;background:#e0e7ff;border:2px solid #4f46e5;border-radius:8px;">
<div style="font-size:14px;color:#3730a3;font-weight:bold;">📋 What Changed:</div>
<div style="margin-top:8px;font-size:13px;color:#312e81;">{admin_note}</div>
</div>

<div style="margin:20px 0;text-align:center;">
<a href="{form_link}" style="background:#dc2626;color:#fff;display:inline-block;font-size:16px;font-weight:bold;line-height:50px;text-align:center;text-decoration:none;width:280px;border-radius:8px;">OPEN NEW RESPONSE FORM</a>
</div>

<table cellpadding="8" cellspacing="0" width="100%" style="border-collapse:collapse;margin-top:15px;">
<tr><td colspan="2" style="background:#f2f2f2;font-weight:bold;border:1px solid #ddd;">Item Information</td></tr>
<tr><td style="width:160px;border:1px solid #ddd;font-weight:bold;">Type</td><td style="border:1px solid #ddd;">{item['type']}</td></tr>
<tr><td style="border:1px solid #ddd;font-weight:bold;">Identifier</td><td style="border:1px solid #ddd;">{item['identifier']}</td></tr>
<tr><td style="border:1px solid #ddd;font-weight:bold;">Title</td><td style="border:1px solid #ddd;">{item['title'] or 'N/A'}</td></tr>
<tr><td style="border:1px solid #ddd;font-weight:bold;">Your Due Date</td><td style="border:1px solid #ddd;color:#2980b9;font-weight:bold;">{fmt(item['initial_reviewer_due_date'])}</td></tr>
<tr><td style="border:1px solid #ddd;font-weight:bold;">Contractor Due Date</td><td style="border:1px solid #ddd;color:#c0392b;font-weight:bold;">{fmt(item['previous_due_date'])} → {fmt(item['due_date'])}</td></tr>
</table>

<div style="margin-top:18px;">
<div style="font-weight:bold;margin-bottom:4px;">📁 Item Folder:</div>
<div style="padding:10px;border:1px solid #ddd;background:#fafafa;font-family:Consolas;font-size:12px;border-radius:4px;">{folder_html}</div>
</div>
<p style="margin-top:20px;font-size:12px;color:#777;"><em>Response forms are in the "{resp_folder}" folder.</em></p>
</div>'''

pythoncom.CoInitialize()
try:
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = reviewer['reviewer_email']
    if qcr_email:
        mail.CC = qcr_email
    mail.Subject = f'[LEB] {item["identifier"]} - REOPENED: Review Restart Required'
    mail.HTMLBody = html
    mail.Send()
    print(f'Email sent to {reviewer["reviewer_email"]} (CC: {qcr_email})')
except Exception as e:
    print(f'Error: {e}')
finally:
    pythoncom.CoUninitialize()
