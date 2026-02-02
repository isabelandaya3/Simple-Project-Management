"""Send manual reminder for RFI #34."""
import sqlite3
import win32com.client
import pythoncom
from datetime import datetime

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

# Get RFI #34 and its reviewer
c.execute('SELECT * FROM item WHERE id = 44')
item = dict(c.fetchone())
c.execute('SELECT * FROM item_reviewers WHERE item_id = 44')
reviewer = dict(c.fetchone())
c.execute('SELECT email, display_name FROM user WHERE id = ?', (item['qcr_id'],))
qcr = dict(c.fetchone())

print(f'Item: {item["identifier"]}')
print(f'Reviewer: {reviewer["reviewer_name"]} ({reviewer["reviewer_email"]})')
print(f'Reviewer Due: {item["initial_reviewer_due_date"]} (OVERDUE)')
print(f'Contractor Due: {item["due_date"]}')
print(f'QCR: {qcr["display_name"]} ({qcr["email"]})')

folder = item['folder_link']
folder_html = f'<a href="file:///{folder.replace(chr(92), "/")}">{folder}</a>' if folder else 'Not set'

# Build form link
from pathlib import Path
form_path = Path(folder) / "Responses" / f"[EOR] RESPONSE FORM - {item['identifier']}.hta"
form_link = f'file:///{str(form_path).replace(chr(92), "/")}'

def fmt(d):
    if not d: return 'Not set'
    try: return datetime.strptime(d[:10], '%Y-%m-%d').strftime('%b %d, %Y')
    except: return d

html = f'''<div style="font-family:Segoe UI,Arial;color:#333;font-size:14px;line-height:1.5;">
<h2 style="color:#dc2626;">⚠️ OVERDUE REMINDER: {item['identifier']}</h2>
<p>Your response for this item is <strong>OVERDUE</strong>. Please submit your response as soon as possible.</p>

<div style="margin:15px 0;padding:15px;background:#fef2f2;border:2px solid #dc2626;border-radius:8px;">
<div style="font-size:15px;color:#991b1b;font-weight:bold;">⚠️ RESPONSE OVERDUE</div>
<div style="font-size:13px;color:#991b1b;margin-top:8px;">Your response was due on <strong>{fmt(item['initial_reviewer_due_date'])}</strong>. Please respond immediately.</div>
</div>

<div style="margin:20px 0;text-align:center;">
<a href="{form_link}" style="background:#dc2626;color:#fff;display:inline-block;font-size:16px;font-weight:bold;line-height:50px;text-align:center;text-decoration:none;width:280px;border-radius:8px;">OPEN RESPONSE FORM</a>
</div>

<table cellpadding="8" cellspacing="0" width="100%" style="border-collapse:collapse;margin-top:15px;">
<tr><td colspan="2" style="background:#f2f2f2;font-weight:bold;border:1px solid #ddd;">Item Information</td></tr>
<tr><td style="width:160px;border:1px solid #ddd;font-weight:bold;">Type</td><td style="border:1px solid #ddd;">{item['type']}</td></tr>
<tr><td style="border:1px solid #ddd;font-weight:bold;">Identifier</td><td style="border:1px solid #ddd;">{item['identifier']}</td></tr>
<tr><td style="border:1px solid #ddd;font-weight:bold;">Title</td><td style="border:1px solid #ddd;">{item['title'] or 'N/A'}</td></tr>
<tr><td style="border:1px solid #ddd;font-weight:bold;">Your Due Date</td><td style="border:1px solid #ddd;color:#dc2626;font-weight:bold;">{fmt(item['initial_reviewer_due_date'])} (OVERDUE)</td></tr>
<tr><td style="border:1px solid #ddd;font-weight:bold;">Contractor Due Date</td><td style="border:1px solid #ddd;color:#c0392b;font-weight:bold;">{fmt(item['due_date'])}</td></tr>
</table>

<div style="margin:15px 0;padding:15px;background:#e0e7ff;border:2px solid #4f46e5;border-radius:8px;">
<div style="font-size:14px;color:#3730a3;font-weight:bold;">📝 RFI Question:</div>
<div style="margin-top:8px;font-size:13px;color:#312e81;white-space:pre-wrap;">{item['rfi_question'] if item['rfi_question'] else 'See attached documents'}</div>
</div>

<div style="margin-top:18px;">
<div style="font-weight:bold;margin-bottom:4px;">📁 Item Folder:</div>
<div style="padding:10px;border:1px solid #ddd;background:#fafafa;font-family:Consolas;font-size:12px;border-radius:4px;">{folder_html}</div>
</div>
</div>'''

pythoncom.CoInitialize()
try:
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = reviewer['reviewer_email']
    mail.CC = qcr['email']  # CC the QCR
    mail.Subject = f'[LEB] {item["identifier"]} - OVERDUE: Response Required'
    mail.HTMLBody = html
    mail.Send()
    print(f'\nReminder sent to {reviewer["reviewer_email"]} (CC: {qcr["email"]})')
except Exception as e:
    print(f'Error: {e}')
finally:
    pythoncom.CoUninitialize()

conn.close()
