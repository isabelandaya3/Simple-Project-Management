"""Send sample workflow restart email for verification."""
import sqlite3
import win32com.client
import pythoncom
from datetime import datetime
from pathlib import Path
import sys

# Add the current directory to path so we can import from app
sys.path.insert(0, str(Path(__file__).parent))

def format_date_for_email(date_str):
    if not date_str:
        return 'Not set'
    try:
        dt = datetime.strptime(date_str[:10], '%Y-%m-%d')
        return dt.strftime('%b %d, %Y')
    except:
        return date_str

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

item_id = 36  # Submittal #4
cursor.execute('SELECT * FROM item WHERE id = ?', (item_id,))
item = dict(cursor.fetchone())
conn.close()

admin_note = 'Contractor is adding additional drawing to this submittal for review'
was_closed = True

folder_path = item['folder_link'] or 'Not set'
if folder_path != 'Not set':
    folder_link_html = f'<a href="file:///{folder_path.replace(chr(92), "/")}" style="color:#0078D4;">{folder_path}</a>'
else:
    folder_link_html = 'Not set'

# Generate form file link using the versioned folder
form_file_link = ''
if folder_path != 'Not set':
    # Determine the responses folder name
    reopen_count = item.get('reopen_count') or 0
    if reopen_count > 0:
        responses_folder_name = f"Responses R{reopen_count + 1}"
    else:
        responses_folder_name = "Responses"
    
    # Build path to response form
    form_path = Path(folder_path) / responses_folder_name / f"[EOR] RESPONSE FORM - {item['identifier']}.hta"
    
    # Check if form already exists
    if form_path.exists():
        form_file_link = f'file:///{str(form_path).replace(chr(92), "/")}'
        print(f"Form exists at: {form_path}")
    else:
        print(f"Note: Form would be generated at: {form_path}")
        print("(Form will be created when 'Restart Workflow' action is executed)")

# Determine the responses folder name for display
reopen_count = item.get('reopen_count') or 0
if reopen_count > 0:
    responses_folder_name = f"Responses R{reopen_count + 1}"
else:
    responses_folder_name = "Responses"
print(f"Response forms will be in: {responses_folder_name}")

# Build previous response section
previous_response_html = ''
if was_closed and item['final_response_category']:
    previous_response_html = f'''
    <!-- PREVIOUS RESPONSE -->
    <div style="margin:15px 0; padding:15px; background:#f0fdf4; border:2px solid #22c55e; border-radius:8px;">
        <div style="font-size:14px; color:#166534; font-weight:bold;">📄 Our Previous Response (when item was closed):</div>
        <div style="margin-top:10px;">
            <div style="font-size:13px; color:#166534;">
                <strong>Response Category:</strong> {item['final_response_category']}
            </div>
            <div style="font-size:13px; color:#166534; margin-top:6px;"><strong>Comments:</strong> {item["final_response_text"] or "None"}</div>
            <div style="font-size:13px; color:#166534; margin-top:6px;"><strong>Files:</strong> {item["final_response_files"] or "None"}</div>
        </div>
    </div>'''

status_msg = 'This item was previously <strong>CLOSED</strong> but has been reopened due to contractor changes.'
header_color = '#dc2626'
icon = '🔄⚠️'
priority_color = '#27ae60'

subject = f"[LEB] {item['identifier']} – REOPENED: Review Restart Required (SAMPLE v2)"

# Build due date display with change indicator
due_date_display = format_date_for_email(item['due_date'])
if item['previous_due_date'] and item['previous_due_date'] != item['due_date']:
    due_date_display = f"{format_date_for_email(item['previous_due_date'])} → {format_date_for_email(item['due_date'])}"

# Build form button HTML
form_button_html = ''
if form_file_link:
    form_button_html = f'''
    <!-- ACTION BUTTON -->
    <div style="margin:20px 0; text-align:center;">
        <table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;">
            <tr>
                <td align="center" bgcolor="#dc2626" style="background:#dc2626; border-radius:8px; padding:0;">
                    <a href="{form_file_link}" target="_blank"
                       style="background:#dc2626; color:#ffffff; display:inline-block; font-family:Segoe UI,Arial,sans-serif;
                              font-size:16px; font-weight:bold; line-height:50px; text-align:center; text-decoration:none;
                              width:280px; -webkit-text-size-adjust:none; border-radius:8px;">
                        OPEN NEW RESPONSE FORM
                    </a>
                </td>
            </tr>
        </table>
    </div>'''

html_body = f'''<div style="font-family:Segoe UI, Helvetica, Arial, sans-serif; color:#333; font-size:14px; line-height:1.5;">

    <!-- HEADER -->
    <h2 style="color:{header_color}; margin-bottom:6px;">
        {icon} ITEM REOPENED: {item['identifier']}
    </h2>

    <p style="margin-top:0; font-size:13px; color:#666;">
        {status_msg}
    </p>

    <!-- CHANGE NOTICE -->
    <div style="margin:15px 0; padding:15px; background:#fef3c7; border:2px solid #f59e0b; border-radius:8px;">
        <div style="font-size:15px; color:#92400e; font-weight:bold;">
            ⚠️ ITEM REOPENED - NEW REVIEW REQUIRED
        </div>
        <div style="font-size:13px; color:#92400e; margin-top:8px;">
            Your previous response has been cleared. Please review the updated materials and submit a new response.
        </div>
    </div>
    
    <!-- ADMIN NOTE -->
    <div style="margin:15px 0; padding:15px; background:#e0e7ff; border:2px solid #4f46e5; border-radius:8px;">
        <div style="font-size:14px; color:#3730a3; font-weight:bold;">📋 What Changed (Note from Administrator):</div>
        <div style="margin-top:8px; font-size:13px; color:#312e81;">{admin_note}</div>
    </div>
    
    {previous_response_html}
    
    {form_button_html}

    <!-- INFO TABLE -->
    <table cellpadding="8" cellspacing="0" width="100%" style="border-collapse:collapse; margin-top:15px;">
        <tr>
            <td colspan="2" style="background:#f2f2f2; font-weight:bold; border:1px solid #ddd;">
                Item Information
            </td>
        </tr>
        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Type</td>
            <td style="border:1px solid #ddd;">{item['type']}</td>
        </tr>
        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Identifier</td>
            <td style="border:1px solid #ddd;">{item['identifier']}</td>
        </tr>
        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Title</td>
            <td style="border:1px solid #ddd;">{item['title'] or 'N/A'}</td>
        </tr>
        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Priority</td>
            <td style="border:1px solid #ddd; color:{priority_color}; font-weight:bold;">{item['priority'] or 'Normal'}</td>
        </tr>
        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Your Due Date</td>
            <td style="border:1px solid #ddd; color:#2980b9; font-weight:bold;">{format_date_for_email(item['initial_reviewer_due_date'])}</td>
        </tr>
        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Contractor Due Date</td>
            <td style="border:1px solid #ddd; color:#c0392b; font-weight:bold;">
                {due_date_display}
            </td>
        </tr>
    </table>

    <!-- FOLDER LINK -->
    <div style="margin-top:18px;">
        <div style="font-weight:bold; margin-bottom:4px;">📁 Item Folder (review updated materials here):</div>
        <div style="padding:10px; border:1px solid #ddd; background:#fafafa; font-family:Consolas, monospace; font-size:12px; border-radius:4px;">
            {folder_link_html}
        </div>
    </div>

    <!-- FOOTER -->
    <p style="margin-top:20px; font-size:12px; color:#777;">
        <em>This is a SAMPLE email for verification purposes. Response forms are in the "{responses_folder_name}" folder.</em>
    </p>

</div>'''

pythoncom.CoInitialize()
try:
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = 'isabel.andaya@hdrinc.com'
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.Send()
    print('Sample email sent to isabel.andaya@hdrinc.com (1 email only)')
except Exception as e:
    print(f'Error: {e}')
finally:
    pythoncom.CoUninitialize()

