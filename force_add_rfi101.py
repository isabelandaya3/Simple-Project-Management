"""Force process RFI 101 email"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import sqlite3
import re
from datetime import datetime
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
    print("Error: win32com not available")
    sys.exit(1)

from app import (
    parse_item_type, parse_identifier, parse_title, parse_due_date, parse_priority,
    parse_rfi_question, is_user_in_rfi_reviewers,
    get_project_by_subject, get_default_project, get_db, CONFIG
)

def force_process_rfi101():
    """Find and force process the RFI 101 email"""
    pythoncom.CoInitialize()
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        print("Looking for RFI 101 email...")
        
        count = 0
        for message in messages:
            count += 1
            if count > 300:
                break
            
            try:
                subject = getattr(message, 'Subject', '') or ''
                
                if 'RFI' in subject and '101' in subject:
                    body = getattr(message, 'Body', '') or ''
                    entry_id = getattr(message, 'EntryID', None)
                    received_time = getattr(message, 'ReceivedTime', None)
                    
                    print(f"Found: {subject[:70]}...")
                    
                    # Check if already processed
                    conn = get_db()
                    cursor = conn.cursor()
                    cursor.execute('SELECT id FROM email_log WHERE entry_id = ?', (entry_id,))
                    if cursor.fetchone():
                        print("Already in email_log, skipping...")
                        conn.close()
                        return
                    
                    # Parse the email
                    project = get_project_by_subject(subject)
                    if not project:
                        if 'LEB' in subject.upper():
                            project = get_default_project()
                        else:
                            print("No project found!")
                            conn.close()
                            return
                    
                    print(f"Project: {project.get('name')}")
                    project_id = project.get('id')
                    
                    item_type = parse_item_type(subject)
                    print(f"Item type: {item_type}")
                    
                    identifier = parse_identifier(subject, item_type)
                    print(f"Identifier: {identifier}")
                    
                    user_names = project.get('user_names', CONFIG.get('user_names', []))
                    if not is_user_in_rfi_reviewers(body, user_names):
                        print(f"User not in reviewers - would skip!")
                        print(f"Checking user_names: {user_names}")
                    else:
                        print("User IS in reviewers - should process")
                    
                    rfi_question = parse_rfi_question(body)
                    print(f"Question: {rfi_question[:80] if rfi_question else 'None'}...")
                    
                    bucket = project.get('matched_bucket', 'ALL')
                    title = parse_title(subject, identifier, body)
                    due_date = parse_due_date(body)
                    priority = parse_priority(body)
                    
                    print(f"Bucket: {bucket}")
                    print(f"Title: {title[:60] if title else 'None'}...")
                    print(f"Due date: {due_date}")
                    print(f"Priority: {priority}")
                    
                    received_at = None
                    if received_time:
                        try:
                            received_at = datetime(
                                received_time.year, received_time.month, received_time.day,
                                received_time.hour, received_time.minute, received_time.second
                            ).isoformat()
                        except:
                            received_at = datetime.now().isoformat()
                    
                    # Check if item exists
                    cursor.execute('''
                        SELECT id FROM item WHERE identifier = ? AND bucket = ?
                    ''', (identifier, bucket))
                    existing = cursor.fetchone()
                    
                    if existing:
                        print(f"Item already exists with id {existing[0]}")
                    else:
                        print("Creating new item...")
                        cursor.execute('''
                            INSERT INTO item (
                                type, identifier, bucket, title, due_date, priority,
                                status, date_received, rfi_question, project_id,
                                source_subject, created_at, email_entry_id
                            ) VALUES (?, ?, ?, ?, ?, ?, 'Unassigned', ?, ?, ?, ?, ?, ?)
                        ''', (item_type, identifier, bucket, title, due_date, priority, 
                              received_at, rfi_question, project_id, subject, 
                              datetime.now().isoformat(), entry_id))
                        item_id = cursor.lastrowid
                        print(f"Created item with id: {item_id}")
                    
                    # Log the email
                    cursor.execute('''
                        INSERT INTO email_log (entry_id, subject, received_at, processed)
                        VALUES (?, ?, ?, 1)
                    ''', (entry_id, subject, datetime.now().isoformat()))
                    
                    conn.commit()
                    conn.close()
                    print("Done! RFI 101 has been added to the tracker.")
                    return
                    
            except Exception as e:
                print(f"Error: {e}")
                import traceback
                traceback.print_exc()
                continue
        
        print(f"Scanned {count} emails - RFI 101 not found")
        
    finally:
        pythoncom.CoUninitialize()

if __name__ == '__main__':
    force_process_rfi101()
