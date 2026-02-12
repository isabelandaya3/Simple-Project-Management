"""Test script to trigger email poll manually and debug RFI 101"""
import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import from app.py
import sqlite3
import re
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
    print("Warning: win32com not available")

from app import (
    parse_item_type, parse_identifier, parse_rfi_question,
    is_user_in_rfi_reviewers, parse_rfi_reviewers, parse_rfi_coReviewers,
    CONFIG
)

def manual_scan_for_rfi101():
    """Manually scan Outlook for RFI 101 email"""
    if not HAS_WIN32COM:
        print("Cannot scan - win32com not available")
        return
    
    pythoncom.CoInitialize()
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # Inbox
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        print(f"Scanning inbox for RFI 101...")
        print(f"User names in config: {CONFIG.get('user_names', [])}")
        
        count = 0
        for message in messages:
            count += 1
            if count > 300:
                break
            
            try:
                subject = getattr(message, 'Subject', '') or ''
                
                # Look for RFI 101
                if 'RFI' in subject and '101' in subject:
                    body = getattr(message, 'Body', '') or ''
                    sender = getattr(message, 'SenderEmailAddress', '') or ''
                    received = getattr(message, 'ReceivedTime', None)
                    entry_id = getattr(message, 'EntryID', None)
                    
                    print(f"\n=== Found RFI 101 email ===")
                    print(f"Subject: {subject}")
                    print(f"Sender: {sender}")
                    print(f"Received: {received}")
                    print(f"Entry ID: {entry_id[:30]}..." if entry_id else "No Entry ID")
                    
                    # Test parsing
                    item_type = parse_item_type(subject)
                    print(f"\nItem type: {item_type}")
                    
                    identifier = parse_identifier(subject, item_type)
                    print(f"Identifier: {identifier}")
                    
                    # Test reviewer parsing
                    reviewers = parse_rfi_reviewers(body)
                    print(f"Reviewers: {reviewers}")
                    
                    coReviewers = parse_rfi_coReviewers(body)
                    print(f"Co-reviewers: {coReviewers}")
                    
                    user_names = CONFIG.get('user_names', [])
                    is_user = is_user_in_rfi_reviewers(body, user_names)
                    print(f"Is user in reviewers: {is_user}")
                    
                    # Check if already in email_log
                    conn = sqlite3.connect('tracker.db')
                    c = conn.cursor()
                    c.execute('SELECT id FROM email_log WHERE entry_id = ?', (entry_id,))
                    existing = c.fetchone()
                    print(f"Already in email_log: {existing is not None}")
                    conn.close()
                    
                    return  # Found it, stop scanning
                    
            except Exception as e:
                print(f"Error processing message: {e}")
                continue
        
        print(f"\nScanned {count} emails - RFI 101 not found in inbox")
        
    finally:
        pythoncom.CoUninitialize()

if __name__ == '__main__':
    manual_scan_for_rfi101()
