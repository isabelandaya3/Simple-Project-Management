#!/usr/bin/env python3
"""
Comprehensive test script for the reminder email system.
Tests all 8 reminder scenarios:
1. Single Reviewer - Reviewer Due Today
2. Single Reviewer - Reviewer Overdue
3. Single Reviewer - QCR Due Today
4. Single Reviewer - QCR Overdue
5. Multi-Reviewer - Individual Reviewer Due Today
6. Multi-Reviewer - Individual Reviewer Overdue
7. Multi-Reviewer - QCR Due Today
8. Multi-Reviewer - QCR Overdue
"""

import sys
import os
import sqlite3
from datetime import datetime, timedelta

# Add the app directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import (
    get_db, init_db,
    send_single_reviewer_reminder_email,
    send_multi_reviewer_reminder_email,
    send_multi_reviewer_qcr_reminder_email,
    generate_token
)

# Test email addresses
EMAIL_1 = "isabel.andaya@hdrinc.com"
EMAIL_2 = "isabelandaya@gmail.com"
EMAIL_3 = "iandaya@scu.edu"

# Today's date
TODAY = datetime.now().date()
YESTERDAY = TODAY - timedelta(days=1)

def cleanup_test_items():
    """Remove any previous test items."""
    conn = get_db()
    cursor = conn.cursor()
    
    # Delete test items by identifier pattern
    cursor.execute("DELETE FROM item_reviewers WHERE item_id IN (SELECT id FROM item WHERE identifier LIKE 'TEST-REMINDER-%')")
    cursor.execute("DELETE FROM reminder_log WHERE item_id IN (SELECT id FROM item WHERE identifier LIKE 'TEST-REMINDER-%')")
    cursor.execute("DELETE FROM item WHERE identifier LIKE 'TEST-REMINDER-%'")
    
    conn.commit()
    conn.close()
    print("Cleaned up previous test items")

def ensure_test_users():
    """Ensure test users exist in the database."""
    conn = get_db()
    cursor = conn.cursor()
    
    test_users = [
        (EMAIL_1, 'Isabel Andaya (HDR)'),
        (EMAIL_2, 'Isabel Andaya (Gmail)'),
        (EMAIL_3, 'Isabel Andaya (SCU)')
    ]
    
    user_ids = {}
    for email, name in test_users:
        cursor.execute('SELECT id FROM user WHERE email = ?', (email,))
        result = cursor.fetchone()
        if result:
            user_ids[email] = result[0]
        else:
            cursor.execute('INSERT INTO user (email, display_name, role) VALUES (?, ?, ?)', 
                          (email, name, 'user'))
            user_ids[email] = cursor.lastrowid
    
    conn.commit()
    conn.close()
    print(f"Test users ready: {user_ids}")
    return user_ids

def create_test_item(identifier, title, reviewer_email, reviewer_name, qcr_email, qcr_name,
                     reviewer_due_date, qcr_due_date, status, multi_reviewer_mode=False,
                     reviewer_response_at=None, qcr_response_at=None,
                     reviewer_email_sent_at=None, qcr_email_sent_at=None):
    """Create a test item in the database."""
    conn = get_db()
    cursor = conn.cursor()
    
    # Get or create user IDs
    cursor.execute('SELECT id FROM user WHERE email = ?', (reviewer_email,))
    reviewer_id = cursor.fetchone()
    reviewer_id = reviewer_id[0] if reviewer_id else None
    
    cursor.execute('SELECT id FROM user WHERE email = ?', (qcr_email,))
    qcr_id = cursor.fetchone()
    qcr_id = qcr_id[0] if qcr_id else None
    
    # Generate tokens
    reviewer_token = generate_token()
    qcr_token = generate_token()
    
    # Set email sent times if not provided
    if reviewer_email_sent_at is None:
        reviewer_email_sent_at = datetime.now().isoformat()
    if qcr_email_sent_at is None and status == 'In QC':
        qcr_email_sent_at = datetime.now().isoformat()
    
    # Create item
    cursor.execute('''
        INSERT INTO item (
            type, bucket, identifier, title, status, 
            date_received, due_date, priority,
            initial_reviewer_id, qcr_id,
            initial_reviewer_due_date, qcr_due_date,
            email_token_reviewer, email_token_qcr,
            reviewer_email_sent_at, reviewer_response_at,
            qcr_email_sent_at, qcr_response_at,
            multi_reviewer_mode,
            folder_link
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        'Submittal', 'TEST', identifier, title, status,
        (TODAY - timedelta(days=5)).strftime('%Y-%m-%d'),
        (TODAY + timedelta(days=5)).strftime('%Y-%m-%d'),
        'Medium',
        reviewer_id, qcr_id,
        reviewer_due_date.strftime('%Y-%m-%d') if reviewer_due_date else None,
        qcr_due_date.strftime('%Y-%m-%d') if qcr_due_date else None,
        reviewer_token, qcr_token,
        reviewer_email_sent_at, reviewer_response_at,
        qcr_email_sent_at, qcr_response_at,
        1 if multi_reviewer_mode else 0,
        r"C:\Users\IANDAYA\Documents\Project Management -Simple\TrackerFiles\Turner\Submittals\Test-Reminder"
    ))
    
    item_id = cursor.lastrowid
    conn.commit()
    conn.close()
    
    print(f"Created test item: {identifier} (ID: {item_id}, Status: {status})")
    return item_id

def create_multi_reviewer_item(identifier, title, reviewers, qcr_email, qcr_name,
                               reviewer_due_date, qcr_due_date, status,
                               qcr_email_sent_at=None):
    """Create a multi-reviewer test item."""
    conn = get_db()
    cursor = conn.cursor()
    
    # Get QCR user ID
    cursor.execute('SELECT id FROM user WHERE email = ?', (qcr_email,))
    qcr_id = cursor.fetchone()
    qcr_id = qcr_id[0] if qcr_id else None
    
    qcr_token = generate_token()
    
    if qcr_email_sent_at is None and status == 'In QC':
        qcr_email_sent_at = datetime.now().isoformat()
    
    # Create item
    cursor.execute('''
        INSERT INTO item (
            type, bucket, identifier, title, status, 
            date_received, due_date, priority,
            qcr_id, initial_reviewer_due_date, qcr_due_date,
            email_token_qcr, qcr_email_sent_at,
            multi_reviewer_mode,
            folder_link
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        'Submittal', 'TEST', identifier, title, status,
        (TODAY - timedelta(days=5)).strftime('%Y-%m-%d'),
        (TODAY + timedelta(days=5)).strftime('%Y-%m-%d'),
        'Medium',
        qcr_id,
        reviewer_due_date.strftime('%Y-%m-%d') if reviewer_due_date else None,
        qcr_due_date.strftime('%Y-%m-%d') if qcr_due_date else None,
        qcr_token, qcr_email_sent_at,
        1,
        r"C:\Users\IANDAYA\Documents\Project Management -Simple\TrackerFiles\Turner\Submittals\Test-Reminder"
    ))
    
    item_id = cursor.lastrowid
    
    # Add reviewers
    for reviewer in reviewers:
        token = generate_token()
        cursor.execute('''
            INSERT INTO item_reviewers (
                item_id, reviewer_name, reviewer_email, email_token,
                email_sent_at, response_at, needs_response
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (
            item_id, reviewer['name'], reviewer['email'], token,
            datetime.now().isoformat(),
            reviewer.get('response_at'),
            1 if reviewer.get('needs_response', True) else 0
        ))
    
    conn.commit()
    conn.close()
    
    print(f"Created multi-reviewer test item: {identifier} (ID: {item_id}, Status: {status})")
    return item_id

def get_item_dict(item_id):
    """Get item as a dictionary with joined user info."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT i.*, 
               ir.email as reviewer_email, ir.display_name as reviewer_name,
               qcr.email as qcr_email, qcr.display_name as qcr_name
        FROM item i
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE i.id = ?
    ''', (item_id,))
    item = dict(cursor.fetchone())
    conn.close()
    return item

def get_reviewer_record(item_id, email):
    """Get reviewer record from item_reviewers table."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM item_reviewers WHERE item_id = ? AND reviewer_email = ?', (item_id, email))
    reviewer = cursor.fetchone()
    conn.close()
    return dict(reviewer) if reviewer else None

def test_single_reviewer_due_today():
    """Test Case 1: Single Reviewer - Reviewer Due Today"""
    print("\n" + "=" * 60)
    print("TEST CASE 1: Single Reviewer - Reviewer Due Today")
    print("=" * 60)
    
    item_id = create_test_item(
        identifier='TEST-REMINDER-SR-DT',
        title='Single Reviewer - Due Today Test',
        reviewer_email=EMAIL_1,
        reviewer_name='Isabel Andaya (HDR)',
        qcr_email=EMAIL_2,
        qcr_name='Isabel Andaya (Gmail)',
        reviewer_due_date=TODAY,
        qcr_due_date=TODAY + timedelta(days=2),
        status='Assigned'
    )
    
    item = get_item_dict(item_id)
    result = send_single_reviewer_reminder_email(item, 'reviewer', TODAY, 'due_today')
    print(f"Result: {result}")
    return result

def test_single_reviewer_overdue():
    """Test Case 2: Single Reviewer - Reviewer Overdue"""
    print("\n" + "=" * 60)
    print("TEST CASE 2: Single Reviewer - Reviewer Overdue")
    print("=" * 60)
    
    item_id = create_test_item(
        identifier='TEST-REMINDER-SR-OD',
        title='Single Reviewer - Overdue Test',
        reviewer_email=EMAIL_2,
        reviewer_name='Isabel Andaya (Gmail)',
        qcr_email=EMAIL_3,
        qcr_name='Isabel Andaya (SCU)',
        reviewer_due_date=YESTERDAY,
        qcr_due_date=TODAY + timedelta(days=1),
        status='In Review'
    )
    
    item = get_item_dict(item_id)
    result = send_single_reviewer_reminder_email(item, 'reviewer', YESTERDAY, 'overdue')
    print(f"Result: {result}")
    return result

def test_single_qcr_due_today():
    """Test Case 3: Single Reviewer - QCR Due Today"""
    print("\n" + "=" * 60)
    print("TEST CASE 3: Single Reviewer - QCR Due Today")
    print("=" * 60)
    
    item_id = create_test_item(
        identifier='TEST-REMINDER-SQ-DT',
        title='Single QCR - Due Today Test',
        reviewer_email=EMAIL_1,
        reviewer_name='Isabel Andaya (HDR)',
        qcr_email=EMAIL_3,
        qcr_name='Isabel Andaya (SCU)',
        reviewer_due_date=TODAY - timedelta(days=2),
        qcr_due_date=TODAY,
        status='In QC',
        reviewer_response_at=datetime.now().isoformat()  # Reviewer already responded
    )
    
    item = get_item_dict(item_id)
    result = send_single_reviewer_reminder_email(item, 'qcr', TODAY, 'due_today')
    print(f"Result: {result}")
    return result

def test_single_qcr_overdue():
    """Test Case 4: Single Reviewer - QCR Overdue"""
    print("\n" + "=" * 60)
    print("TEST CASE 4: Single Reviewer - QCR Overdue")
    print("=" * 60)
    
    item_id = create_test_item(
        identifier='TEST-REMINDER-SQ-OD',
        title='Single QCR - Overdue Test',
        reviewer_email=EMAIL_2,
        reviewer_name='Isabel Andaya (Gmail)',
        qcr_email=EMAIL_1,
        qcr_name='Isabel Andaya (HDR)',
        reviewer_due_date=TODAY - timedelta(days=3),
        qcr_due_date=YESTERDAY,
        status='In QC',
        reviewer_response_at=datetime.now().isoformat()  # Reviewer already responded
    )
    
    item = get_item_dict(item_id)
    result = send_single_reviewer_reminder_email(item, 'qcr', YESTERDAY, 'overdue')
    print(f"Result: {result}")
    return result

def test_multi_reviewer_due_today():
    """Test Case 5: Multi-Reviewer - Individual Reviewer Due Today"""
    print("\n" + "=" * 60)
    print("TEST CASE 5: Multi-Reviewer - Individual Reviewer Due Today")
    print("=" * 60)
    
    # Create with 2 reviewers - one has responded, one hasn't
    reviewers = [
        {'name': 'Isabel Andaya (HDR)', 'email': EMAIL_1, 'response_at': None, 'needs_response': True},
        {'name': 'Isabel Andaya (Gmail)', 'email': EMAIL_2, 'response_at': datetime.now().isoformat(), 'needs_response': False}  # Already responded
    ]
    
    item_id = create_multi_reviewer_item(
        identifier='TEST-REMINDER-MR-DT',
        title='Multi-Reviewer - Due Today Test',
        reviewers=reviewers,
        qcr_email=EMAIL_3,
        qcr_name='Isabel Andaya (SCU)',
        reviewer_due_date=TODAY,
        qcr_due_date=TODAY + timedelta(days=2),
        status='In Review'
    )
    
    item = get_item_dict(item_id)
    reviewer = get_reviewer_record(item_id, EMAIL_1)  # Only the one who hasn't responded
    
    result = send_multi_reviewer_reminder_email(item, reviewer, 'reviewer', TODAY, 'due_today')
    print(f"Result: {result}")
    return result

def test_multi_reviewer_overdue():
    """Test Case 6: Multi-Reviewer - Individual Reviewer Overdue"""
    print("\n" + "=" * 60)
    print("TEST CASE 6: Multi-Reviewer - Individual Reviewer Overdue")
    print("=" * 60)
    
    reviewers = [
        {'name': 'Isabel Andaya (Gmail)', 'email': EMAIL_2, 'response_at': None, 'needs_response': True},
        {'name': 'Isabel Andaya (SCU)', 'email': EMAIL_3, 'response_at': datetime.now().isoformat(), 'needs_response': False}
    ]
    
    item_id = create_multi_reviewer_item(
        identifier='TEST-REMINDER-MR-OD',
        title='Multi-Reviewer - Overdue Test',
        reviewers=reviewers,
        qcr_email=EMAIL_1,
        qcr_name='Isabel Andaya (HDR)',
        reviewer_due_date=YESTERDAY,
        qcr_due_date=TODAY + timedelta(days=1),
        status='In Review'
    )
    
    item = get_item_dict(item_id)
    reviewer = get_reviewer_record(item_id, EMAIL_2)
    
    result = send_multi_reviewer_reminder_email(item, reviewer, 'reviewer', YESTERDAY, 'overdue')
    print(f"Result: {result}")
    return result

def test_multi_qcr_due_today():
    """Test Case 7: Multi-Reviewer - QCR Due Today"""
    print("\n" + "=" * 60)
    print("TEST CASE 7: Multi-Reviewer - QCR Due Today")
    print("=" * 60)
    
    # All reviewers have responded
    reviewers = [
        {'name': 'Isabel Andaya (HDR)', 'email': EMAIL_1, 'response_at': datetime.now().isoformat(), 'needs_response': False},
        {'name': 'Isabel Andaya (Gmail)', 'email': EMAIL_2, 'response_at': datetime.now().isoformat(), 'needs_response': False}
    ]
    
    item_id = create_multi_reviewer_item(
        identifier='TEST-REMINDER-MQ-DT',
        title='Multi-Reviewer QCR - Due Today Test',
        reviewers=reviewers,
        qcr_email=EMAIL_3,
        qcr_name='Isabel Andaya (SCU)',
        reviewer_due_date=TODAY - timedelta(days=2),
        qcr_due_date=TODAY,
        status='In QC',
        qcr_email_sent_at=datetime.now().isoformat()
    )
    
    item = get_item_dict(item_id)
    result = send_multi_reviewer_qcr_reminder_email(item, TODAY, 'due_today')
    print(f"Result: {result}")
    return result

def test_multi_qcr_overdue():
    """Test Case 8: Multi-Reviewer - QCR Overdue"""
    print("\n" + "=" * 60)
    print("TEST CASE 8: Multi-Reviewer - QCR Overdue")
    print("=" * 60)
    
    # All reviewers have responded
    reviewers = [
        {'name': 'Isabel Andaya (Gmail)', 'email': EMAIL_2, 'response_at': datetime.now().isoformat(), 'needs_response': False},
        {'name': 'Isabel Andaya (SCU)', 'email': EMAIL_3, 'response_at': datetime.now().isoformat(), 'needs_response': False}
    ]
    
    item_id = create_multi_reviewer_item(
        identifier='TEST-REMINDER-MQ-OD',
        title='Multi-Reviewer QCR - Overdue Test',
        reviewers=reviewers,
        qcr_email=EMAIL_1,
        qcr_name='Isabel Andaya (HDR)',
        reviewer_due_date=TODAY - timedelta(days=3),
        qcr_due_date=YESTERDAY,
        status='In QC',
        qcr_email_sent_at=datetime.now().isoformat()
    )
    
    item = get_item_dict(item_id)
    result = send_multi_reviewer_qcr_reminder_email(item, YESTERDAY, 'overdue')
    print(f"Result: {result}")
    return result

def main():
    """Run all reminder tests."""
    print("=" * 60)
    print("REMINDER EMAIL WORKFLOW TEST")
    print("=" * 60)
    print(f"Today: {TODAY}")
    print(f"Yesterday: {YESTERDAY}")
    print(f"Test emails: {EMAIL_1}, {EMAIL_2}, {EMAIL_3}")
    print()
    
    # Initialize
    print("Initializing database...")
    init_db()
    
    # Cleanup previous test items
    cleanup_test_items()
    
    # Ensure test users exist
    user_ids = ensure_test_users()
    
    # Run all tests
    results = []
    
    input("\nPress Enter to run TEST 1: Single Reviewer - Reviewer Due Today...")
    results.append(('Single Reviewer - Reviewer Due Today', test_single_reviewer_due_today()))
    
    input("\nPress Enter to run TEST 2: Single Reviewer - Reviewer Overdue...")
    results.append(('Single Reviewer - Reviewer Overdue', test_single_reviewer_overdue()))
    
    input("\nPress Enter to run TEST 3: Single Reviewer - QCR Due Today...")
    results.append(('Single QCR - Due Today', test_single_qcr_due_today()))
    
    input("\nPress Enter to run TEST 4: Single Reviewer - QCR Overdue...")
    results.append(('Single QCR - Overdue', test_single_qcr_overdue()))
    
    input("\nPress Enter to run TEST 5: Multi-Reviewer - Individual Due Today...")
    results.append(('Multi-Reviewer - Due Today', test_multi_reviewer_due_today()))
    
    input("\nPress Enter to run TEST 6: Multi-Reviewer - Individual Overdue...")
    results.append(('Multi-Reviewer - Overdue', test_multi_reviewer_overdue()))
    
    input("\nPress Enter to run TEST 7: Multi-Reviewer - QCR Due Today...")
    results.append(('Multi-Reviewer QCR - Due Today', test_multi_qcr_due_today()))
    
    input("\nPress Enter to run TEST 8: Multi-Reviewer - QCR Overdue...")
    results.append(('Multi-Reviewer QCR - Overdue', test_multi_qcr_overdue()))
    
    # Summary
    print("\n" + "=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)
    for name, result in results:
        status = "✅ SUCCESS" if result.get('success') else "❌ FAILED"
        print(f"{status}: {name}")
        if not result.get('success'):
            print(f"   Error: {result.get('error')}")
    
    print("\n" + "=" * 60)
    print("All tests complete! Check your email inboxes.")
    print("=" * 60)

if __name__ == '__main__':
    main()
