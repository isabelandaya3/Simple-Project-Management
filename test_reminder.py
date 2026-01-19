#!/usr/bin/env python3
"""Test script to verify the reminder email system and due date calculations."""

import sys
import os
from datetime import datetime, timedelta

# Add the app directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import (
    calculate_review_due_dates, 
    subtract_business_days,
    business_days_between,
    QCR_REVIEW_DAYS,
    QCR_DAYS_BEFORE_DUE,
    get_items_needing_reminders,
    is_past_reminder_time_today,
    get_pst_now
)

def test_due_date_calculation():
    """Test that QCR gets 2 business days after reviewer due date."""
    print("=" * 60)
    print("Testing Due Date Calculation")
    print("=" * 60)
    print(f"QCR_REVIEW_DAYS = {QCR_REVIEW_DAYS}")
    print(f"QCR_DAYS_BEFORE_DUE = {QCR_DAYS_BEFORE_DUE}")
    print()
    
    # Test with a specific date
    date_received = datetime(2026, 1, 19).date()  # Monday
    contractor_due_date = datetime(2026, 1, 30).date()  # Friday
    
    result = calculate_review_due_dates(
        date_received.strftime('%Y-%m-%d'),
        contractor_due_date.strftime('%Y-%m-%d'),
        'Medium'
    )
    
    print(f"Date Received: {date_received} ({date_received.strftime('%A')})")
    print(f"Contractor Due: {contractor_due_date} ({contractor_due_date.strftime('%A')})")
    print()
    print(f"Reviewer Due: {result['initial_reviewer_due_date']}")
    print(f"QCR Due: {result['qcr_due_date']}")
    
    # Calculate gap between reviewer and QCR due dates
    reviewer_due = datetime.strptime(result['initial_reviewer_due_date'], '%Y-%m-%d').date()
    qcr_due = datetime.strptime(result['qcr_due_date'], '%Y-%m-%d').date()
    
    gap = business_days_between(reviewer_due, qcr_due)
    print(f"\nBusiness days between Reviewer Due and QCR Due: {gap}")
    print(f"Expected: {QCR_REVIEW_DAYS}")
    
    if gap == QCR_REVIEW_DAYS:
        print("✅ Gap is correct!")
    else:
        print(f"❌ Gap should be {QCR_REVIEW_DAYS}, but got {gap}")
    
    print()

def test_reminder_timing():
    """Test the PST time check."""
    print("=" * 60)
    print("Testing Reminder Timing")
    print("=" * 60)
    
    pst_now = get_pst_now()
    print(f"Current PST time: {pst_now.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Is past 8 AM PST: {is_past_reminder_time_today()}")
    print()

def test_items_needing_reminders():
    """Check which items need reminders."""
    print("=" * 60)
    print("Testing Items Needing Reminders")
    print("=" * 60)
    
    items = get_items_needing_reminders()
    
    print(f"Single reviewer items: {len(items['single_reviewer'])}")
    for item, role, due_date, stage in items['single_reviewer']:
        print(f"  - {item['identifier']}: {role} due {due_date} ({stage})")
    
    print(f"\nMulti-reviewer items: {len(items['multi_reviewer'])}")
    for item, reviewer, role, due_date, stage in items['multi_reviewer']:
        print(f"  - {item['identifier']}: {reviewer['reviewer_name']} due {due_date} ({stage})")
    
    print(f"\nMulti-reviewer QCR items: {len(items['multi_reviewer_qcr'])}")
    for item, due_date, stage in items['multi_reviewer_qcr']:
        print(f"  - {item['identifier']}: QCR due {due_date} ({stage})")
    
    print()

if __name__ == '__main__':
    test_due_date_calculation()
    test_reminder_timing()
    test_items_needing_reminders()
