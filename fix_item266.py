#!/usr/bin/env python3
"""Fix item 266 - send QCR email and update status."""
import sys
sys.path.insert(0, 'c:/Users/IANDAYA/Documents/Project Management -Simple')

from app import send_qcr_assignment_email, get_db

item_id = 39  # Submittal #266

# Check current status
conn = get_db()
cursor = conn.cursor()
cursor.execute("SELECT id, identifier, status, qcr_email_sent_at, qcr_id FROM item WHERE id = ?", (item_id,))
item = cursor.fetchone()
print(f"Before fix:")
print(f"  Item: {item['identifier']}")
print(f"  Status: {item['status']}")
print(f"  QCR Email Sent: {item['qcr_email_sent_at']}")
print(f"  QCR ID: {item['qcr_id']}")
conn.close()

# Send QCR email
print("\nSending QCR email...")
result = send_qcr_assignment_email(item_id)
print(f"Result: {result}")

# Check status after
conn = get_db()
cursor = conn.cursor()
cursor.execute("SELECT id, identifier, status, qcr_email_sent_at FROM item WHERE id = ?", (item_id,))
item = cursor.fetchone()
print(f"\nAfter fix:")
print(f"  Status: {item['status']}")
print(f"  QCR Email Sent: {item['qcr_email_sent_at']}")
conn.close()
