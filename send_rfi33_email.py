"""Manually send QCR completion confirmation email for RFI 33"""
import sys
sys.path.insert(0, '.')

from app import send_qcr_completion_confirmation_email, get_db

item_id = 37

# Get item details
conn = get_db()
cursor = conn.cursor()
cursor.execute('SELECT * FROM item WHERE id = ?', (item_id,))
item = cursor.fetchone()
conn.close()

if item:
    print(f"Sending completion email for: {item['identifier']}")
    print(f"  QCR Action: {item['qcr_action']}")
    print(f"  Final Category: {item['final_response_category']}")
    
    result = send_qcr_completion_confirmation_email(
        item_id,
        item['qcr_action'],
        item['qcr_notes'] or '',
        final_category=item['final_response_category'],
        final_text=item['final_response_text']
    )
    
    print(f"\nResult: {result}")
else:
    print("Item not found")
