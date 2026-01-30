import sqlite3
from app import send_qcr_assignment_email

# Connect to database
conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Get item 35 details
cursor.execute('SELECT * FROM item WHERE id = 35')
item = cursor.fetchone()

if item:
    print("=== Item 35 Details ===")
    print(f"ID: {item['id']}")
    print(f"Identifier: {item['identifier']}")
    print(f"Status: {item['status']}")
    print()
    print("All columns:")
    for key in item.keys():
        print(f"  {key}: {item[key]}")
    print()
    
    # Send QCR email
    print("Sending QCR email...")
    try:
        result = send_qcr_assignment_email(35)
        if result['success']:
            print(f"✓ QCR email sent successfully!")
            print(f"  Message: {result.get('message', 'N/A')}")
        else:
            print(f"✗ Failed to send QCR email")
            print(f"  Error: {result.get('error', 'Unknown error')}")
    except Exception as e:
        print(f"✗ Exception occurred: {e}")
else:
    print("Item 35 not found")

conn.close()
