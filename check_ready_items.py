import sqlite3

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

cursor.execute("""
    SELECT id, identifier, type, status, 
           response_category, response_text,
           final_response_category, final_response_text, 
           qcr_action, qcr_response_category, reviewer_email_sent_at, multi_reviewer_mode, qcr_response_at
    FROM item 
    WHERE status = 'Ready for Response'
""")

items = cursor.fetchall()

print('\n=== Items with Ready for Response status ===')
for item in items:
    print(f"\nID: {item['id']}")
    print(f"  Type/ID: {item['type']} {item['identifier']}")
    print(f"  Status: {item['status']}")
    print(f"  response_category: {item['response_category']}")
    print(f"  response_text: {item['response_text'][:50] if item['response_text'] else 'None'}")
    print(f"  final_response_category: {item['final_response_category']}")
    print(f"  final_response_text: {item['final_response_text'][:50] if item['final_response_text'] else 'None'}")
    print(f"  QCR Action: {item['qcr_action']}")
    
    # Check item_reviewers table
    cursor.execute("SELECT * FROM item_reviewers WHERE item_id = ?", (item['id'],))
    reviewers = cursor.fetchall()
    if reviewers:
        print(f"  Item Reviewers in table: {len(reviewers)}")

conn.close()
