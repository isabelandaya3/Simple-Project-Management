"""
Fix items where reviewer_email_sent_at is NULL but item_reviewers have email_sent_at.
This should only be run once to fix existing data.
"""
import sqlite3
from datetime import datetime

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Find items with NULL reviewer_email_sent_at but with item_reviewers that have email_sent_at
cursor.execute("""
    SELECT DISTINCT i.id, i.identifier, i.type,
           MIN(ir.email_sent_at) as first_email_sent
    FROM item i
    JOIN item_reviewers ir ON ir.item_id = i.id
    WHERE i.reviewer_email_sent_at IS NULL
    AND ir.email_sent_at IS NOT NULL
    GROUP BY i.id
""")

items_to_fix = cursor.fetchall()

if items_to_fix:
    print(f"\n=== Found {len(items_to_fix)} items to fix ===")
    for item in items_to_fix:
        print(f"\nFixing ID: {item['id']} ({item['type']} {item['identifier']})")
        print(f"  Setting reviewer_email_sent_at = {item['first_email_sent']}")
        
        cursor.execute("""
            UPDATE item 
            SET reviewer_email_sent_at = ?
            WHERE id = ?
        """, (item['first_email_sent'], item['id']))
    
    conn.commit()
    print(f"\n✅ Fixed {len(items_to_fix)} items")
else:
    print("\n✅ No items need fixing")

conn.close()
