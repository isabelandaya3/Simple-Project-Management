import sqlite3
import sys
sys.path.insert(0, '.')
from app import generate_multi_reviewer_form

# Get item 37 details
conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

cursor.execute('SELECT * FROM item WHERE id = 37')
item = cursor.fetchone()

print(f"Item found: {item['identifier']} - {item['title']}")
print(f"Type: {item['type']}")
print(f"Folder: {item['folder_link']}")

# Get the reviewer for this item
cursor.execute('SELECT * FROM item_reviewers WHERE item_id = 37')
reviewer = cursor.fetchone()

if reviewer:
    print(f"\nReviewer: {reviewer['reviewer_name']} ({reviewer['reviewer_email']})")
    
    # Generate the form
    print("\nGenerating RFI response form...")
    result = generate_multi_reviewer_form(37, dict(reviewer))
    
    if result['success']:
        print(f"✅ Form generated successfully!")
        print(f"   Path: {result['path']}")
    else:
        print(f"❌ Error: {result['error']}")
else:
    print("\n❌ No reviewer found for this item")

conn.close()
