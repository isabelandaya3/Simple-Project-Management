import sqlite3

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

cursor.execute("SELECT id, type, identifier, response_category, final_response_category FROM item WHERE identifier LIKE '%31%' AND type LIKE '%RFI%'")
rows = cursor.fetchall()

for r in rows:
    print(f"ID: {r['id']}")
    print(f"  Type: '{r['type']}'")
    print(f"  Identifier: {r['identifier']}")
    print(f"  response_category: {r['response_category']}")
    print(f"  final_response_category: {r['final_response_category']}")
    print()

conn.close()
