import sqlite3

# Connect to database
conn = sqlite3.connect('tracker.db')
cursor = conn.cursor()

# Update RFI #33 with the question
question = """It is contractor preference to hold concrete down 6" below bottom of anchor bolts for interior piers and bottom of grade beam for exterior perimeter piers with the intent to pour back to final elevation upon completion of backfilling operations to prevent damage to anchor bolts. Please advise if this acceptable for all piers in locations highlighted below."""

cursor.execute('UPDATE item SET rfi_question = ? WHERE id = 37', (question,))
conn.commit()

# Verify the update
cursor.execute('SELECT id, identifier, rfi_question FROM item WHERE id = 37')
result = cursor.fetchone()
print(f"✅ Updated item {result[0]} ({result[1]})")
print(f"Question: {result[2][:100]}...")

conn.close()
