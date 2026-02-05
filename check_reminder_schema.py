"""Check reminder_log table schema."""
import sqlite3

conn = sqlite3.connect('tracker.db')
cursor = conn.cursor()

cursor.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='reminder_log'")
result = cursor.fetchone()
if result:
    print("Table schema:")
    print(result[0])

# Check for unique constraints
cursor.execute("SELECT * FROM pragma_index_list('reminder_log')")
indexes = cursor.fetchall()
print("\nIndexes:")
for idx in indexes:
    print(idx)

conn.close()
