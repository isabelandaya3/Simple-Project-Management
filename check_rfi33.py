import sqlite3

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

c.execute("""
    SELECT id, identifier, status, qcr_action, final_response_category, final_response_text 
    FROM item 
    WHERE identifier LIKE '%33%'
""")

for r in c.fetchall():
    print(dict(r))

conn.close()
