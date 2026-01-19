import sqlite3
conn = sqlite3.connect('tracker.db')
c = conn.cursor()
c.execute("UPDATE item SET status = 'In QC', reviewer_response_status = 'All Responded' WHERE id = 17")
conn.commit()
print('Fixed item 17 status to In QC')
conn.close()
