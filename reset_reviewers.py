import sqlite3
conn = sqlite3.connect('c:/Users/IANDAYA/Documents/Project Management -Simple/tracker.db')
cursor = conn.cursor()
cursor.execute('UPDATE item_reviewers SET response_at = NULL, response_category = NULL, internal_notes = NULL WHERE item_id = 17')
cursor.execute("UPDATE item SET status = 'Assigned', reviewer_response_status = 'Emails Sent', qcr_email_sent_at = NULL, qcr_response_status = NULL WHERE id = 17")
conn.commit()
print('Reset complete')
cursor.execute('SELECT reviewer_name, response_at, response_category FROM item_reviewers WHERE item_id = 17')
for row in cursor.fetchall():
    print(row)
conn.close()
