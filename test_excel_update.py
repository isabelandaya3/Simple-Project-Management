"""Test Submittal Excel update function."""
from app import get_db, update_submittal_tracker_excel

conn = get_db()
c = conn.cursor()
c.execute("SELECT * FROM item WHERE type = 'Submittal' AND status = 'Closed' LIMIT 1")
item = dict(c.fetchone())

# Get reviewer info
c.execute("SELECT reviewer_name, reviewer_email, response_at, response_category, internal_notes FROM item_reviewers WHERE item_id = ?", (item['id'],))
reviewers = [dict(row) for row in c.fetchall()]

conn.close()

print(f"Testing with: {item.get('identifier')}")
print(f"Reviewers: {[r.get('reviewer_name') for r in reviewers]}")
result = update_submittal_tracker_excel(item, reviewers=reviewers, action='close')
print(f"Result: {result}")
