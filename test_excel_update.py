"""Backfill all closed items to Excel trackers (RFI and Submittal)."""
from app import get_db, update_submittal_tracker_excel, update_rfi_tracker_excel

conn = get_db()
c = conn.cursor()

# Process all closed RFIs
print("=" * 50)
print("Processing closed RFIs...")
print("=" * 50)
c.execute("SELECT * FROM item WHERE type = 'RFI' AND status = 'Closed'")
rfis = c.fetchall()
rfi_success = 0
rfi_fail = 0
for rfi in rfis:
    item = dict(rfi)
    result = update_rfi_tracker_excel(item, action='close')
    if result.get('success'):
        print(f"  ✓ {item.get('identifier')}: {result.get('message')}")
        rfi_success += 1
    else:
        print(f"  ✗ {item.get('identifier')}: {result.get('error')}")
        rfi_fail += 1

print(f"\nRFI Results: {rfi_success} success, {rfi_fail} failed")

# Process all closed Submittals
print("\n" + "=" * 50)
print("Processing closed Submittals...")
print("=" * 50)
c.execute("SELECT * FROM item WHERE type = 'Submittal' AND status = 'Closed'")
submittals = c.fetchall()
sub_success = 0
sub_fail = 0
for sub in submittals:
    item = dict(sub)
    # Get reviewer info
    c.execute("SELECT reviewer_name, reviewer_email, response_at, response_category, internal_notes FROM item_reviewers WHERE item_id = ?", (item['id'],))
    reviewers = [dict(row) for row in c.fetchall()]
    
    result = update_submittal_tracker_excel(item, reviewers=reviewers, action='close')
    if result.get('success'):
        print(f"  ✓ {item.get('identifier')}: {result.get('message')}")
        sub_success += 1
    else:
        print(f"  ✗ {item.get('identifier')}: {result.get('error')}")
        sub_fail += 1

print(f"\nSubmittal Results: {sub_success} success, {sub_fail} failed")

conn.close()

print("\n" + "=" * 50)
print(f"TOTAL: {rfi_success + sub_success} items updated, {rfi_fail + sub_fail} failed")
print("=" * 50)
