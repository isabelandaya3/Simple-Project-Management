import sqlite3

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
cur = conn.cursor()

# Get all closed items
cur.execute("SELECT id, identifier, type, status, closed_at, final_response_category FROM item WHERE status = 'Closed' ORDER BY closed_at DESC")
print("=== CLOSED ITEMS ===")
for r in cur.fetchall():
    print(f"  {r['identifier']:30s}  type={r['type']:10s}  closed={r['closed_at'][:10] if r['closed_at'] else 'None':12s}  cat={r['final_response_category']}")

# Get all items with status Ready for Response (completed but not closed)
cur.execute("SELECT id, identifier, type, status, qcr_response_status, final_response_category FROM item WHERE status = 'Ready for Response'")
print("\n=== READY FOR RESPONSE (Not yet closed) ===")
for r in cur.fetchall():
    print(f"  {r['identifier']:30s}  type={r['type']:10s}  qcr={r['qcr_response_status']:15s}  cat={r['final_response_category']}")

conn.close()


