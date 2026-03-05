"""Regenerate response form for Submittal #13 34 19-3 (Item 75)
Reviewer: Iannelli, Michael - form was accidentally deleted.
"""
import sqlite3
from pathlib import Path

# First, fix the folder_link in the database (actual folder name differs)
CORRECT_FOLDER = r"\\sac-filsrv1\Projects\Structural-028\Projects\LEB\9.0_Const_Svcs\Turner\Submittals\Submittal - 13 34 19-3 - LEB01_133419-3 Shp Dwg_Anchor Bolts"

conn = sqlite3.connect('tracker.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

# Check current folder_link
c.execute('SELECT folder_link FROM item WHERE id=75')
current = c.fetchone()['folder_link']
print(f"Current folder_link: {current}")
print(f"Correct folder_link: {CORRECT_FOLDER}")

if current != CORRECT_FOLDER:
    c.execute('UPDATE item SET folder_link=? WHERE id=75', (CORRECT_FOLDER,))
    conn.commit()
    print("=> Updated folder_link in database")
else:
    print("=> folder_link already correct")

# Get the reviewer record for Iannelli, Michael
c.execute('SELECT * FROM item_reviewers WHERE item_id=75 AND reviewer_name LIKE ?', ('%Iannelli%',))
reviewer = c.fetchone()
if not reviewer:
    print("ERROR: Could not find Iannelli reviewer record!")
    conn.close()
    exit(1)

print(f"\nReviewer: {reviewer['reviewer_name']} ({reviewer['reviewer_email']})")
print(f"Token: {reviewer['email_token']}")
conn.close()

# Now generate the form
from app import generate_multi_reviewer_form
reviewer_dict = dict(reviewer)
result = generate_multi_reviewer_form(75, reviewer_dict)

if result['success']:
    print(f"\nSUCCESS! Form regenerated at:\n  {result['path']}")
    p = Path(result['path'])
    print(f"  File size: {p.stat().st_size:,} bytes")
else:
    print(f"\nFAILED: {result['error']}")
