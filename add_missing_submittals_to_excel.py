"""
Add all missing closed Submittals to the Excel tracker.
"""
import sqlite3
import sys
import json
from pathlib import Path

# Load config
config_path = Path(__file__).parent / 'config.json'
with open(config_path) as f:
    CONFIG = json.load(f)

# Import the update function from app
from app import update_submittal_tracker_excel, get_db

def get_closed_submittals():
    """Get all closed Submittals from the database with reviewer info."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT * FROM item 
        WHERE type = 'Submittal' AND status = 'Closed'
        ORDER BY id
    ''')
    submittals = [dict(row) for row in cursor.fetchall()]
    
    # Get reviewer info for each submittal
    for sub in submittals:
        cursor.execute('''
            SELECT reviewer_name, reviewer_email, response_at, response_category, internal_notes
            FROM item_reviewers WHERE item_id = ?
        ''', (sub['id'],))
        sub['reviewers'] = [dict(row) for row in cursor.fetchall()]
    
    conn.close()
    return submittals

def main():
    print("=" * 60)
    print("Adding Missing Closed Submittals to Excel Tracker")
    print("=" * 60)
    
    excel_path = CONFIG.get('submittal_tracker_excel_path')
    print(f"\nExcel path: {excel_path}")
    
    if not excel_path:
        print("ERROR: Submittal tracker Excel path not configured in config.json")
        return
    
    # Note: The update function will create the file if it doesn't exist
    if not Path(excel_path).exists():
        print(f"Note: Excel file does not exist yet, will be created: {excel_path}")
    
    # Get all closed Submittals
    closed_submittals = get_closed_submittals()
    print(f"\nFound {len(closed_submittals)} closed Submittals in database")
    
    if not closed_submittals:
        print("No closed Submittals to add.")
        return
    
    # Process each Submittal
    added = 0
    updated = 0
    errors = 0
    skipped = 0
    
    for sub in closed_submittals:
        identifier = sub.get('identifier', '')
        print(f"\nProcessing {identifier}...")
        
        result = update_submittal_tracker_excel(sub, reviewers=sub.get('reviewers'), action='close')
        
        if result.get('success'):
            msg = result.get('message', '')
            if 'added' in msg.lower():
                added += 1
                print(f"  ✓ Added: {msg}")
            elif 'updated' in msg.lower():
                updated += 1
                print(f"  ✓ Updated: {msg}")
            elif 'skipping' in msg.lower() or 'not a' in msg.lower():
                skipped += 1
                print(f"  - Skipped: {msg}")
            else:
                print(f"  ✓ {msg}")
        else:
            errors += 1
            print(f"  ✗ Error: {result.get('error', 'Unknown error')}")
    
    print("\n" + "=" * 60)
    print(f"Summary:")
    print(f"  Added:   {added}")
    print(f"  Updated: {updated}")
    print(f"  Skipped: {skipped}")
    print(f"  Errors:  {errors}")
    print("=" * 60)

if __name__ == '__main__':
    main()
