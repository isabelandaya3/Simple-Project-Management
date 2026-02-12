"""
Add all missing closed RFIs to the Excel bulletin tracker.
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
from app import update_rfi_tracker_excel, get_db

def get_closed_rfis():
    """Get all closed RFIs from the database."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT * FROM item 
        WHERE type = 'RFI' AND status = 'Closed'
        ORDER BY id
    ''')
    rfis = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return rfis

def main():
    print("=" * 60)
    print("Adding Missing Closed RFIs to Excel Tracker")
    print("=" * 60)
    
    excel_path = CONFIG.get('rfi_tracker_excel_path')
    print(f"\nExcel path: {excel_path}")
    
    if not excel_path:
        print("ERROR: RFI tracker Excel path not configured in config.json")
        return
    
    if not Path(excel_path).exists():
        print(f"ERROR: Excel file not found: {excel_path}")
        return
    
    # Get all closed RFIs
    closed_rfis = get_closed_rfis()
    print(f"\nFound {len(closed_rfis)} closed RFIs in database")
    
    if not closed_rfis:
        print("No closed RFIs to add.")
        return
    
    # Process each RFI
    added = 0
    updated = 0
    errors = 0
    
    for rfi in closed_rfis:
        identifier = rfi.get('identifier', '')
        print(f"\nProcessing {identifier}...")
        
        result = update_rfi_tracker_excel(rfi, action='close')
        
        if result.get('success'):
            msg = result.get('message', '')
            if 'added' in msg.lower():
                added += 1
                print(f"  ✓ Added: {msg}")
            elif 'updated' in msg.lower():
                updated += 1
                print(f"  ✓ Updated: {msg}")
            else:
                print(f"  ✓ {msg}")
        else:
            errors += 1
            print(f"  ✗ Error: {result.get('error', 'Unknown error')}")
    
    print("\n" + "=" * 60)
    print(f"Summary:")
    print(f"  Added:   {added}")
    print(f"  Updated: {updated}")
    print(f"  Errors:  {errors}")
    print("=" * 60)

if __name__ == '__main__':
    main()
