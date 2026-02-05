"""
Add the missed Submittal #03 30 00-3 to the tracker database.
"""

import sqlite3
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent.absolute()
DATABASE_PATH = BASE_DIR / "tracker.db"

def add_missed_submittal():
    """Add the missed submittal from the ACC email."""
    
    # Submittal details from the email
    identifier = "Submittal #03 30 00-3"
    title = "LEB10_033000_Prdt_Data_IMI_Mix_Designs_Early_Works"
    bucket = "ACC_TURNER"
    item_type = "Submittal"
    due_date = "2026-02-09"
    priority = "Medium"
    date_received = "2026-02-03"  # Today's date when email was received
    
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # Check if already exists
    cursor.execute('''
        SELECT id FROM item WHERE identifier = ? AND bucket = ?
    ''', (identifier, bucket))
    existing = cursor.fetchone()
    
    if existing:
        print(f"Item already exists with ID: {existing['id']}")
        conn.close()
        return existing['id']
    
    # Create folder for the item
    import json
    config_path = BASE_DIR / "config.json"
    with open(config_path, 'r') as f:
        config = json.load(f)
    
    base_path = Path(config['base_folder_path'])
    folder_path = base_path / "Turner" / "Submittals" / f"Submittal - 03 30 00-3"
    
    try:
        folder_path.mkdir(parents=True, exist_ok=True)
        folder_link = str(folder_path)
        print(f"Created folder: {folder_link}")
    except Exception as e:
        print(f"Could not create folder: {e}")
        folder_link = None
    
    # Calculate review due dates (simple calculation)
    # Initial reviewer due: 2 days before contractor due date
    # QCR due: same as contractor due date
    initial_reviewer_due = "2026-02-07"  # 2 days before Feb 9
    qcr_due = "2026-02-09"
    
    now = datetime.now().isoformat()
    
    # Insert the item
    cursor.execute('''
        INSERT INTO item (
            type, bucket, identifier, title, 
            created_at, last_email_at, due_date, priority, 
            folder_link, date_received, 
            initial_reviewer_due_date, qcr_due_date,
            status
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        item_type, bucket, identifier, title,
        now, now, due_date, priority,
        folder_link, date_received,
        initial_reviewer_due, qcr_due,
        'Open'
    ))
    
    item_id = cursor.lastrowid
    conn.commit()
    conn.close()
    
    print(f"Successfully added {identifier} with ID: {item_id}")
    print(f"  Title: {title}")
    print(f"  Due Date: {due_date}")
    print(f"  Folder: {folder_link}")
    
    return item_id

if __name__ == "__main__":
    add_missed_submittal()
