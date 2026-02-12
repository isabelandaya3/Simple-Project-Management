"""Create folder for RFI #101"""
import sqlite3
from pathlib import Path
import json
import re

# Load config
with open('config.json') as f:
    CONFIG = json.load(f)

def sanitize_folder_name(name):
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '-')
    name = re.sub(r'-+', '-', name)
    name = name.strip('- ')
    return name

def create_item_folder(item_type, identifier, bucket, title=None, base_folder=None):
    base_path = Path(base_folder) if base_folder else Path(CONFIG['base_folder_path'])
    bucket_folder = bucket.replace('ACC_', '').title()
    if bucket == 'ALL':
        bucket_folder = 'General'
    type_folder = 'Submittals' if item_type == 'Submittal' else 'RFIs'
    clean_id = sanitize_folder_name(identifier)
    folder_id = clean_id.replace(f'{item_type} #', '')
    if title:
        clean_title = sanitize_folder_name(title)[:100]
        item_folder = f'{item_type} - {folder_id} - {clean_title}'
    else:
        item_folder = f'{item_type} - {folder_id}'
    full_path = base_path / bucket_folder / type_folder / item_folder
    print(f"Target folder: {full_path}")
    try:
        full_path.mkdir(parents=True, exist_ok=True)
        return str(full_path)
    except Exception as e:
        print(f'Error creating folder {full_path}: {e}')
        return None

# Get RFI 101 details
conn = sqlite3.connect('tracker.db')
c = conn.cursor()
c.execute('SELECT id, type, identifier, bucket, title FROM item WHERE id = 70')
row = c.fetchone()
print(f'Item: {row}')

# Create folder
folder_path = create_item_folder(row[1], row[2], row[3], row[4])
print(f'Created folder: {folder_path}')

# Update database
if folder_path:
    c.execute('UPDATE item SET folder_link = ? WHERE id = ?', (folder_path, row[0]))
    conn.commit()
    print('Database updated with folder path')
    
    # Verify
    c.execute('SELECT folder_link FROM item WHERE id = 70')
    result = c.fetchone()
    print(f'Verified folder_link: {result[0]}')

conn.close()
