"""
Migration script to add multi-project support tables.
Run this once to add the project tables to an existing database.
"""
import sqlite3
import json
from pathlib import Path

# Load config
config_path = Path(__file__).parent / 'config.json'
with open(config_path) as f:
    CONFIG = json.load(f)

db_path = Path(__file__).parent / 'tracker.db'
conn = sqlite3.connect(str(db_path))
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

print("Starting migration for multi-project support...")

# 1. Create project table if not exists
print("\n1. Creating project table...")
cursor.execute('''
    CREATE TABLE IF NOT EXISTS project (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        short_name TEXT NOT NULL UNIQUE,
        description TEXT,
        base_folder_path TEXT,
        rfi_tracker_excel_path TEXT,
        submittal_tracker_excel_path TEXT,
        outlook_folder TEXT,
        user_names TEXT,
        bucket_patterns TEXT,
        settings TEXT,
        created_by INTEGER,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        is_active INTEGER DEFAULT 1
    )
''')
print("   project table created/verified")

# 2. Create project_user table if not exists
print("\n2. Creating project_user table...")
cursor.execute('''
    CREATE TABLE IF NOT EXISTS project_user (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER NOT NULL,
        user_id INTEGER NOT NULL,
        role TEXT DEFAULT 'member',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (project_id) REFERENCES project(id) ON DELETE CASCADE,
        FOREIGN KEY (user_id) REFERENCES user(id) ON DELETE CASCADE,
        UNIQUE(project_id, user_id)
    )
''')
print("   project_user table created/verified")

# 3. Add current_project_id to user table if not exists
print("\n3. Adding current_project_id to user table...")
try:
    cursor.execute("ALTER TABLE user ADD COLUMN current_project_id INTEGER")
    print("   current_project_id column added")
except sqlite3.OperationalError as e:
    if "duplicate column" in str(e).lower():
        print("   current_project_id column already exists")
    else:
        print(f"   Warning: {e}")

# 4. Add project_id to item table if not exists
print("\n4. Adding project_id to item table...")
try:
    cursor.execute("ALTER TABLE item ADD COLUMN project_id INTEGER")
    print("   project_id column added")
except sqlite3.OperationalError as e:
    if "duplicate column" in str(e).lower():
        print("   project_id column already exists")
    else:
        print(f"   Warning: {e}")

# 5. Create default LEB project from config
print("\n5. Creating default LEB project from config...")
cursor.execute("SELECT id FROM project WHERE short_name = 'LEB'")
existing = cursor.fetchone()

if not existing:
    user_names = CONFIG.get('user_names', [])
    user_names_json = json.dumps(user_names) if user_names else '[]'
    
    settings = {
        'outlook_folder': CONFIG.get('outlook_folder', 'Inbox'),
        'rfi_tracker_excel_path': CONFIG.get('rfi_tracker_excel_path', ''),
        'submittal_tracker_excel_path': CONFIG.get('submittal_tracker_excel_path', '')
    }
    
    cursor.execute('''
        INSERT INTO project (name, short_name, description, base_folder_path, 
                            rfi_tracker_excel_path, submittal_tracker_excel_path,
                            outlook_folder, user_names, settings, is_active)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
    ''', (
        'LEB Project',
        'LEB',
        'Lawrence Berkeley Lab Extension Building',
        CONFIG.get('base_folder_path', ''),
        CONFIG.get('rfi_tracker_excel_path', ''),
        CONFIG.get('submittal_tracker_excel_path', ''),
        CONFIG.get('outlook_folder', 'Inbox'),
        user_names_json,
        json.dumps(settings)
    ))
    project_id = cursor.lastrowid
    print(f"   Created LEB project with ID {project_id}")
    
    # 6. Update all existing items to belong to LEB project
    print("\n6. Updating existing items to LEB project...")
    cursor.execute("UPDATE item SET project_id = ? WHERE project_id IS NULL", (project_id,))
    print(f"   Updated {cursor.rowcount} items")
    
    # 7. Add all existing users to LEB project
    print("\n7. Adding existing users to LEB project...")
    cursor.execute("SELECT id FROM user")
    users = cursor.fetchall()
    for user in users:
        try:
            cursor.execute('''
                INSERT OR IGNORE INTO project_user (project_id, user_id, role)
                VALUES (?, ?, 'member')
            ''', (project_id, user['id']))
        except:
            pass
    print(f"   Added {len(users)} users to LEB project")
    
    # 8. Set current_project_id for all users to LEB
    print("\n8. Setting current project for all users...")
    cursor.execute("UPDATE user SET current_project_id = ? WHERE current_project_id IS NULL", (project_id,))
    print(f"   Updated {cursor.rowcount} users")
else:
    print(f"   LEB project already exists with ID {existing['id']}")

conn.commit()
conn.close()

print("\n" + "="*50)
print("Migration completed successfully!")
print("="*50)
