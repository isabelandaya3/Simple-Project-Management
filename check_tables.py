import sqlite3

conn = sqlite3.connect('tracker.db')
c = conn.cursor()
c.execute("SELECT name FROM sqlite_master WHERE type='table'")
tables = [r[0] for r in c.fetchall()]
print('All tables:', tables)

# Check if project table exists
if 'project' in tables:
    print('\nProject table exists!')
    c.execute("PRAGMA table_info(project)")
    print('Columns:', [r[1] for r in c.fetchall()])
    c.execute("SELECT * FROM project")
    projects = c.fetchall()
    print(f'Projects: {len(projects)}')
    for p in projects:
        print(f'  {p}')
else:
    print('\nProject table does NOT exist!')

conn.close()
