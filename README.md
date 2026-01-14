# LEB RFI/Submittal Tracker

A local-only Windows application for tracking RFIs and Submittals from Autodesk Construction Cloud (ACC) email notifications. Features an Apple-style web dashboard for managing items, assignments, and comments.

![Dashboard Preview](docs/dashboard-preview.png)

## Features

- 📧 **Automatic Email Monitoring**: Polls Outlook for ACC notifications and extracts RFI/Submittal information
- 🎨 **Beautiful Dashboard**: Apple-style web interface with clean design
- 👥 **Multi-User Support**: Simple user management with admin and regular user roles
- 📁 **Folder Management**: Automatic creation of organized folder structure for files
- 💬 **Comments & Notes**: Track discussions and notes on each item
- 🏷️ **Status Tracking**: Track items through workflow states (Unassigned → Closed)
- 🔗 **ACC Placeholder**: UI ready for future direct ACC API integration

## Requirements

- **Windows 10/11** with Outlook desktop (M365) installed and configured
- **Python 3.8+** (must be in PATH or use virtual environment)
- **No admin rights required** - runs entirely in user space

## Quick Start

### 1. Install Python Dependencies

Open a terminal in this folder and run:

```bash
# Create a virtual environment (recommended)
python -m venv venv

# Activate it
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Start the Tracker

Double-click `start_all.bat` or run:

```bash
python app.py
```

### 3. Access the Dashboard

Open your browser to: **http://localhost:5000**

Default login:
- **Email**: `admin@local`
- **Password**: `admin123`

> ⚠️ **Important**: Change the default admin password after first login!

## Batch Files

| File | Description |
|------|-------------|
| `start_tracker.bat` | Starts only the Python server |
| `open_dashboard.bat` | Opens the dashboard in your browser |
| `start_all.bat` | Starts server (if not running) and opens dashboard |

## Project Structure

```
LEB Tracker/
├── app.py                 # Main Python backend (Flask + email poller)
├── tracker.db             # SQLite database (created on first run)
├── config.json            # Configuration file (created on first run)
├── requirements.txt       # Python dependencies
├── setup_startup.py       # Windows startup helper
├── start_tracker.bat      # Start server
├── open_dashboard.bat     # Open browser
├── start_all.bat          # Start everything
├── static/
│   ├── index.html         # Dashboard HTML
│   ├── style.css          # Apple-style CSS
│   └── app.js             # Frontend JavaScript
└── TrackerFiles/          # Default folder for item files
    ├── Turner/
    │   ├── RFIs/
    │   └── Submittals/
    ├── Mortenson/
    │   ├── RFIs/
    │   └── Submittals/
    └── ...
```

## Email Polling

The tracker automatically monitors your Outlook inbox for ACC notifications:

### How It Works

1. Every 5 minutes (configurable), the tracker connects to Outlook
2. Scans for emails containing "LEB" and "RFI" or "Submittal"
3. Parses the email subject to extract:
   - Type (RFI or Submittal)
   - Identifier (e.g., "Submittal #13 34 19-2")
   - Bucket (Turner, Mortenson, FTI, or General)
4. Parses the email body for:
   - Due date
   - Priority
5. Creates or updates items in the database
6. Creates a folder for new items

### Supported Email Patterns

Subject patterns recognized:
- `Action Required: LEB - Turner (NB.TypeF2.0) - Submittal #13 34 19-2 was assigned to you`
- `LEB - Mortenson (NB.TypeF2.0) - RFI #123 needs your review`
- `LEB (NB.TypeF2.0) - Submittal #456 requires action`

Body patterns for due date:
- `Due date: March 15, 2026`
- `Due Date: 03/15/2026`
- `Due: 2026-03-15`

Body patterns for priority:
- `Priority: High`
- `Priority: Medium`
- `Priority: Low`

### Custom Outlook Folder

By default, the tracker scans your Inbox. To use a specific folder:

1. Create a folder in Outlook (e.g., "LEB ACC")
2. Set up an Outlook rule to move ACC emails to that folder
3. Update `config.json`:

```json
{
    "outlook_folder": "LEB ACC"
}
```

## Configuration

The `config.json` file (created on first run) contains:

```json
{
    "base_folder_path": "C:\\Users\\YOU\\Documents\\Project Management -Simple\\TrackerFiles",
    "outlook_folder": "Inbox",
    "poll_interval_minutes": 5,
    "server_port": 5000,
    "project_name": "LEB – Local Tracker"
}
```

You can also change these settings from the dashboard (Settings menu).

## Windows Startup (Optional)

To have the tracker start automatically when you log in:

```bash
# Add to startup
python setup_startup.py --add

# Remove from startup
python setup_startup.py --remove

# Check status
python setup_startup.py --status
```

This creates a shortcut in your user Startup folder - **no admin rights required**.

## User Management

### Default Admin

On first run, a default admin user is created:
- Email: `admin@local`
- Password: `admin123`

### Creating Users

1. Log in as admin
2. Click your avatar → "Manage Users"
3. Fill in the new user form

### Roles

| Role | Permissions |
|------|-------------|
| `admin` | Create users, change settings, update any item |
| `user` | View items, update assigned items, add comments |

## API Reference

The tracker exposes a REST API:

### Authentication
- `POST /api/auth/login` - Login
- `POST /api/auth/logout` - Logout
- `GET /api/auth/me` - Get current user

### Items
- `GET /api/items` - List items (supports `?bucket=`, `?type=`, `?status=`)
- `GET /api/item/<id>` - Get single item
- `POST /api/item/<id>` - Update item
- `POST /api/items` - Create new item

### Comments
- `GET /api/comments/<item_id>` - List comments
- `POST /api/comments/<item_id>` - Add comment

### Other
- `GET /api/stats` - Dashboard statistics
- `GET /api/users` - List users
- `POST /api/users` - Create user (admin only)
- `GET /api/poll-status` - Email polling status
- `POST /api/poll-now` - Trigger immediate poll (admin only)
- `GET /api/config` - Get configuration (admin only)
- `POST /api/config` - Update configuration (admin only)

## Future: ACC Integration

The UI includes placeholder buttons for future ACC API integration:
- "Connect to ACC (Coming Soon)"
- "Import from ACC (Future)"

To prepare for this, the code is structured to allow adding an `acc_integration.py` module with functions like:

```python
def authenticate_acc(client_id, client_secret):
    """Authenticate with ACC API."""
    pass

def import_from_acc(project_id):
    """Import RFIs and Submittals from ACC."""
    pass
```

## Troubleshooting

### "Outlook not available"

- Ensure Microsoft Outlook desktop is installed (not just web version)
- Ensure you're logged into Outlook
- Install `pywin32`: `pip install pywin32`

### "Python not found"

- Add Python to your PATH, or
- Use the full path to Python in batch files

### Items not appearing

1. Check email polling status in the sidebar
2. Verify emails contain "LEB" in the subject
3. Check that emails also contain "RFI" or "Submittal"
4. Try the admin "Poll Now" action

### Port 5000 already in use

Edit `config.json` to change the port:

```json
{
    "server_port": 5001
}
```

## License

Internal use only - LEB Project Team

## Support

For issues or questions, contact your team administrator.
