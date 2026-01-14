"""
LEB Tracker - Windows Startup Setup
====================================
This script adds or removes the LEB Tracker from Windows startup.
It creates a shortcut in the user's Startup folder (no admin required).

Usage:
    python setup_startup.py --add      Add to startup
    python setup_startup.py --remove   Remove from startup
    python setup_startup.py --status   Check current status
"""

import os
import sys
import argparse
from pathlib import Path

# Get the directory where this script is located
SCRIPT_DIR = Path(__file__).parent.absolute()
START_ALL_BAT = SCRIPT_DIR / "start_all.bat"
SHORTCUT_NAME = "LEB Tracker.lnk"

def get_startup_folder():
    """Get the user's Startup folder path."""
    # This is the user-level Startup folder (no admin required)
    startup = Path(os.environ.get('APPDATA', '')) / 'Microsoft' / 'Windows' / 'Start Menu' / 'Programs' / 'Startup'
    return startup

def get_shortcut_path():
    """Get the full path to the shortcut in the Startup folder."""
    return get_startup_folder() / SHORTCUT_NAME

def create_shortcut_with_powershell(target_path, shortcut_path, working_dir, description=""):
    """Create a Windows shortcut using PowerShell (no extra dependencies)."""
    
    # Escape paths for PowerShell
    target_str = str(target_path).replace("'", "''")
    shortcut_str = str(shortcut_path).replace("'", "''")
    working_str = str(working_dir).replace("'", "''")
    desc_str = description.replace("'", "''")
    
    # PowerShell script to create shortcut
    ps_script = f'''
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut('{shortcut_str}')
$Shortcut.TargetPath = '{target_str}'
$Shortcut.WorkingDirectory = '{working_str}'
$Shortcut.Description = '{desc_str}'
$Shortcut.WindowStyle = 7
$Shortcut.Save()
'''
    
    # Run PowerShell
    import subprocess
    result = subprocess.run(
        ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_script],
        capture_output=True,
        text=True
    )
    
    return result.returncode == 0

def create_shortcut_with_winshell(target_path, shortcut_path, working_dir, description=""):
    """Create a Windows shortcut using winshell library (if available)."""
    try:
        import winshell
        
        winshell.CreateShortcut(
            Path=str(shortcut_path),
            Target=str(target_path),
            StartIn=str(working_dir),
            Description=description
        )
        return True
    except ImportError:
        return None  # Library not available
    except Exception as e:
        print(f"Error with winshell: {e}")
        return False

def create_shortcut_with_win32com(target_path, shortcut_path, working_dir, description=""):
    """Create a Windows shortcut using win32com (if available)."""
    try:
        import win32com.client
        
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(str(shortcut_path))
        shortcut.TargetPath = str(target_path)
        shortcut.WorkingDirectory = str(working_dir)
        shortcut.Description = description
        shortcut.WindowStyle = 7  # Minimized
        shortcut.save()
        return True
    except ImportError:
        return None  # Library not available
    except Exception as e:
        print(f"Error with win32com: {e}")
        return False

def add_to_startup():
    """Add LEB Tracker to Windows startup."""
    
    # Verify start_all.bat exists
    if not START_ALL_BAT.exists():
        print(f"ERROR: {START_ALL_BAT} not found!")
        print("Make sure you're running this script from the LEB Tracker directory.")
        return False
    
    # Get paths
    startup_folder = get_startup_folder()
    shortcut_path = get_shortcut_path()
    
    # Verify startup folder exists
    if not startup_folder.exists():
        print(f"ERROR: Startup folder not found: {startup_folder}")
        return False
    
    print(f"Creating startup shortcut...")
    print(f"  Target: {START_ALL_BAT}")
    print(f"  Shortcut: {shortcut_path}")
    
    # Try different methods to create shortcut
    description = "LEB RFI/Submittal Tracker"
    
    # Method 1: Try win32com (most reliable if available)
    result = create_shortcut_with_win32com(START_ALL_BAT, shortcut_path, SCRIPT_DIR, description)
    if result is True:
        print("\nSuccess! LEB Tracker added to Windows startup.")
        print("The tracker will now start automatically when you log in.")
        return True
    
    # Method 2: Try winshell
    result = create_shortcut_with_winshell(START_ALL_BAT, shortcut_path, SCRIPT_DIR, description)
    if result is True:
        print("\nSuccess! LEB Tracker added to Windows startup.")
        print("The tracker will now start automatically when you log in.")
        return True
    
    # Method 3: Fall back to PowerShell
    result = create_shortcut_with_powershell(START_ALL_BAT, shortcut_path, SCRIPT_DIR, description)
    if result:
        print("\nSuccess! LEB Tracker added to Windows startup.")
        print("The tracker will now start automatically when you log in.")
        return True
    
    print("\nERROR: Failed to create shortcut.")
    print("You can manually add the shortcut:")
    print(f"  1. Right-click on {START_ALL_BAT}")
    print(f"  2. Select 'Create shortcut'")
    print(f"  3. Move the shortcut to: {startup_folder}")
    return False

def remove_from_startup():
    """Remove LEB Tracker from Windows startup."""
    
    shortcut_path = get_shortcut_path()
    
    if not shortcut_path.exists():
        print("LEB Tracker is not currently in Windows startup.")
        return True
    
    try:
        shortcut_path.unlink()
        print("Success! LEB Tracker removed from Windows startup.")
        return True
    except Exception as e:
        print(f"ERROR: Failed to remove shortcut: {e}")
        print(f"Please manually delete: {shortcut_path}")
        return False

def check_status():
    """Check if LEB Tracker is in Windows startup."""
    
    shortcut_path = get_shortcut_path()
    
    print("LEB Tracker Startup Status")
    print("=" * 40)
    print(f"Startup folder: {get_startup_folder()}")
    print(f"Shortcut path: {shortcut_path}")
    print()
    
    if shortcut_path.exists():
        print("Status: ENABLED")
        print("LEB Tracker will start automatically when you log in.")
    else:
        print("Status: DISABLED")
        print("Run 'python setup_startup.py --add' to enable.")

def main():
    parser = argparse.ArgumentParser(
        description="Add or remove LEB Tracker from Windows startup"
    )
    
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--add', action='store_true', help='Add to Windows startup')
    group.add_argument('--remove', action='store_true', help='Remove from Windows startup')
    group.add_argument('--status', action='store_true', help='Check current status')
    
    args = parser.parse_args()
    
    if args.add:
        success = add_to_startup()
        sys.exit(0 if success else 1)
    elif args.remove:
        success = remove_from_startup()
        sys.exit(0 if success else 1)
    elif args.status:
        check_status()
        sys.exit(0)

if __name__ == '__main__':
    main()
