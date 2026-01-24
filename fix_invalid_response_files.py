import os
from pathlib import Path

# Base path
base_path = Path(r"\\sac-filsrv1\Projects\Structural-028\Projects\LEB\9.0_Const_Svcs")

# List of problematic response files from the error log
problem_files = [
    r"Turner\Submittals\Submittal - 13 34 19-01 - (Rev 1) LEB10_133419_ShpDwg_Reactions-AnchorBoltDwg_Rev01\Responses\_reviewer_response.json",
    r"Fti\Submittals\Submittal - 4 - 260 - MES (Mechanical Equipment Skid)\Responses\_reviewer_response.json"
]

print("Renaming old/invalid response files to prevent watcher errors...\n")

for rel_path in problem_files:
    full_path = base_path / rel_path
    
    if full_path.exists():
        # Rename to .old to exclude from watcher
        new_path = full_path.with_suffix('.json.old')
        try:
            full_path.rename(new_path)
            print(f"✅ Renamed: {rel_path}")
        except Exception as e:
            print(f"❌ Failed to rename {rel_path}: {e}")
    else:
        print(f"⚠️  Not found: {rel_path}")

print("\nDone! These files are now ignored by the watcher.")
