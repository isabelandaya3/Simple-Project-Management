import sys
sys.path.insert(0, '.')
from pathlib import Path
from app import process_reviewer_response_json

# Process the restored response
json_file = Path(r"\\sac-filsrv1\Projects\Structural-028\Projects\LEB\9.0_Const_Svcs\Fti\Submittals\Submittal - 4 - 260 - MES (Mechanical Equipment Skid)\Responses\_reviewer_response.json")

print(f"Processing: {json_file}")
result = process_reviewer_response_json(json_file)

if result['success']:
    print(f"✅ Success!")
    print(f"   Item ID: {result.get('item_id')}")
    print(f"   Version: {result.get('version')}")
else:
    print(f"❌ Error: {result.get('error')}")
