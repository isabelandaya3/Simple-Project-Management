import sys
sys.path.insert(0, '.')
from pathlib import Path
from app import process_reviewer_response_json

# Process the RFI #33 response
json_file = Path(r"\\sac-filsrv1\Projects\Structural-028\Projects\LEB\9.0_Const_Svcs\General\RFIs\RFI - 33 - Concrete_Leave Down_Anchor Bolts\Responses\_RESPONSE_2026-01-24T00-00-22-653Z.json")

print(f"Processing: {json_file.name}")
result = process_reviewer_response_json(json_file)

if result['success']:
    print(f"✅ Success!")
    print(f"   Item ID: {result.get('item_id')}")
    print(f"   Version: {result.get('version')}")
else:
    print(f"❌ Error: {result.get('error')}")
