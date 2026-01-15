"""
Airtable Integration for LEB Tracker
=====================================
This module provides integration with Airtable for collecting form responses
when the local server is not running. Responses are stored in Airtable and
synced back when the local server starts.

Setup:
1. Create a free Airtable account at https://airtable.com
2. Create a new base called "LEB Tracker Responses"
3. Create two tables: "Reviewer Responses" and "QCR Responses"
4. Get your API key from https://airtable.com/create/tokens
5. Add the configuration to config.json

Airtable Tables Structure:

Reviewer Responses:
- item_id (Single line text)
- token (Single line text)  
- identifier (Single line text)
- title (Single line text)
- response_category (Single select: Approved, Approved as Noted, For Record Only, Rejected, Revise and Resubmit)
- selected_files (Long text)
- notes (Long text)
- response_text (Long text)
- submitted_at (Date time)
- synced (Checkbox)

QCR Responses:
- item_id (Single line text)
- token (Single line text)
- identifier (Single line text)
- title (Single line text)
- qc_action (Single select: Approve, Modify, Send Back)
- response_mode (Single select: Keep, Tweak, Revise)
- response_category (Single select: Approved, Approved as Noted, For Record Only, Rejected, Revise and Resubmit)
- response_text (Long text)
- selected_files (Long text)
- qcr_notes (Long text)
- submitted_at (Date time)
- synced (Checkbox)
"""

import json
import urllib.parse
from pathlib import Path
from datetime import datetime

# Configuration
BASE_DIR = Path(__file__).parent.absolute()
CONFIG_PATH = BASE_DIR / "config.json"


def load_airtable_config():
    """Load Airtable configuration from config.json."""
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, 'r') as f:
            config = json.load(f)
            return config.get('airtable', {})
    return {}


def get_airtable_form_url(form_type, item_data, token):
    """
    Generate an Airtable form URL with pre-filled data.
    
    Args:
        form_type: 'reviewer' or 'qcr'
        item_data: dict containing item information
        token: the magic link token
    
    Returns:
        Airtable form URL with prefilled parameters
    """
    config = load_airtable_config()
    
    if form_type == 'reviewer':
        form_id = config.get('reviewer_form_id', '')
    else:
        form_id = config.get('qcr_form_id', '')
    
    if not form_id:
        return None
    
    # Build prefill parameters
    # Airtable uses prefill_{Field Name} format in URL
    prefill_params = {
        'prefill_item_id': str(item_data.get('id', '')),
        'prefill_token': token,
        'prefill_identifier': item_data.get('identifier', ''),
        'prefill_title': item_data.get('title', ''),
    }
    
    # Add QCR-specific prefills
    if form_type == 'qcr':
        prefill_params['prefill_reviewer_category'] = item_data.get('reviewer_response_category', '')
        prefill_params['prefill_reviewer_notes'] = item_data.get('reviewer_notes', '')
        prefill_params['prefill_reviewer_files'] = item_data.get('reviewer_selected_files', '')
        prefill_params['prefill_reviewer_response_text'] = item_data.get('reviewer_response_text', '')
    
    # Build URL
    base_url = f"https://airtable.com/{form_id}"
    query_string = urllib.parse.urlencode(prefill_params)
    
    return f"{base_url}?{query_string}"


def get_airtable_api_headers():
    """Get headers for Airtable API requests."""
    config = load_airtable_config()
    api_key = config.get('api_key', '')
    
    return {
        'Authorization': f'Bearer {api_key}',
        'Content-Type': 'application/json'
    }


def sync_airtable_responses():
    """
    Sync responses from Airtable back to the local database.
    Call this when the server starts or on demand.
    
    Returns:
        dict with success status and count of synced records
    """
    import requests
    
    config = load_airtable_config()
    
    if not config.get('api_key') or not config.get('base_id'):
        return {'success': False, 'error': 'Airtable not configured'}
    
    base_id = config['base_id']
    api_key = config['api_key']
    
    headers = {
        'Authorization': f'Bearer {api_key}',
        'Content-Type': 'application/json'
    }
    
    synced_count = 0
    errors = []
    
    # Sync Reviewer Responses
    reviewer_table = config.get('reviewer_table_name', 'Reviewer Responses')
    try:
        url = f"https://api.airtable.com/v0/{base_id}/{urllib.parse.quote(reviewer_table)}"
        params = {
            'filterByFormula': 'NOT({synced})',  # Only get un-synced records
            'sort[0][field]': 'submitted_at',
            'sort[0][direction]': 'asc'
        }
        
        response = requests.get(url, headers=headers, params=params)
        
        if response.status_code == 200:
            records = response.json().get('records', [])
            for record in records:
                result = process_reviewer_response(record)
                if result['success']:
                    # Mark as synced in Airtable
                    mark_synced(base_id, reviewer_table, record['id'], headers)
                    synced_count += 1
                else:
                    errors.append(f"Reviewer {record['id']}: {result.get('error')}")
        else:
            errors.append(f"Failed to fetch reviewer responses: {response.status_code}")
            
    except Exception as e:
        errors.append(f"Reviewer sync error: {str(e)}")
    
    # Sync QCR Responses
    qcr_table = config.get('qcr_table_name', 'QCR Responses')
    try:
        url = f"https://api.airtable.com/v0/{base_id}/{urllib.parse.quote(qcr_table)}"
        params = {
            'filterByFormula': 'NOT({synced})',
            'sort[0][field]': 'submitted_at',
            'sort[0][direction]': 'asc'
        }
        
        response = requests.get(url, headers=headers, params=params)
        
        if response.status_code == 200:
            records = response.json().get('records', [])
            for record in records:
                result = process_qcr_response(record)
                if result['success']:
                    mark_synced(base_id, qcr_table, record['id'], headers)
                    synced_count += 1
                else:
                    errors.append(f"QCR {record['id']}: {result.get('error')}")
        else:
            errors.append(f"Failed to fetch QCR responses: {response.status_code}")
            
    except Exception as e:
        errors.append(f"QCR sync error: {str(e)}")
    
    return {
        'success': len(errors) == 0,
        'synced_count': synced_count,
        'errors': errors
    }


def mark_synced(base_id, table_name, record_id, headers):
    """Mark a record as synced in Airtable."""
    import requests
    
    url = f"https://api.airtable.com/v0/{base_id}/{urllib.parse.quote(table_name)}/{record_id}"
    data = {
        'fields': {
            'synced': True
        }
    }
    
    try:
        requests.patch(url, headers=headers, json=data)
    except:
        pass  # Non-critical if this fails


def process_reviewer_response(record):
    """
    Process a reviewer response from Airtable and update local database.
    
    Args:
        record: Airtable record dict
    
    Returns:
        dict with success status
    """
    from app import get_db, send_qcr_assignment_email
    
    fields = record.get('fields', {})
    
    token = fields.get('token', '')
    if not token:
        return {'success': False, 'error': 'No token provided'}
    
    try:
        conn = get_db()
        cursor = conn.cursor()
        
        # Find item by token
        cursor.execute('SELECT * FROM item WHERE email_token_reviewer = ?', (token,))
        item = cursor.fetchone()
        
        if not item:
            conn.close()
            return {'success': False, 'error': 'Item not found for token'}
        
        # Check if already responded
        if item['reviewer_response_status'] == 'Responded':
            conn.close()
            return {'success': False, 'error': 'Already responded'}
        
        # Get response data
        response_category = fields.get('response_category', '')
        selected_files = fields.get('selected_files', '')
        notes = fields.get('notes', '')
        response_text = fields.get('response_text', '')
        submitted_at = fields.get('submitted_at', datetime.now().isoformat())
        
        # Parse selected files (comma or newline separated)
        if selected_files:
            files_list = [f.strip() for f in selected_files.replace('\n', ',').split(',') if f.strip()]
            selected_files_json = json.dumps(files_list)
        else:
            selected_files_json = '[]'
        
        # Get current version
        current_version = item['reviewer_response_version'] or 0
        new_version = current_version + 1
        
        # Update item
        cursor.execute('''
            UPDATE item SET
                reviewer_response_category = ?,
                reviewer_selected_files = ?,
                reviewer_notes = ?,
                reviewer_response_text = ?,
                reviewer_response_at = ?,
                reviewer_response_status = 'Responded',
                reviewer_response_version = ?,
                status = 'Pending QC'
            WHERE id = ?
        ''', (
            response_category,
            selected_files_json,
            notes,
            response_text,
            submitted_at,
            new_version,
            item['id']
        ))
        
        conn.commit()
        conn.close()
        
        # Send notification to QCR
        try:
            send_qcr_assignment_email(item['id'], is_revision=(new_version > 1), version=new_version)
        except:
            pass  # Non-critical
        
        return {'success': True}
        
    except Exception as e:
        return {'success': False, 'error': str(e)}


def process_qcr_response(record):
    """
    Process a QCR response from Airtable and update local database.
    
    Args:
        record: Airtable record dict
    
    Returns:
        dict with success status
    """
    from app import get_db
    
    fields = record.get('fields', {})
    
    token = fields.get('token', '')
    if not token:
        return {'success': False, 'error': 'No token provided'}
    
    try:
        conn = get_db()
        cursor = conn.cursor()
        
        # Find item by token
        cursor.execute('SELECT * FROM item WHERE email_token_qcr = ?', (token,))
        item = cursor.fetchone()
        
        if not item:
            conn.close()
            return {'success': False, 'error': 'Item not found for token'}
        
        # Get response data
        qc_action = fields.get('qc_action', '')
        response_mode = fields.get('response_mode', '')
        response_category = fields.get('response_category', '')
        response_text = fields.get('response_text', '')
        selected_files = fields.get('selected_files', '')
        qcr_notes = fields.get('qcr_notes', '')
        submitted_at = fields.get('submitted_at', datetime.now().isoformat())
        
        # Parse selected files
        if selected_files:
            files_list = [f.strip() for f in selected_files.replace('\n', ',').split(',') if f.strip()]
            selected_files_json = json.dumps(files_list)
        else:
            selected_files_json = item['reviewer_selected_files'] or '[]'
        
        if qc_action == 'Send Back':
            # Send back to reviewer
            cursor.execute('''
                UPDATE item SET
                    qcr_notes = ?,
                    qcr_response_at = ?,
                    qcr_action = 'Send Back',
                    reviewer_response_status = 'Revision Requested',
                    status = 'Revision Requested'
                WHERE id = ?
            ''', (qcr_notes, submitted_at, item['id']))
            
            # Would trigger send_back_to_reviewer_email here
            
        else:
            # Approve or Modify
            final_text = response_text if response_text else item['reviewer_response_text']
            final_category = response_category if response_category else item['reviewer_response_category']
            
            cursor.execute('''
                UPDATE item SET
                    qcr_action = ?,
                    qcr_response_mode = ?,
                    final_response_category = ?,
                    final_response_text = ?,
                    final_selected_files = ?,
                    qcr_notes = ?,
                    qcr_response_at = ?,
                    status = 'Complete'
                WHERE id = ?
            ''', (
                qc_action,
                response_mode,
                final_category,
                final_text,
                selected_files_json,
                qcr_notes,
                submitted_at,
                item['id']
            ))
        
        conn.commit()
        conn.close()
        
        return {'success': True}
        
    except Exception as e:
        return {'success': False, 'error': str(e)}


def generate_email_body_with_airtable(item, form_type, token, local_url):
    """
    Generate email HTML body that includes both local and Airtable form links.
    The Airtable form is presented as a fallback when the local server is not running.
    
    Args:
        item: Item data dict
        form_type: 'reviewer' or 'qcr'
        token: Magic link token
        local_url: The localhost URL for the form
    
    Returns:
        HTML string for the email body
    """
    config = load_airtable_config()
    airtable_url = get_airtable_form_url(form_type, dict(item), token)
    
    # Build the fallback section if Airtable is configured
    airtable_section = ""
    if airtable_url:
        airtable_section = f"""
    <!-- FALLBACK OPTION -->
    <div style="margin-top:24px; padding:16px; background:#fff8e6; border:1px solid #ffd666; border-radius:8px;">
        <div style="font-weight:bold; color:#d48806; margin-bottom:8px;">
            🌐 Working Remotely or Server Unavailable?
        </div>
        <p style="font-size:13px; color:#666; margin-bottom:12px;">
            If the button above doesn't work (server offline), use this alternative form:
        </p>
        <a href="{airtable_url}"
           style="background:#ffd666; color:#333; padding:10px 20px; 
                  font-size:14px; font-weight:600; text-decoration:none; border-radius:6px; display:inline-block;">
            📝 Use Online Form (Airtable)
        </a>
        <p style="font-size:11px; color:#888; margin-top:8px;">
            Your response will sync automatically when the server is back online.
        </p>
    </div>
"""
    
    return airtable_section


# Airtable form field mappings for easy setup reference
AIRTABLE_SETUP_GUIDE = """
=============================================================================
AIRTABLE SETUP GUIDE
=============================================================================

1. CREATE AIRTABLE ACCOUNT
   Go to https://airtable.com and sign up (free tier is sufficient)

2. CREATE A NEW BASE
   Name it "LEB Tracker Responses"

3. CREATE "Reviewer Responses" TABLE
   Add these fields:
   - item_id (Single line text) - Required
   - token (Single line text) - Required, Hidden in form
   - identifier (Single line text) - Read-only in form
   - title (Single line text) - Read-only in form
   - response_category (Single select)
     Options: Approved, Approved as Noted, For Record Only, Rejected, Revise and Resubmit
   - selected_files (Long text) - Instructions: "Enter file names, one per line"
   - notes (Long text)
   - response_text (Long text)
   - submitted_at (Created time) - Auto-filled
   - synced (Checkbox) - Hidden, default false

4. CREATE "QCR Responses" TABLE
   Add these fields:
   - item_id (Single line text) - Required
   - token (Single line text) - Required, Hidden in form
   - identifier (Single line text) - Read-only in form  
   - title (Single line text) - Read-only in form
   - reviewer_category (Single line text) - Read-only, shows what reviewer selected
   - reviewer_notes (Long text) - Read-only, shows reviewer's notes
   - reviewer_files (Long text) - Read-only, shows reviewer's selected files
   - qc_action (Single select)
     Options: Approve, Modify, Send Back
   - response_mode (Single select) - Conditional, shown only if qc_action is Approve/Modify
     Options: Keep, Tweak, Revise
   - response_category (Single select) - Conditional, for final category
     Options: Approved, Approved as Noted, For Record Only, Rejected, Revise and Resubmit
   - response_text (Long text) - For modified/revised text
   - selected_files (Long text)
   - qcr_notes (Long text)
   - submitted_at (Created time) - Auto-filled
   - synced (Checkbox) - Hidden, default false

5. CREATE FORMS
   - Click "Create form" for each table
   - Customize which fields are visible and required
   - Get the form URL (e.g., https://airtable.com/appXXXXX/shrYYYYY)

6. GET API CREDENTIALS
   - Go to https://airtable.com/create/tokens
   - Create a new token with:
     - Scopes: data.records:read, data.records:write
     - Access: Your "LEB Tracker Responses" base
   - Copy the API key

7. UPDATE config.json
   Add this to your config.json:

   "airtable": {
       "api_key": "pat_YOUR_API_KEY_HERE",
       "base_id": "appXXXXXXXXXXXXXX",
       "reviewer_form_id": "shrYYYYYYYYYYYYYY",
       "qcr_form_id": "shrZZZZZZZZZZZZZZ",
       "reviewer_table_name": "Reviewer Responses",
       "qcr_table_name": "QCR Responses"
   }

   The base_id can be found in the Airtable URL when viewing your base:
   https://airtable.com/appXXXXXXXXXXXXXX/...

=============================================================================
"""

def print_setup_guide():
    """Print the Airtable setup guide."""
    print(AIRTABLE_SETUP_GUIDE)


if __name__ == '__main__':
    print_setup_guide()
