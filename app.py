"""
LEB RFI/Submittal Tracker - Local Windows Application
=====================================================
A local-only tracker that monitors Outlook emails from Autodesk Construction Cloud
and provides a web dashboard for managing RFIs and Submittals.

Run: python app.py
Access: http://localhost:5000
"""

import os
import re
import json
import sqlite3
import hashlib
import secrets
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path
from functools import wraps

from flask import Flask, request, jsonify, send_from_directory, session, redirect, url_for, render_template_string
import bcrypt

# Optional: dateutil for flexible date parsing
try:
    from dateutil import parser as date_parser
    HAS_DATEUTIL = True
except ImportError:
    HAS_DATEUTIL = False

# Optional: win32com for Outlook integration
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
    print("WARNING: pywin32 not installed. Email polling will be disabled.")
    print("Install with: pip install pywin32")

# Optional: Windows toast notifications
try:
    from winotify import Notification, audio
    HAS_WINOTIFY = True
except ImportError:
    HAS_WINOTIFY = False
    print("WARNING: winotify not installed. System notifications will be disabled.")
    print("Install with: pip install winotify")

# =============================================================================
# CONFIGURATION
# =============================================================================

BASE_DIR = Path(__file__).parent.absolute()
DATABASE_PATH = BASE_DIR / "tracker.db"
CONFIG_PATH = BASE_DIR / "config.json"

# Default configuration
DEFAULT_CONFIG = {
    "base_folder_path": r"\\sac-filsrv1\Projects\Structural-028\Projects\LEB\9.0_Const_Svcs",
    "outlook_folder": "Inbox",  # Can be "Inbox" or a subfolder like "LEB ACC"
    "poll_interval_minutes": 5,
    "server_port": 5000,
    "project_name": "LEB – Local Tracker"
}

def load_config():
    """Load configuration from config.json or create default."""
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, 'r') as f:
            config = json.load(f)
            # Merge with defaults for any missing keys
            for key, value in DEFAULT_CONFIG.items():
                if key not in config:
                    config[key] = value
            return config
    else:
        # Create default config file
        with open(CONFIG_PATH, 'w') as f:
            json.dump(DEFAULT_CONFIG, f, indent=2)
        return DEFAULT_CONFIG.copy()

CONFIG = load_config()

# =============================================================================
# REVIEW DUE DATE CONFIGURATION
# =============================================================================

# Priority-based minimum work days required from date_received to contractor due_date
PRIORITY_MIN_DAYS = {
    'High': 5,
    'Medium': 10,
    'Low': 15,
    None: 10  # Default to Medium if no priority set
}

# QCR requires at least this many work days before contractor due date
QCR_DAYS_BEFORE_DUE = 3

# =============================================================================
# WORKDAY HELPER FUNCTIONS
# =============================================================================

def is_business_day(date):
    """Check if a date is a business day (Mon-Fri)."""
    return date.weekday() < 5  # 0=Monday, 4=Friday

def business_days_between(start_date, end_date):
    """
    Count business days (Mon-Fri) from start_date to end_date.
    Inclusive of start_date, exclusive of end_date.
    """
    if isinstance(start_date, str):
        start_date = datetime.strptime(start_date, '%Y-%m-%d').date()
    if isinstance(end_date, str):
        end_date = datetime.strptime(end_date, '%Y-%m-%d').date()
    
    if hasattr(start_date, 'date'):
        start_date = start_date.date()
    if hasattr(end_date, 'date'):
        end_date = end_date.date()
    
    if start_date >= end_date:
        return 0
    
    count = 0
    current = start_date
    while current < end_date:
        if is_business_day(current):
            count += 1
        current += timedelta(days=1)
    return count

def subtract_business_days(end_date, n):
    """
    Return the date that is n business days before end_date.
    If n=3 and end_date is Friday, result is Tuesday (skipping Sat/Sun).
    """
    if isinstance(end_date, str):
        end_date = datetime.strptime(end_date, '%Y-%m-%d').date()
    if hasattr(end_date, 'date'):
        end_date = end_date.date()
    
    current = end_date
    days_subtracted = 0
    while days_subtracted < n:
        current -= timedelta(days=1)
        if is_business_day(current):
            days_subtracted += 1
    return current

def add_business_days(start_date, n):
    """
    Return the date that is n business days after start_date.
    """
    if isinstance(start_date, str):
        start_date = datetime.strptime(start_date, '%Y-%m-%d').date()
    if hasattr(start_date, 'date'):
        start_date = start_date.date()
    
    current = start_date
    days_added = 0
    while days_added < n:
        current += timedelta(days=1)
        if is_business_day(current):
            days_added += 1
    return current

def calculate_review_due_dates(date_received, contractor_due_date, priority):
    """
    Calculate internal due dates for Initial Reviewer and QCR.
    
    Returns dict with:
        - initial_reviewer_due_date
        - qcr_due_date
        - contractor_window_days
        - is_contractor_window_insufficient
        - required_days
    """
    if not date_received or not contractor_due_date:
        return {
            'initial_reviewer_due_date': None,
            'qcr_due_date': None,
            'contractor_window_days': None,
            'is_contractor_window_insufficient': False,
            'required_days': None
        }
    
    # Parse dates
    if isinstance(date_received, str):
        date_received = datetime.strptime(date_received, '%Y-%m-%d').date()
    if isinstance(contractor_due_date, str):
        contractor_due_date = datetime.strptime(contractor_due_date, '%Y-%m-%d').date()
    if hasattr(date_received, 'date'):
        date_received = date_received.date()
    if hasattr(contractor_due_date, 'date'):
        contractor_due_date = contractor_due_date.date()
    
    # Get required minimum days based on priority
    required_days = PRIORITY_MIN_DAYS.get(priority, PRIORITY_MIN_DAYS[None])
    
    # Calculate contractor window (business days available)
    contractor_window_days = business_days_between(date_received, contractor_due_date)
    
    # Check if window is insufficient
    is_insufficient = contractor_window_days < required_days
    
    # Calculate QCR due date (3 business days before contractor due date)
    qcr_due_date = subtract_business_days(contractor_due_date, QCR_DAYS_BEFORE_DUE)
    
    # Calculate Initial Reviewer due date
    # Reserve QCR_DAYS_BEFORE_DUE for QCR, rest for Initial Reviewer
    review_window_days = max(required_days - QCR_DAYS_BEFORE_DUE, 1)
    initial_reviewer_due_date = subtract_business_days(qcr_due_date, review_window_days)
    
    # Clamp to not be before date_received
    if initial_reviewer_due_date < date_received:
        initial_reviewer_due_date = date_received
    
    return {
        'initial_reviewer_due_date': initial_reviewer_due_date.strftime('%Y-%m-%d'),
        'qcr_due_date': qcr_due_date.strftime('%Y-%m-%d'),
        'contractor_window_days': contractor_window_days,
        'is_contractor_window_insufficient': is_insufficient,
        'required_days': required_days
    }

def get_due_date_status(due_date):
    """
    Get status color for a due date based on business days remaining.
    Returns: 'green', 'yellow', 'red', or None if no due date.
    """
    if not due_date:
        return None
    
    if isinstance(due_date, str):
        due_date = datetime.strptime(due_date, '%Y-%m-%d').date()
    if hasattr(due_date, 'date'):
        due_date = due_date.date()
    
    today = datetime.now().date()
    days_remaining = business_days_between(today, due_date)
    
    if due_date < today:
        return 'red'  # Overdue
    elif days_remaining <= 3:
        return 'yellow'  # 1-3 days remaining
    else:
        return 'green'  # More than 3 days

# =============================================================================
# FLASK APP SETUP
# =============================================================================

app = Flask(__name__, static_folder='static', template_folder='templates')

# Use a persistent secret key (stored in config or generate once)
SECRET_KEY_FILE = BASE_DIR / ".secret_key"
if SECRET_KEY_FILE.exists():
    app.secret_key = SECRET_KEY_FILE.read_text().strip()
else:
    app.secret_key = secrets.token_hex(32)
    SECRET_KEY_FILE.write_text(app.secret_key)

app.config['SESSION_TYPE'] = 'filesystem'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=24)
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SECURE'] = False  # Set True if using HTTPS

# =============================================================================
# HTML TEMPLATES FOR MAGIC-LINK RESPONSE PAGES
# =============================================================================

BASE_PAGE_STYLE = '''
<style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        min-height: 100vh;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 20px;
    }
    .container {
        background: white;
        border-radius: 16px;
        box-shadow: 0 25px 50px rgba(0,0,0,0.15);
        padding: 40px;
        max-width: 700px;
        width: 100%;
    }
    h1 {
        color: #1a1a2e;
        font-size: 24px;
        margin-bottom: 8px;
    }
    .subtitle {
        color: #666;
        font-size: 14px;
        margin-bottom: 24px;
    }
    .info-box {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 24px;
    }
    .info-row {
        display: flex;
        margin-bottom: 8px;
    }
    .info-row:last-child { margin-bottom: 0; }
    .info-label {
        font-weight: 600;
        color: #444;
        width: 140px;
        flex-shrink: 0;
    }
    .info-value {
        color: #666;
    }
    .section-title {
        font-size: 16px;
        font-weight: 600;
        color: #1a1a2e;
        margin-bottom: 12px;
    }
    .form-group {
        margin-bottom: 20px;
    }
    label {
        display: block;
        font-weight: 500;
        color: #444;
        margin-bottom: 6px;
    }
    select, textarea {
        width: 100%;
        padding: 10px 12px;
        border: 1px solid #ddd;
        border-radius: 8px;
        font-size: 14px;
        font-family: inherit;
    }
    select:focus, textarea:focus {
        outline: none;
        border-color: #667eea;
        box-shadow: 0 0 0 3px rgba(102,126,234,0.1);
    }
    textarea { resize: vertical; min-height: 100px; }
    .file-list {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 12px;
        max-height: 200px;
        overflow-y: auto;
    }
    .file-item {
        display: flex;
        align-items: center;
        padding: 8px;
        border-radius: 6px;
        margin-bottom: 4px;
    }
    .file-item:hover { background: #e9ecef; }
    .file-item input { margin-right: 10px; }
    .file-item label {
        margin-bottom: 0;
        font-weight: normal;
        color: #333;
        cursor: pointer;
    }
    .btn {
        display: inline-block;
        padding: 12px 24px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        font-size: 16px;
        font-weight: 600;
        cursor: pointer;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 20px rgba(102,126,234,0.4);
    }
    .reviewer-notes-box {
        background: #fff3cd;
        border: 1px solid #ffc107;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 24px;
    }
    .reviewer-notes-box h3 {
        font-size: 14px;
        color: #856404;
        margin-bottom: 8px;
    }
    .reviewer-notes-box p {
        color: #856404;
        font-size: 14px;
        white-space: pre-wrap;
    }
    .error-container {
        text-align: center;
    }
    .error-icon {
        font-size: 64px;
        margin-bottom: 20px;
    }
    .success-container {
        text-align: center;
    }
    .success-icon {
        font-size: 64px;
        margin-bottom: 20px;
    }
</style>
'''

ERROR_PAGE_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Error - LEB Tracker</title>
    ''' + BASE_PAGE_STYLE + '''
</head>
<body>
    <div class="container error-container">
        <div class="error-icon">❌</div>
        <h1>Error</h1>
        <p style="color: #666; margin-top: 12px;">{{ error }}</p>
    </div>
</body>
</html>
'''

ALREADY_RESPONDED_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Already Submitted - LEB Tracker</title>
    ''' + BASE_PAGE_STYLE + '''
</head>
<body>
    <div class="container success-container">
        <div class="success-icon">✅</div>
        <h1>Already Submitted</h1>
        <p style="color: #666; margin-top: 12px;">
            This {{ 'review' if response_type == 'reviewer' else 'QC review' }} for 
            <strong>{{ item.type }} {{ item.identifier }}</strong> has already been submitted.
        </p>
    </div>
</body>
</html>
'''

SUCCESS_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Success - LEB Tracker</title>
    ''' + BASE_PAGE_STYLE + '''
</head>
<body>
    <div class="container success-container">
        <div class="success-icon">✅</div>
        <h1>{{ message }}</h1>
        <p style="color: #666; margin-top: 12px;">{{ details }}</p>
    </div>
</body>
</html>
'''

REVIEWER_RESPONSE_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Review Response - {{ item.type }} {{ item.identifier }}</title>
    ''' + BASE_PAGE_STYLE + '''
    <style>
        .version-badge {
            display: inline-block;
            background: #667eea;
            color: white;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            margin-left: 10px;
        }
        .resubmit-notice {
            background: #fef3c7;
            border: 1px solid #f59e0b;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .resubmit-notice h3 {
            margin: 0 0 8px 0;
            color: #92400e;
        }
        .previous-response {
            background: #f0f9ff;
            border: 1px solid #bae6fd;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .previous-response h4 {
            margin: 0 0 10px 0;
            color: #0369a1;
        }
        .qcr-feedback {
            background: #fef2f2;
            border: 1px solid #fecaca;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .qcr-feedback h4 {
            margin: 0 0 10px 0;
            color: #991b1b;
        }
        .version-history {
            font-size: 12px;
            color: #666;
            margin-top: 10px;
        }
        .closed-notice {
            background: #f3f4f6;
            border: 1px solid #d1d5db;
            border-radius: 8px;
            padding: 20px;
            text-align: center;
        }
        .closed-notice h3 {
            color: #6b7280;
            margin: 0 0 10px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Initial Review Response <span class="version-badge">v{{ version }}</span></h1>
        <p class="subtitle">{{ item.type }} {{ item.identifier }}</p>
        
        {% if is_closed %}
        <div class="closed-notice">
            <h3>🔒 This Item Has Been Closed</h3>
            <p>No further changes can be submitted. Contact the project administrator if this is unexpected.</p>
        </div>
        {% elif not can_submit %}
        <div class="closed-notice">
            <h3>⚠️ Submission Not Allowed</h3>
            <p>This item has been finalized in QC. Contact the project administrator if additional changes are required.</p>
        </div>
        {% else %}
        
        {% if is_resubmit and qcr_feedback %}
        <div class="qcr-feedback">
            <h4>↩️ QC Reviewer Requested Revisions</h4>
            <p><strong>Feedback on your v{{ version - 1 }} response:</strong></p>
            <div style="background: white; padding: 10px; border-radius: 4px; margin-top: 8px;">
                {{ qcr_feedback|replace('\n', '<br>')|safe }}
            </div>
        </div>
        {% endif %}
        
        {% if is_resubmit and previous_response %}
        <div class="previous-response">
            <h4>📄 Your Previous Response (v{{ version - 1 }})</h4>
            <p><strong>Category:</strong> {{ previous_response.category or 'N/A' }}</p>
            <p><strong>Notes:</strong></p>
            <div style="background: white; padding: 10px; border-radius: 4px; margin-top: 8px;">
                {{ (previous_response.text or 'No notes')|replace('\n', '<br>')|safe }}
            </div>
            {% if version_history %}
            <div class="version-history">
                <strong>Version History:</strong> {{ version_history }}
            </div>
            {% endif %}
        </div>
        {% endif %}
        
        {% if is_resubmit %}
        <div class="resubmit-notice">
            <h3>📝 Resubmitting Response</h3>
            <p>You are updating your response to version {{ version }}. The QC Reviewer will be notified of your changes.</p>
        </div>
        {% endif %}
        
        <div class="info-box">
            <div class="info-row">
                <span class="info-label">Title:</span>
                <span class="info-value">{{ item.title or 'N/A' }}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Date Received:</span>
                <span class="info-value">{{ item.date_received or 'N/A' }}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Your Due Date:</span>
                <span class="info-value">{{ item.initial_reviewer_due_date or 'N/A' }}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Folder:</span>
                <span class="info-value">{{ item.folder_link or 'N/A' }}</span>
            </div>
        </div>
        
        <form method="POST">
            <input type="hidden" name="token" value="{{ token }}">
            
            <div class="form-group">
                <label for="response_category">Response Category *</label>
                <select name="response_category" id="response_category" required>
                    <option value="">-- Select --</option>
                    <option value="Approved" {% if previous_response and previous_response.category == 'Approved' %}selected{% endif %}>Approved</option>
                    <option value="Approved as Noted" {% if previous_response and previous_response.category == 'Approved as Noted' %}selected{% endif %}>Approved as Noted</option>
                    <option value="For Record Only" {% if previous_response and previous_response.category == 'For Record Only' %}selected{% endif %}>For Record Only</option>
                    <option value="Rejected" {% if previous_response and previous_response.category == 'Rejected' %}selected{% endif %}>Rejected</option>
                    <option value="Revise and Resubmit" {% if previous_response and previous_response.category == 'Revise and Resubmit' %}selected{% endif %}>Revise and Resubmit</option>
                </select>
            </div>
            
            <div class="form-group">
                <p class="section-title">Select Files to Include in Response</p>
                <div class="file-list">
                    {% if files %}
                        {% for file in files %}
                        <div class="file-item">
                            <input type="checkbox" name="selected_files" value="{{ file }}" id="file_{{ loop.index }}" {% if previous_files and file in previous_files %}checked{% endif %}>
                            <label for="file_{{ loop.index }}">{{ file }}</label>
                        </div>
                        {% endfor %}
                    {% else %}
                        <p style="color: #666; padding: 8px;">No files found in folder. Please add your response documents first.</p>
                    {% endif %}
                </div>
            </div>
            
            <div class="form-group">
                <label for="notes">Notes / Comments</label>
                <textarea name="notes" id="notes" placeholder="Add any notes for the QC Reviewer and project record...">{{ previous_response.text if previous_response else '' }}</textarea>
            </div>
            
            <button type="submit" class="btn">{% if is_resubmit %}Submit Revision (v{{ version }}){% else %}Submit Review{% endif %}</button>
        </form>
        {% endif %}
    </div>
</body>
</html>
'''

QCR_RESPONSE_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>QC Review - {{ item.type }} {{ item.identifier }}</title>
    ''' + BASE_PAGE_STYLE + '''
    <style>
        .reviewer-response-box {
            background: #f0fdf4;
            border: 1px solid #86efac;
            border-radius: 8px;
            padding: 15px;
            margin: 20px 0;
        }
        .reviewer-response-box h3 {
            margin: 0 0 12px 0;
            color: #166534;
        }
        .action-group {
            background: #fef3c7;
            border: 1px solid #fbbf24;
            border-radius: 8px;
            padding: 15px;
            margin: 20px 0;
        }
        .action-group h3 {
            margin: 0 0 12px 0;
            color: #92400e;
        }
        .radio-group {
            display: flex;
            flex-direction: column;
            gap: 10px;
        }
        .radio-option {
            display: flex;
            align-items: flex-start;
            gap: 10px;
            padding: 10px;
            background: white;
            border-radius: 6px;
            border: 1px solid #e5e7eb;
            cursor: pointer;
        }
        .radio-option:hover {
            border-color: #667eea;
        }
        .radio-option input[type="radio"] {
            margin-top: 3px;
        }
        .radio-option-content {
            flex: 1;
        }
        .radio-option-label {
            font-weight: 600;
            color: #333;
        }
        .radio-option-desc {
            font-size: 12px;
            color: #666;
            margin-top: 4px;
        }
        .response-text-container {
            margin-top: 15px;
        }
        .response-text-label {
            font-weight: 600;
            margin-bottom: 8px;
            color: #333;
        }
        .response-text-readonly {
            background: #f9fafb;
            border: 1px solid #e5e7eb;
            border-radius: 6px;
            padding: 12px;
            white-space: pre-wrap;
            color: #666;
            min-height: 80px;
        }
        #response_text_area {
            display: none;
        }
        .send-back-warning {
            display: none;
            background: #fef2f2;
            border: 1px solid #fecaca;
            border-radius: 8px;
            padding: 12px;
            margin-top: 15px;
            color: #991b1b;
        }
        .version-badge {
            display: inline-block;
            background: #667eea;
            color: white;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            margin-left: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>QC Review <span class="version-badge">v{{ version }}</span></h1>
        <p class="subtitle">{{ item.type }} {{ item.identifier }}</p>
        
        {% if version_history %}
        <div style="background: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px; padding: 12px; margin-bottom: 15px; font-size: 13px;">
            <strong>📋 Reviewer Response Version History:</strong><br>
            Current: <strong>v{{ version }}</strong> ({{ item.reviewer_response_at[:16]|replace('T', ' ') if item.reviewer_response_at else 'N/A' }})<br>
            {% for v in version_history %}
            Previous: v{{ v.version }} ({{ v.submitted_at[:16]|replace('T', ' ') }}) - {{ v.response_category or 'N/A' }}{% if not loop.last %}<br>{% endif %}
            {% endfor %}
        </div>
        {% endif %}
        
        <div class="info-box">
            <div class="info-row">
                <span class="info-label">Title:</span>
                <span class="info-value">{{ item.title or 'N/A' }}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Date Received:</span>
                <span class="info-value">{{ item.date_received or 'N/A' }}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Contractor Due Date:</span>
                <span class="info-value">{{ item.due_date or 'N/A' }}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Priority:</span>
                <span class="info-value">{{ item.priority or 'Normal' }}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Initial Reviewer:</span>
                <span class="info-value">{{ item.reviewer_name or 'N/A' }}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Your Due Date:</span>
                <span class="info-value" style="color: #d97706;">{{ item.qcr_due_date or 'N/A' }}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Folder:</span>
                <span class="info-value">{{ item.folder_link or 'N/A' }}</span>
            </div>
        </div>
        
        <!-- Reviewer's Response Section -->
        <div class="reviewer-response-box">
            <h3>📝 Initial Reviewer's Submitted Response</h3>
            <div class="info-row">
                <span class="info-label">Category:</span>
                <span class="info-value" style="font-weight: 600; color: #166534;">{{ item.reviewer_response_category or 'Not specified' }}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Selected Files:</span>
                <span class="info-value">{{ reviewer_files|join('; ') if reviewer_files else 'None selected' }}</span>
            </div>
            {% if item.reviewer_notes %}
            <div class="info-row" style="flex-direction: column; align-items: flex-start;">
                <span class="info-label">Reviewer's Notes:</span>
                <div class="response-text-readonly" style="margin-top: 8px; width: 100%;">{{ item.reviewer_notes }}</div>
            </div>
            {% endif %}
            {% if item.reviewer_response_text %}
            <div class="info-row" style="flex-direction: column; align-items: flex-start;">
                <span class="info-label">Reviewer's Response Text:</span>
                <div class="response-text-readonly" id="reviewer_original_text" style="margin-top: 8px; width: 100%;">{{ item.reviewer_response_text }}</div>
            </div>
            {% endif %}
        </div>
        
        <form method="POST" id="qcr-form">
            <input type="hidden" name="token" value="{{ token }}">
            
            <!-- QC Action Selection -->
            <div class="action-group">
                <h3>🎯 Your QC Decision</h3>
                <div class="radio-group">
                    <label class="radio-option">
                        <input type="radio" name="qc_action" value="Approve" required>
                        <div class="radio-option-content">
                            <div class="radio-option-label">✅ Approve</div>
                            <div class="radio-option-desc">Accept the reviewer's response as final. No changes needed.</div>
                        </div>
                    </label>
                    <label class="radio-option">
                        <input type="radio" name="qc_action" value="Modify" required>
                        <div class="radio-option-content">
                            <div class="radio-option-label">✏️ Modify</div>
                            <div class="radio-option-desc">Make adjustments to the response. You can tweak or revise the text.</div>
                        </div>
                    </label>
                    <label class="radio-option">
                        <input type="radio" name="qc_action" value="Send Back" required>
                        <div class="radio-option-content">
                            <div class="radio-option-label">↩️ Send Back to Reviewer</div>
                            <div class="radio-option-desc">Return to the reviewer for revisions. You'll specify what needs to change.</div>
                        </div>
                    </label>
                </div>
            </div>
            
            <div class="send-back-warning" id="send-back-warning">
                ⚠️ <strong>Sending Back:</strong> The item will be returned to the Initial Reviewer for revision. They will receive an email with your notes explaining what changes are needed.
            </div>
            
            <!-- Response Text Handling (shown only for Approve/Modify) -->
            <div class="form-group" id="response-mode-group" style="display: none;">
                <label style="font-weight: 600; margin-bottom: 12px; display: block;">📄 Response Text Handling</label>
                <div class="radio-group">
                    <label class="radio-option">
                        <input type="radio" name="response_mode" value="Keep" checked>
                        <div class="radio-option-content">
                            <div class="radio-option-label">Keep as is</div>
                            <div class="radio-option-desc">Use the reviewer's original response text without changes.</div>
                        </div>
                    </label>
                    <label class="radio-option">
                        <input type="radio" name="response_mode" value="Tweak">
                        <div class="radio-option-content">
                            <div class="radio-option-label">Tweak</div>
                            <div class="radio-option-desc">Start with the reviewer's text and make minor edits.</div>
                        </div>
                    </label>
                    <label class="radio-option">
                        <input type="radio" name="response_mode" value="Revise">
                        <div class="radio-option-content">
                            <div class="radio-option-label">Revise</div>
                            <div class="radio-option-desc">Write a completely new response from scratch.</div>
                        </div>
                    </label>
                </div>
                
                <div class="response-text-container" id="response-text-container">
                    <div class="response-text-label">Final Response Text:</div>
                    <div class="response-text-readonly" id="response_text_readonly">{{ item.reviewer_response_text or item.reviewer_notes or '' }}</div>
                    <textarea name="response_text" id="response_text_area" placeholder="Enter your response text...">{{ item.reviewer_response_text or item.reviewer_notes or '' }}</textarea>
                </div>
            </div>
            
            <!-- Final Response Category (shown only for Approve/Modify) -->
            <div class="form-group" id="category-group" style="display: none;">
                <label for="response_category">Final Response Category *</label>
                <select name="response_category" id="response_category">
                    <option value="">-- Select --</option>
                    <option value="Approved" {% if item.reviewer_response_category == 'Approved' %}selected{% endif %}>Approved</option>
                    <option value="Approved as Noted" {% if item.reviewer_response_category == 'Approved as Noted' %}selected{% endif %}>Approved as Noted</option>
                    <option value="For Record Only" {% if item.reviewer_response_category == 'For Record Only' %}selected{% endif %}>For Record Only</option>
                    <option value="Rejected" {% if item.reviewer_response_category == 'Rejected' %}selected{% endif %}>Rejected</option>
                    <option value="Revise and Resubmit" {% if item.reviewer_response_category == 'Revise and Resubmit' %}selected{% endif %}>Revise and Resubmit</option>
                </select>
            </div>
            
            <!-- File Selection (shown only for Approve/Modify) -->
            <div class="form-group" id="files-group" style="display: none;">
                <p class="section-title">Confirm Files to Include in Response</p>
                <p style="color: #666; font-size: 13px; margin-bottom: 8px;">Files selected by the Initial Reviewer are pre-checked. You can adjust as needed.</p>
                <div class="file-list">
                    {% if files %}
                        {% for file in files %}
                        <div class="file-item">
                            <input type="checkbox" name="selected_files" value="{{ file }}" id="file_{{ loop.index }}"
                                {% if file in reviewer_files %}checked{% endif %}>
                            <label for="file_{{ loop.index }}">{{ file }}</label>
                        </div>
                        {% endfor %}
                    {% else %}
                        <p style="color: #666; padding: 8px;">No files found in folder.</p>
                    {% endif %}
                </div>
            </div>
            
            <!-- QC Notes (always shown) -->
            <div class="form-group">
                <label for="qcr_notes">QC Notes / Comments <span id="notes-required-hint" style="display: none; color: #dc2626;">* Required for Send Back</span></label>
                <textarea name="qcr_notes" id="qcr_notes" placeholder="Add any QC notes, conditions, or explain what changes are needed..."></textarea>
            </div>
            
            <button type="submit" class="btn" id="submit-btn">Complete QC Review</button>
        </form>
    </div>
    
    <script>
        const reviewerText = document.getElementById('reviewer_original_text')?.innerText || '{{ (item.reviewer_response_text or item.reviewer_notes or "")|e }}';
        const responseModeGroup = document.getElementById('response-mode-group');
        const categoryGroup = document.getElementById('category-group');
        const filesGroup = document.getElementById('files-group');
        const sendBackWarning = document.getElementById('send-back-warning');
        const notesRequiredHint = document.getElementById('notes-required-hint');
        const responseTextReadonly = document.getElementById('response_text_readonly');
        const responseTextArea = document.getElementById('response_text_area');
        const qcrNotes = document.getElementById('qcr_notes');
        const submitBtn = document.getElementById('submit-btn');
        const categorySelect = document.getElementById('response_category');
        
        // Handle QC Action change
        document.querySelectorAll('input[name="qc_action"]').forEach(radio => {
            radio.addEventListener('change', function() {
                const action = this.value;
                
                if (action === 'Send Back') {
                    responseModeGroup.style.display = 'none';
                    categoryGroup.style.display = 'none';
                    filesGroup.style.display = 'none';
                    sendBackWarning.style.display = 'block';
                    notesRequiredHint.style.display = 'inline';
                    qcrNotes.required = true;
                    categorySelect.required = false;
                    submitBtn.textContent = '↩️ Send Back to Reviewer';
                    submitBtn.style.background = '#f59e0b';
                } else {
                    responseModeGroup.style.display = 'block';
                    categoryGroup.style.display = 'block';
                    filesGroup.style.display = 'block';
                    sendBackWarning.style.display = 'none';
                    notesRequiredHint.style.display = 'none';
                    qcrNotes.required = false;
                    categorySelect.required = true;
                    
                    if (action === 'Approve') {
                        submitBtn.textContent = '✅ Approve & Complete';
                        submitBtn.style.background = '#10b981';
                    } else {
                        submitBtn.textContent = '✏️ Submit Modifications';
                        submitBtn.style.background = 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)';
                    }
                }
            });
        });
        
        // Handle Response Mode change
        document.querySelectorAll('input[name="response_mode"]').forEach(radio => {
            radio.addEventListener('change', function() {
                const mode = this.value;
                
                if (mode === 'Keep') {
                    responseTextReadonly.style.display = 'block';
                    responseTextArea.style.display = 'none';
                    responseTextArea.value = reviewerText;
                } else if (mode === 'Tweak') {
                    responseTextReadonly.style.display = 'none';
                    responseTextArea.style.display = 'block';
                    responseTextArea.value = reviewerText;
                } else if (mode === 'Revise') {
                    responseTextReadonly.style.display = 'none';
                    responseTextArea.style.display = 'block';
                    responseTextArea.value = '';
                    responseTextArea.placeholder = 'Write your new response text here...';
                }
            });
        });
        
        // Form validation
        document.getElementById('qcr-form').addEventListener('submit', function(e) {
            const action = document.querySelector('input[name="qc_action"]:checked')?.value;
            
            if (!action) {
                e.preventDefault();
                alert('Please select a QC decision.');
                return false;
            }
            
            if (action === 'Send Back' && !qcrNotes.value.trim()) {
                e.preventDefault();
                alert('Please provide notes explaining what revisions are needed.');
                return false;
            }
            
            if ((action === 'Approve' || action === 'Modify') && !categorySelect.value) {
                e.preventDefault();
                alert('Please select a final response category.');
                return false;
            }
        });
    </script>
</body>
</html>
'''

# =============================================================================
# DATABASE INITIALIZATION
# =============================================================================

def get_db():
    """Get database connection with row factory."""
    conn = sqlite3.connect(str(DATABASE_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Initialize database tables."""
    conn = get_db()
    cursor = conn.cursor()
    
    # User table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS user (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT,
            display_name TEXT,
            role TEXT DEFAULT 'user' CHECK(role IN ('admin', 'user')),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Item table (RFI/Submittal)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS item (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL CHECK(type IN ('RFI', 'Submittal')),
            bucket TEXT NOT NULL DEFAULT 'ALL',
            identifier TEXT NOT NULL,
            title TEXT,
            source_subject TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            last_email_at TIMESTAMP,
            due_date DATE,
            priority TEXT CHECK(priority IN ('High', 'Medium', 'Low', NULL)),
            status TEXT DEFAULT 'Unassigned' CHECK(status IN ('Unassigned', 'Assigned', 'In Review', 'In QC', 'Ready for Response', 'Closed')),
            assigned_to_user_id INTEGER,
            notes TEXT,
            folder_link TEXT,
            response_category TEXT CHECK(response_category IN ('Approved', 'Approved as Noted', 'For Record Only', 'Rejected', 'Revise and Resubmit', NULL)),
            response_text TEXT,
            response_files TEXT,
            closed_at TIMESTAMP,
            read_by TEXT,
            FOREIGN KEY (assigned_to_user_id) REFERENCES user(id),
            UNIQUE(identifier, bucket)
        )
    ''')
    
    # Comment table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS comment (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            body TEXT NOT NULL,
            FOREIGN KEY (item_id) REFERENCES item(id) ON DELETE CASCADE,
            FOREIGN KEY (user_id) REFERENCES user(id)
        )
    ''')
    
    # Email log table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS email_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id INTEGER,
            message_id TEXT UNIQUE,
            entry_id TEXT,
            subject TEXT,
            body_preview TEXT,
            received_at TIMESTAMP,
            raw_type TEXT,
            processed INTEGER DEFAULT 0,
            FOREIGN KEY (item_id) REFERENCES item(id)
        )
    ''')
    
    # Migration: Add new columns to existing databases
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN response_category TEXT')
    except:
        pass
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN response_text TEXT')
    except:
        pass
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN response_files TEXT')
    except:
        pass
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN closed_at TIMESTAMP')
    except:
        pass
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN read_by TEXT')
    except:
        pass
    # New columns for two-level review system
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN date_received DATE')
    except:
        pass
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN initial_reviewer_id INTEGER REFERENCES user(id)')
    except:
        pass
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN qcr_id INTEGER REFERENCES user(id)')
    except:
        pass
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN initial_reviewer_due_date DATE')
    except:
        pass
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN qcr_due_date DATE')
    except:
        pass
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN is_contractor_window_insufficient INTEGER DEFAULT 0')
    except:
        pass
    
    # Backfill date_received for existing items that don't have it
    cursor.execute('''
        UPDATE item 
        SET date_received = DATE(created_at)
        WHERE date_received IS NULL AND created_at IS NOT NULL
    ''')
    
    # Recalculate review due dates for items that have date_received and due_date but missing calculated dates
    cursor.execute('''
        SELECT id, date_received, due_date, priority 
        FROM item 
        WHERE date_received IS NOT NULL 
        AND due_date IS NOT NULL 
        AND (initial_reviewer_due_date IS NULL OR qcr_due_date IS NULL)
    ''')
    items_to_update = cursor.fetchall()
    for item_id, date_received, due_date, priority in items_to_update:
        try:
            due_dates = calculate_review_due_dates(date_received, due_date, priority)
            cursor.execute('''
                UPDATE item SET 
                    initial_reviewer_due_date = ?,
                    qcr_due_date = ?,
                    is_contractor_window_insufficient = ?
                WHERE id = ?
            ''', (
                due_dates['initial_reviewer_due_date'],
                due_dates['qcr_due_date'],
                1 if due_dates['is_contractor_window_insufficient'] else 0,
                item_id
            ))
        except Exception as e:
            print(f"Could not calculate due dates for item {item_id}: {e}")
    
    # Email workflow columns
    email_workflow_columns = [
        ('email_token_reviewer', 'TEXT'),
        ('email_token_qcr', 'TEXT'),
        ('reviewer_email_sent_at', 'TIMESTAMP'),
        ('reviewer_response_at', 'TIMESTAMP'),
        ('qcr_email_sent_at', 'TIMESTAMP'),
        ('qcr_response_at', 'TIMESTAMP'),
        ('reviewer_response_status', "TEXT DEFAULT 'Not Sent'"),
        ('qcr_response_status', "TEXT DEFAULT 'Not Sent'"),
        # Reviewer response fields
        ('reviewer_notes', 'TEXT'),
        ('reviewer_response_category', 'TEXT'),
        ('reviewer_selected_files', 'TEXT'),
        ('reviewer_response_text', 'TEXT'),
        # QCR response fields
        ('qcr_notes', 'TEXT'),
        ('qcr_response_category', 'TEXT'),
        ('qcr_selected_files', 'TEXT'),
        ('qcr_action', 'TEXT'),  # Approve, Modify, Send Back
        ('qcr_response_mode', 'TEXT'),  # Keep, Tweak, Revise
        ('qcr_response_text', 'TEXT'),
        # Final official response fields
        ('final_response_category', 'TEXT'),
        ('final_response_text', 'TEXT'),
        ('final_response_files', 'TEXT'),
    ]
    for col_name, col_type in email_workflow_columns:
        try:
            cursor.execute(f'ALTER TABLE item ADD COLUMN {col_name} {col_type}')
        except:
            pass
    
    # Add reviewer_response_version column for version tracking
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN reviewer_response_version INTEGER DEFAULT 1')
    except:
        pass
    
    # Add email_entry_id column for opening original email
    try:
        cursor.execute('ALTER TABLE item ADD COLUMN email_entry_id TEXT')
    except:
        pass
    
    # Reviewer response history table for version tracking
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS reviewer_response_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id INTEGER NOT NULL,
            version INTEGER NOT NULL,
            submitted_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            response_category TEXT,
            response_text TEXT,
            response_files TEXT,
            submitted_by_user_id INTEGER,
            FOREIGN KEY (item_id) REFERENCES item(id) ON DELETE CASCADE,
            FOREIGN KEY (submitted_by_user_id) REFERENCES user(id)
        )
    ''')
    
    # Notification table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS notification (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL,
            title TEXT NOT NULL,
            message TEXT,
            item_id INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            read_at TIMESTAMP,
            action_url TEXT,
            action_label TEXT,
            FOREIGN KEY (item_id) REFERENCES item(id) ON DELETE CASCADE
        )
    ''')
    
    # Create default admin user if no users exist
    cursor.execute('SELECT COUNT(*) FROM user')
    if cursor.fetchone()[0] == 0:
        default_password = 'admin123'  # Change this!
        password_hash = bcrypt.hashpw(default_password.encode('utf-8'), bcrypt.gensalt())
        cursor.execute('''
            INSERT INTO user (email, password_hash, display_name, role)
            VALUES (?, ?, ?, ?)
        ''', ('admin@local', password_hash.decode('utf-8'), 'Administrator', 'admin'))
        print(f"Created default admin user: admin@local / {default_password}")
    
    conn.commit()
    conn.close()

# =============================================================================
# NOTIFICATION HELPERS
# =============================================================================

def show_windows_toast(title, message):
    """Show a Windows toast notification in the system tray."""
    if HAS_WINOTIFY:
        try:
            toast = Notification(
                app_id="LEB Tracker",
                title=title,
                msg=message,
                duration="short"
            )
            toast.set_audio(audio.Default, loop=False)
            # Add action to open the app
            toast.add_actions(label="Open App", launch="http://localhost:5000")
            toast.show()
        except Exception as e:
            print(f"Toast notification error: {e}")

def create_notification(notification_type, title, message, item_id=None, action_url=None, action_label=None):
    """Create a new notification and show Windows toast."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO notification (type, title, message, item_id, action_url, action_label)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', (notification_type, title, message, item_id, action_url, action_label))
    conn.commit()
    notification_id = cursor.lastrowid
    conn.close()
    
    # Also show Windows toast notification
    show_windows_toast(title, message)
    
    return notification_id

# =============================================================================
# EMAIL PARSING UTILITIES
# =============================================================================

# Bucket mapping based on subject patterns
BUCKET_PATTERNS = [
    (r'LEB\s*-\s*Turner', 'ACC_TURNER'),
    (r'LEB\s*-\s*Mortenson', 'ACC_MORTENSON'),
    (r'LEB\s*-\s*Faith', 'ACC_FTI'),
    (r'LEB\s*-\s*FTI', 'ACC_FTI'),
    (r'LEB', 'ALL'),  # Default fallback
]

def determine_bucket(subject):
    """Determine the bucket from email subject."""
    for pattern, bucket in BUCKET_PATTERNS:
        if re.search(pattern, subject, re.IGNORECASE):
            return bucket
    return 'ALL'

def parse_item_type(subject):
    """Determine if this is an RFI or Submittal."""
    if re.search(r'\bSubmittal\b', subject, re.IGNORECASE):
        return 'Submittal'
    elif re.search(r'\bRFI\b', subject, re.IGNORECASE):
        return 'RFI'
    return None

def parse_identifier(subject, item_type):
    """Extract the identifier from the subject."""
    if item_type == 'Submittal':
        # Match patterns like "Submittal #13 34 19-2" or "Submittal #123"
        match = re.search(r'Submittal\s*#?([\d\s\-\.]+)', subject, re.IGNORECASE)
        if match:
            return f"Submittal #{match.group(1).strip()}"
    elif item_type == 'RFI':
        # Match patterns like "RFI #123" or "RFI-123"
        match = re.search(r'RFI\s*[#\-]?\s*(\d+)', subject, re.IGNORECASE)
        if match:
            return f"RFI #{match.group(1).strip()}"
    return None

def parse_title(subject, identifier, body=None):
    """Extract a title from the email body (Spec Section) or subject."""
    title = None
    
    # First, try to get Spec Section from body (this is the best title)
    if body:
        # ACC format: "Spec Section    25 00 00 - INTEGRATED AUTOMATION"
        spec_match = re.search(r'Spec\s*Section[:\s\t]+([^\n\r]+)', body, re.IGNORECASE)
        if spec_match:
            title = spec_match.group(1).strip()
            if title:
                return title
    
    # Fallback: extract from subject
    if identifier and subject:
        title = subject
        # Remove common prefixes
        title = re.sub(r'^(Re:\s*)?(Fwd?:\s*)?(Action Required:\s*)?', '', title, flags=re.IGNORECASE)
        # Remove project prefix like "LEB - Turner (NB.TypeF2.0) -"
        title = re.sub(r'LEB\s*-?\s*[\w\s]*\([^)]+\)\s*-?\s*', '', title)
        # Remove the identifier (handle both "Submittal #25 00 00-3" formats)
        title = re.sub(re.escape(identifier), '', title, flags=re.IGNORECASE)
        title = re.sub(r'Submittal\s*#?[\d\s\-\.]+', '', title, flags=re.IGNORECASE)
        title = re.sub(r'RFI\s*#?[\d\s\-\.]+', '', title, flags=re.IGNORECASE)
        # Remove trailing actions
        title = re.sub(r'\s*(was assigned to you|was assigned to your role|needs your review|requires action).*$', '', title, flags=re.IGNORECASE)
        title = title.strip(' -–—:,')
        
    return title if title else None

def parse_due_date(body):
    """Extract due date from email body."""
    if not body:
        return None
    
    # Patterns to look for (ACC uses tab/whitespace separation, not just colons)
    patterns = [
        r'Due\s*Date[:\s\t]+([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})',  # "Due Date    Jan 22, 2026"
        r'Due\s*Date[:\s\t]+(\d{1,2}/\d{1,2}/\d{4})',  # "Due Date    01/22/2026"
        r'Due\s*Date[:\s\t]+(\d{4}-\d{2}-\d{2})',  # "Due Date    2026-01-22"
        r'Due[:\s\t]+([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})',  # "Due: Jan 22, 2026"
        r'Due[:\s\t]+(\d{1,2}/\d{1,2}/\d{4})',
        r'Response\s*Due[:\s\t]+(.+?)(?:\n|$)',
        r'Required\s*[Bb]y[:\s\t]+(.+?)(?:\n|$)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, body, re.IGNORECASE)
        if match:
            date_str = match.group(1).strip()
            # Try to parse the date
            if HAS_DATEUTIL:
                try:
                    parsed_date = date_parser.parse(date_str, fuzzy=True)
                    return parsed_date.strftime('%Y-%m-%d')
                except:
                    pass
            else:
                # Basic date parsing without dateutil
                # Try common formats
                for fmt in ['%m/%d/%Y', '%Y-%m-%d', '%B %d, %Y', '%b %d, %Y', '%b %d %Y']:
                    try:
                        parsed_date = datetime.strptime(date_str, fmt)
                        return parsed_date.strftime('%Y-%m-%d')
                    except:
                        pass
    return None

def parse_priority(body):
    """Extract priority from email body."""
    if not body:
        return None
    
    # ACC uses tab/whitespace separation and may use "Normal" instead of "Medium"
    match = re.search(r'Priority[:\s\t]+(High|Medium|Normal|Low|Urgent|Critical)', body, re.IGNORECASE)
    if match:
        priority = match.group(1).capitalize()
        # Map ACC priority values to our standard values
        priority_map = {
            'Normal': 'Medium',
            'Urgent': 'High',
            'Critical': 'High'
        }
        return priority_map.get(priority, priority)
    return None

def sanitize_folder_name(name):
    """Create a filesystem-safe folder name."""
    # Remove or replace invalid characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '-')
    # Remove multiple dashes
    name = re.sub(r'-+', '-', name)
    # Remove leading/trailing dashes and spaces
    name = name.strip('- ')
    return name

def create_item_folder(item_type, identifier, bucket, title=None):
    """Create a folder for the item and return the path."""
    base_path = Path(CONFIG['base_folder_path'])
    
    # Create bucket subfolder
    bucket_folder = bucket.replace('ACC_', '').title()
    if bucket == 'ALL':
        bucket_folder = 'General'
    
    # Create type subfolder
    type_folder = 'Submittals' if item_type == 'Submittal' else 'RFIs'
    
    # Create item folder name - include title for easier distinction
    clean_id = sanitize_folder_name(identifier)
    folder_id = clean_id.replace(f'{item_type} #', '')
    if title:
        clean_title = sanitize_folder_name(title)[:50]  # Limit title length
        item_folder = f"{item_type} - {folder_id} - {clean_title}"
    else:
        item_folder = f"{item_type} - {folder_id}"
    
    # Full path
    full_path = base_path / bucket_folder / type_folder / item_folder
    
    try:
        full_path.mkdir(parents=True, exist_ok=True)
        return str(full_path)
    except Exception as e:
        print(f"Error creating folder {full_path}: {e}")
        return None

# =============================================================================
# OUTLOOK EMAIL POLLING
# =============================================================================

class EmailPoller:
    """Background email polling service."""
    
    def __init__(self):
        self.running = False
        self.thread = None
        self.last_poll = None
        self.poll_count = 0
        self.error_count = 0
        self.last_error = None
    
    def start(self):
        """Start the polling thread."""
        if not HAS_WIN32COM:
            print("Email polling disabled: pywin32 not available")
            return
        
        self.running = True
        self.thread = threading.Thread(target=self._poll_loop, daemon=True)
        self.thread.start()
        print("Email polling started")
    
    def stop(self):
        """Stop the polling thread."""
        self.running = False
        if self.thread:
            self.thread.join(timeout=5)
        print("Email polling stopped")
    
    def _poll_loop(self):
        """Main polling loop."""
        while self.running:
            try:
                self._poll_emails()
                self.last_poll = datetime.now()
                self.poll_count += 1
            except Exception as e:
                self.error_count += 1
                self.last_error = str(e)
                print(f"Email polling error: {e}")
            
            # Wait for next poll
            interval = CONFIG['poll_interval_minutes'] * 60
            for _ in range(interval):
                if not self.running:
                    break
                time.sleep(1)
    
    def _poll_emails(self):
        """Poll Outlook for new emails."""
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Get the folder (Inbox or custom folder)
            folder_name = CONFIG.get('outlook_folder', 'Inbox')
            if folder_name.lower() == 'inbox':
                folder = namespace.GetDefaultFolder(6)  # 6 = Inbox
            else:
                # Try to find a subfolder
                inbox = namespace.GetDefaultFolder(6)
                try:
                    folder = inbox.Folders[folder_name]
                except:
                    folder = inbox  # Fall back to Inbox
            
            # Get messages
            messages = folder.Items
            messages.Sort("[ReceivedTime]", True)  # Sort by newest first
            
            conn = get_db()
            cursor = conn.cursor()
            
            processed_count = 0
            scanned_count = 0
            for message in messages:
                scanned_count += 1
                if scanned_count > 200:  # Limit to last 200 messages
                    break
                    
                try:
                    # Check if already processed
                    message_id = getattr(message, 'EntryID', None)
                    if not message_id:
                        continue
                    
                    cursor.execute('SELECT id FROM email_log WHERE entry_id = ?', (message_id,))
                    if cursor.fetchone():
                        continue  # Already processed
                    
                    # Check sender - only process emails from ACC
                    sender_email = ''
                    try:
                        sender = getattr(message, 'SenderEmailAddress', '') or ''
                        sender_email = sender.lower()
                        # Sometimes Outlook returns Exchange format, try to get SMTP address
                        if '@' not in sender_email:
                            try:
                                sender_email = message.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
                            except:
                                pass
                    except:
                        pass
                    
                    # Only process emails from Autodesk Construction Cloud
                    if 'no-reply@acc.autodesk.com' not in sender_email and 'acc.autodesk.com' not in sender_email:
                        continue
                    
                    subject = getattr(message, 'Subject', '') or ''
                    body = getattr(message, 'Body', '') or ''
                    received_time = getattr(message, 'ReceivedTime', None)
                    
                    # Check if this is a relevant email
                    if 'LEB' not in subject.upper():
                        continue
                    
                    item_type = parse_item_type(subject)
                    if not item_type:
                        continue
                    
                    # Parse the email
                    identifier = parse_identifier(subject, item_type)
                    if not identifier:
                        continue
                    
                    bucket = determine_bucket(subject)
                    title = parse_title(subject, identifier, body)
                    due_date = parse_due_date(body)
                    priority = parse_priority(body)
                    
                    received_at = None
                    if received_time:
                        try:
                            received_at = datetime(
                                received_time.year, received_time.month, received_time.day,
                                received_time.hour, received_time.minute, received_time.second
                            ).isoformat()
                        except:
                            received_at = datetime.now().isoformat()
                    
                    # Check if item exists
                    cursor.execute('''
                        SELECT id, due_date, priority FROM item 
                        WHERE identifier = ? AND bucket = ?
                    ''', (identifier, bucket))
                    existing = cursor.fetchone()
                    
                    item_id = None
                    if existing:
                        # Update existing item
                        item_id = existing['id']
                        updates = ['last_email_at = ?']
                        params = [received_at]
                        
                        # Only update due_date if currently empty
                        if not existing['due_date'] and due_date:
                            updates.append('due_date = ?')
                            params.append(due_date)
                        
                        # Only update priority if currently empty
                        if not existing['priority'] and priority:
                            updates.append('priority = ?')
                            params.append(priority)
                        
                        params.append(item_id)
                        cursor.execute(f'''
                            UPDATE item SET {', '.join(updates)} WHERE id = ?
                        ''', params)
                    else:
                        # Create new item
                        folder_link = create_item_folder(item_type, identifier, bucket, title)
                        
                        # Extract date_received from email received time
                        if received_at:
                            date_received = received_at[:10]  # Get YYYY-MM-DD part
                        else:
                            date_received = datetime.now().strftime('%Y-%m-%d')
                        
                        # Calculate review due dates if we have enough info
                        initial_reviewer_due = None
                        qcr_due = None
                        is_insufficient = 0
                        
                        if due_date and date_received:
                            due_dates = calculate_review_due_dates(date_received, due_date, priority)
                            initial_reviewer_due = due_dates['initial_reviewer_due_date']
                            qcr_due = due_dates['qcr_due_date']
                            is_insufficient = 1 if due_dates['is_contractor_window_insufficient'] else 0
                        
                        cursor.execute('''
                            INSERT INTO item (type, bucket, identifier, title, source_subject, 
                                            created_at, last_email_at, due_date, priority, folder_link,
                                            date_received, initial_reviewer_due_date, qcr_due_date, 
                                            is_contractor_window_insufficient, email_entry_id)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (item_type, bucket, identifier, title, subject,
                              received_at, received_at, due_date, priority, folder_link,
                              date_received, initial_reviewer_due, qcr_due, is_insufficient, message_id))
                        item_id = cursor.lastrowid
                    
                    # Log the email
                    cursor.execute('''
                        INSERT INTO email_log (item_id, message_id, entry_id, subject, body_preview, 
                                              received_at, raw_type, processed)
                        VALUES (?, ?, ?, ?, ?, ?, ?, 1)
                    ''', (item_id, message_id, message_id, subject, body[:500], 
                          received_at, item_type))
                    
                    processed_count += 1
                    
                except Exception as e:
                    print(f"Error processing message: {e}")
                    continue
            
            conn.commit()
            conn.close()
            
            print(f"Scanned {scanned_count} messages, processed {processed_count} new emails")
                
        finally:
            pythoncom.CoUninitialize()
    
    def get_status(self):
        """Get polling status."""
        return {
            'running': self.running,
            'last_poll': self.last_poll.isoformat() if self.last_poll else None,
            'poll_count': self.poll_count,
            'error_count': self.error_count,
            'last_error': self.last_error,
            'outlook_available': HAS_WIN32COM
        }

# Global email poller instance
email_poller = EmailPoller()

# =============================================================================
# EMAIL SENDING FOR WORKFLOW (Reviewer/QCR Assignment)
# =============================================================================

def generate_token():
    """Generate a secure random token for magic links."""
    return secrets.token_urlsafe(32)

def get_app_host():
    """Get the host URL for the application."""
    # For local development, use localhost
    return f"http://localhost:{CONFIG.get('server_port', 5000)}"

def send_reviewer_assignment_email(item_id):
    """Send assignment email to the Initial Reviewer with magic link."""
    if not HAS_WIN32COM:
        return {'success': False, 'error': 'Outlook not available'}
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Get item with reviewer info
    cursor.execute('''
        SELECT i.*, 
               ir.email as reviewer_email, ir.display_name as reviewer_name,
               qcr.email as qcr_email, qcr.display_name as qcr_name
        FROM item i
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE i.id = ?
    ''', (item_id,))
    item = cursor.fetchone()
    
    if not item:
        conn.close()
        return {'success': False, 'error': 'Item not found'}
    
    if not item['reviewer_email']:
        conn.close()
        return {'success': False, 'error': 'No Initial Reviewer assigned'}
    
    # Generate token if not exists
    token = item['email_token_reviewer']
    if not token:
        token = generate_token()
        cursor.execute('UPDATE item SET email_token_reviewer = ? WHERE id = ?', (token, item_id))
    
    # Build the magic link URL
    respond_url = f"{get_app_host()}/respond/reviewer?token={token}"
    
    # Priority color
    priority_color = '#e67e22' if item['priority'] == 'Medium' else '#c0392b' if item['priority'] == 'High' else '#27ae60'
    
    # Build email content
    subject = f"[LEB] {item['identifier']} – Assigned to You"
    
    # Create clickable folder link
    folder_path = item['folder_link'] or 'Not set'
    if folder_path != 'Not set':
        folder_link_html = f'<a href="file:///{folder_path.replace(chr(92), "/")}" style="color:#0078D4; text-decoration:underline;">{folder_path}</a>'
    else:
        folder_link_html = 'Not set'
    
    html_body = f"""<div style="font-family:Segoe UI, Helvetica, Arial, sans-serif; color:#333; font-size:14px; line-height:1.5;">

    <!-- HEADER -->
    <h2 style="color:#444; margin-bottom:6px;">
        [LEB] {item['identifier']} – Assigned to You
    </h2>

    <p style="margin-top:0; font-size:13px; color:#666;">
        You have been assigned a new review task. Please review the details below.
    </p>

    <!-- ACTION BUTTON - AT TOP -->
    <div style="margin:20px 0; text-align:center;">
        <a href="{respond_url}"
           style="background:linear-gradient(135deg, #0078D4 0%, #106EBE 100%); color:white; padding:14px 32px; 
                  font-size:16px; font-weight:600; text-decoration:none; border-radius:8px; display:inline-block;
                  box-shadow: 0 4px 6px rgba(0,120,212,0.3);">
            📝 Open Response Form
        </a>
    </div>

    <!-- INFO TABLE -->
    <table cellpadding="8" cellspacing="0" width="100%" style="border-collapse:collapse; margin-top:10px;">
        <tr>
            <td colspan="2" style="background:#f2f2f2; font-weight:bold; border:1px solid #ddd;">
                Item Information
            </td>
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Type</td>
            <td style="border:1px solid #ddd;">{item['type']}</td>
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Identifier</td>
            <td style="border:1px solid #ddd;">{item['identifier']}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Title</td>
            <td style="border:1px solid #ddd;">{item['title'] or 'N/A'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Date Received</td>
            <td style="border:1px solid #ddd;">{item['date_received'] or 'N/A'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Contractor Due Date</td>
            <td style="border:1px solid #ddd; color:#c0392b; font-weight:bold;">{item['due_date'] or 'N/A'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Priority</td>
            <td style="border:1px solid #ddd; color:{priority_color}; font-weight:bold;">{item['priority'] or 'Normal'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Your Due Date</td>
            <td style="border:1px solid #ddd; color:#2980b9; font-weight:bold;">{item['initial_reviewer_due_date'] or 'N/A'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">QCR Due Date</td>
            <td style="border:1px solid #ddd; color:#27ae60; font-weight:bold;">{item['qcr_due_date'] or 'N/A'}</td>
        </tr>
    </table>

    <!-- FILE PATH SECTION -->
    <div style="margin-top:18px;">
        <div style="font-weight:bold; margin-bottom:4px;">📁 Designated Folder:</div>
        <div style="padding:10px; border:1px solid #ddd; background:#fafafa; font-family:Consolas, monospace; font-size:12px; border-radius:4px;">
            {folder_link_html}
        </div>
        <p style="font-size:11px; color:#888; margin-top:4px;">Click to open folder location</p>
    </div>

    <!-- INSTRUCTIONS -->
    <div style="margin-top:18px; padding:12px; background:#f7f9fc; border-left:4px solid #0078D4;">
        <strong>In the form, you will:</strong>
        <ul style="margin-top:8px; padding-left:20px;">
            <li>Select a response category (Approved, Approved as Noted, For Record Only, Rejected, Revise/Resubmit)</li>
            <li>Select files from the designated folder to include in the official response</li>
            <li>Add notes/comments for QC and project records</li>
        </ul>
    </div>

    <!-- CC NOTE -->
    <p style="margin-top:16px; font-size:13px; color:#555;">
        <strong>{item['qcr_name'] or 'The QCR'}</strong> has been CC'd so they are aware this review is in progress.
    </p>

    <!-- FOOTER -->
    <p style="margin-top:20px; font-size:12px; color:#777;">
        <em>This message was automatically generated. If you believe you received this by mistake, please contact the project administrator.</em>
    </p>

</div>"""
    
    try:
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = item['reviewer_email']
            if item['qcr_email']:
                mail.CC = item['qcr_email']
            mail.Subject = subject
            mail.HTMLBody = html_body
            mail.Send()
            
            # Update database
            cursor.execute('''
                UPDATE item SET 
                    reviewer_email_sent_at = ?,
                    reviewer_response_status = 'Email Sent',
                    status = 'Assigned'
                WHERE id = ?
            ''', (datetime.now().isoformat(), item_id))
            conn.commit()
            
        finally:
            pythoncom.CoUninitialize()
        
        conn.close()
        return {'success': True, 'message': 'Reviewer assignment email sent'}
        
    except Exception as e:
        conn.close()
        return {'success': False, 'error': str(e)}

def send_qcr_assignment_email(item_id, is_revision=False, version=None):
    """Send assignment email to the QCR with magic link."""
    if not HAS_WIN32COM:
        return {'success': False, 'error': 'Outlook not available'}
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Get item with reviewer info
    cursor.execute('''
        SELECT i.*, 
               ir.email as reviewer_email, ir.display_name as reviewer_name,
               qcr.email as qcr_email, qcr.display_name as qcr_name
        FROM item i
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE i.id = ?
    ''', (item_id,))
    item = cursor.fetchone()
    
    if not item:
        conn.close()
        return {'success': False, 'error': 'Item not found'}
    
    if not item['qcr_email']:
        conn.close()
        return {'success': False, 'error': 'No QCR assigned'}
    
    # Generate token if not exists
    token = item['email_token_qcr']
    if not token:
        token = generate_token()
        cursor.execute('UPDATE item SET email_token_qcr = ? WHERE id = ?', (token, item_id))
    
    # Build the magic link URL
    respond_url = f"{get_app_host()}/respond/qcr?token={token}"
    
    # Get version info
    current_version = version or item['reviewer_response_version'] or 1
    
    # Get version history for display
    cursor.execute('''
        SELECT version, submitted_at 
        FROM reviewer_response_history 
        WHERE item_id = ? 
        ORDER BY version DESC
        LIMIT 5
    ''', (item_id,))
    history = cursor.fetchall()
    version_history_html = ''
    if history:
        history_parts = [f"v{h['version']} ({h['submitted_at'][:16].replace('T', ' ')})" for h in history]
        version_history_html = f"<p style='font-size: 12px; color: #666;'><strong>Previous versions:</strong> {', '.join(history_parts)}</p>"
    
    # Format reviewer response time
    reviewer_response_time = item['reviewer_response_at'] or 'N/A'
    if reviewer_response_time != 'N/A':
        try:
            dt = datetime.fromisoformat(reviewer_response_time.replace('Z', '+00:00'))
            reviewer_response_time = dt.strftime('%Y-%m-%d %H:%M')
        except:
            pass
    
    # Format reviewer selected files
    reviewer_files_display = 'None selected'
    reviewer_files_list = []
    if item['reviewer_selected_files']:
        try:
            reviewer_files_list = json.loads(item['reviewer_selected_files'])
            if reviewer_files_list:
                reviewer_files_display = '<br>'.join([f"• {f}" for f in reviewer_files_list])
        except:
            reviewer_files_display = item['reviewer_selected_files']
    
    # Format reviewer notes
    reviewer_notes_html = (item['reviewer_notes'] or 'No notes provided').replace('\n', '<br>')
    
    # Priority color
    priority_color = '#e67e22' if item['priority'] == 'Medium' else '#c0392b' if item['priority'] == 'High' else '#27ae60'
    
    # Build subject and intro based on whether this is a revision
    if is_revision:
        subject = f"[LEB] {item['type']} {item['identifier']} – Ready for Your Review (v{current_version})"
        intro_text = f"<strong style='color: #f59e0b;'>📝 Revision v{current_version}</strong> - The Initial Reviewer has submitted an updated response after your feedback. Please complete a new QC review."
    else:
        subject = f"[LEB] {item['type']} {item['identifier']} – Ready for Your Review"
        intro_text = "The Initial Reviewer has submitted their response. Please complete the QC review."
    
    # Create clickable folder link for QCR email
    folder_path = item['folder_link'] or 'Not set'
    if folder_path != 'Not set':
        folder_link_html = f'<a href="file:///{folder_path.replace(chr(92), "/")}" style="color:#27ae60; text-decoration:underline;">{folder_path}</a>'
    else:
        folder_link_html = 'Not set'
    
    html_body = f"""<div style="font-family:Segoe UI, Helvetica, Arial, sans-serif; color:#333; font-size:14px; line-height:1.5;">

    <!-- HEADER -->
    <h2 style="color:#444; margin-bottom:6px;">
        [LEB] {item['type']} {item['identifier']} – Ready for Your Review {f'(v{current_version})' if is_revision else ''}
    </h2>

    <p style="margin-top:0; font-size:13px; color:#666;">
        {intro_text}
    </p>

    <!-- ACTION BUTTON - AT TOP -->
    <div style="margin:20px 0; text-align:center;">
        <a href="{respond_url}"
           style="background:linear-gradient(135deg, #27ae60 0%, #1e8449 100%); color:white; padding:14px 32px; 
                  font-size:16px; font-weight:600; text-decoration:none; border-radius:8px; display:inline-block;
                  box-shadow: 0 4px 6px rgba(39,174,96,0.3);">
            ✅ Open QC Review Form
        </a>
    </div>

    <!-- INFO TABLE -->
    <table cellpadding="8" cellspacing="0" width="100%" style="border-collapse:collapse; margin-top:10px;">
        <tr>
            <td colspan="2" style="background:#f2f2f2; font-weight:bold; border:1px solid #ddd;">
                Item Information
            </td>
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Type</td>
            <td style="border:1px solid #ddd;">{item['type']}</td>
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Identifier</td>
            <td style="border:1px solid #ddd;">{item['identifier']}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Title</td>
            <td style="border:1px solid #ddd;">{item['title'] or 'N/A'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Date Received</td>
            <td style="border:1px solid #ddd;">{item['date_received'] or 'N/A'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Contractor Due Date</td>
            <td style="border:1px solid #ddd; color:#c0392b; font-weight:bold;">{item['due_date'] or 'N/A'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Priority</td>
            <td style="border:1px solid #ddd; color:{priority_color}; font-weight:bold;">{item['priority'] or 'Normal'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Initial Reviewer</td>
            <td style="border:1px solid #ddd;">{item['reviewer_name'] or 'N/A'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Your Due Date</td>
            <td style="border:1px solid #ddd; color:#2980b9; font-weight:bold;">{item['qcr_due_date'] or 'N/A'}</td>
        </tr>
    </table>

    <!-- REVIEWER RESPONSE SECTION -->
    <table cellpadding="8" cellspacing="0" width="100%" style="border-collapse:collapse; margin-top:16px;">
        <tr>
            <td colspan="2" style="background:#d5f5e3; font-weight:bold; border:1px solid #82e0aa; color:#1e8449;">
                ✅ Reviewer's Submitted Response (v{current_version})
            </td>
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Category</td>
            <td style="border:1px solid #ddd; color:#1e8449; font-weight:bold;">{item['reviewer_response_category'] or 'Not specified'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold; vertical-align:top;">Selected Files</td>
            <td style="border:1px solid #ddd;">{reviewer_files_display}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold; vertical-align:top;">Notes</td>
            <td style="border:1px solid #ddd;">{reviewer_notes_html}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold;">Responded At</td>
            <td style="border:1px solid #ddd;">{reviewer_response_time}</td>
        </tr>
    </table>

    {version_history_html}

    <!-- FILE PATH SECTION -->
    <div style="margin-top:18px;">
        <div style="font-weight:bold; margin-bottom:4px;">📁 Designated Folder:</div>
        <div style="padding:10px; border:1px solid #ddd; background:#fafafa; font-family:Consolas, monospace; font-size:12px; border-radius:4px;">
            {folder_link_html}
        </div>
        <p style="font-size:11px; color:#888; margin-top:4px;">Click to open folder location</p>
    </div>

    <!-- INSTRUCTIONS -->
    <div style="margin-top:18px; padding:12px; background:#f7f9fc; border-left:4px solid #27ae60;">
        <strong>In the QC form, you will:</strong>
        <ul style="margin-top:8px; padding-left:20px;">
            <li>Review the Initial Reviewer's response and files</li>
            <li>Choose an action: <strong>Approve</strong>, <strong>Modify</strong>, or <strong>Send Back</strong></li>
            <li>Decide how to handle the response text: Keep as is, Tweak, or Revise</li>
            <li>Confirm or adjust file selection for the official response</li>
            <li>Add any QC notes or conditions</li>
        </ul>
    </div>

    <!-- CC NOTE -->
    <p style="margin-top:16px; font-size:13px; color:#555;">
        <strong>{item['reviewer_name'] or 'The Initial Reviewer'}</strong> has been CC'd on this email for visibility.
    </p>

    <!-- FOOTER -->
    <p style="margin-top:20px; font-size:12px; color:#777;">
        <em>This message was automatically generated. If you believe you received this by mistake, please contact the project administrator.</em>
    </p>

</div>"""
    
    try:
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = item['qcr_email']
            if item['reviewer_email']:
                mail.CC = item['reviewer_email']
            mail.Subject = subject
            mail.HTMLBody = html_body
            mail.Send()
            
            # Update database
            cursor.execute('''
                UPDATE item SET 
                    qcr_email_sent_at = ?,
                    qcr_response_status = 'Email Sent',
                    status = 'In QC'
                WHERE id = ?
            ''', (datetime.now().isoformat(), item_id))
            conn.commit()
            
        finally:
            pythoncom.CoUninitialize()
        
        conn.close()
        return {'success': True, 'message': 'QCR assignment email sent'}
        
    except Exception as e:
        conn.close()
        return {'success': False, 'error': str(e)}


def send_qcr_version_update_email(item_id, version):
    """Send an email to QCR notifying them of an updated reviewer response (before QCR has responded)."""
    if not HAS_WIN32COM:
        return {'success': False, 'error': 'Outlook not available'}
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Get item with reviewer info
    cursor.execute('''
        SELECT i.*, 
               ir.email as reviewer_email, ir.display_name as reviewer_name,
               qcr.email as qcr_email, qcr.display_name as qcr_name
        FROM item i
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE i.id = ?
    ''', (item_id,))
    item = cursor.fetchone()
    
    if not item:
        conn.close()
        return {'success': False, 'error': 'Item not found'}
    
    if not item['qcr_email']:
        conn.close()
        return {'success': False, 'error': 'No QCR assigned'}
    
    # Get existing QCR token
    token = item['email_token_qcr']
    if not token:
        token = generate_token()
        cursor.execute('UPDATE item SET email_token_qcr = ? WHERE id = ?', (token, item_id))
        conn.commit()
    
    respond_url = f"{get_app_host()}/respond/qcr?token={token}"
    
    # Get version history
    cursor.execute('''
        SELECT version, submitted_at 
        FROM reviewer_response_history 
        WHERE item_id = ? 
        ORDER BY version DESC
        LIMIT 5
    ''', (item_id,))
    history = cursor.fetchall()
    version_history_html = ''
    if history:
        history_parts = [f"v{h['version']} ({h['submitted_at'][:16].replace('T', ' ')})" for h in history]
        version_history_html = f"<p><strong>Previous versions:</strong> {', '.join(history_parts)}</p>"
    
    conn.close()
    
    # Format reviewer notes
    reviewer_notes_html = (item['reviewer_notes'] or 'No notes provided').replace('\n', '<br>')
    
    # Format files
    reviewer_files_display = 'None selected'
    if item['reviewer_selected_files']:
        try:
            files = json.loads(item['reviewer_selected_files'])
            if files:
                reviewer_files_display = '<br>'.join([f"• {f}" for f in files])
        except:
            pass
    
    subject = f"[LEB] {item['type']} {item['identifier']} – Reviewer response updated (v{version})"
    
    html_body = f"""<div style="font-family:Segoe UI, Helvetica, Arial, sans-serif; color:#333; font-size:14px; line-height:1.5;">

    <h2 style="color:#444; margin-bottom:6px;">
        [LEB] {item['type']} {item['identifier']} – Reviewer Response Updated
    </h2>

    <div style="background: #fef3c7; border: 1px solid #f59e0b; border-radius: 8px; padding: 15px; margin: 15px 0;">
        <p style="margin: 0; color: #92400e;">
            <strong>📝 Update Notice:</strong> The Initial Reviewer has updated their response to <strong>version {version}</strong>.
        </p>
    </div>

    <table cellpadding="8" cellspacing="0" width="100%" style="border-collapse:collapse; margin-top:10px;">
        <tr>
            <td colspan="2" style="background:#d5f5e3; font-weight:bold; border:1px solid #82e0aa; color:#1e8449;">
                ✅ Updated Reviewer Response (v{version})
            </td>
        </tr>

        <tr>
            <td style="width:160px; border:1px solid #ddd; font-weight:bold;">Category</td>
            <td style="border:1px solid #ddd; color:#1e8449; font-weight:bold;">{item['reviewer_response_category'] or 'Not specified'}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold; vertical-align:top;">Selected Files</td>
            <td style="border:1px solid #ddd;">{reviewer_files_display}</td>
        </tr>

        <tr>
            <td style="border:1px solid #ddd; font-weight:bold; vertical-align:top;">Notes</td>
            <td style="border:1px solid #ddd;">{reviewer_notes_html}</td>
        </tr>
    </table>

    {version_history_html}

    <p style="margin-top: 15px;">The QC Review form will always show the <strong>latest version</strong> of the reviewer's response.</p>

    <div style="margin-top:22px; text-align:left;">
        <a href="{respond_url}"
           style="background:#27ae60; color:white; padding:10px 18px; 
                  font-size:15px; text-decoration:none; border-radius:4px; display:inline-block;">
            Open QC Review Form
        </a>
    </div>

    <p style="margin-top:20px; font-size:12px; color:#777;">
        <em>This message was automatically generated. If you believe you received this by mistake, please contact the project administrator.</em>
    </p>

</div>"""
    
    try:
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = item['qcr_email']
            if item['reviewer_email']:
                mail.CC = item['reviewer_email']
            mail.Subject = subject
            mail.HTMLBody = html_body
            mail.Send()
        finally:
            pythoncom.CoUninitialize()
        
        return {'success': True, 'message': 'Version update email sent'}
        
    except Exception as e:
        return {'success': False, 'error': str(e)}


def send_reviewer_notification_email(item_id, qc_action, qcr_notes, final_category=None, final_text=None):
    """Send notification email to reviewer based on QCR action."""
    if not HAS_WIN32COM:
        return {'success': False, 'error': 'Outlook not available'}
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Get item with reviewer info
    cursor.execute('''
        SELECT i.*, 
               ir.email as reviewer_email, ir.display_name as reviewer_name,
               qcr.email as qcr_email, qcr.display_name as qcr_name
        FROM item i
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE i.id = ?
    ''', (item_id,))
    item = cursor.fetchone()
    
    if not item:
        conn.close()
        return {'success': False, 'error': 'Item not found'}
    
    if not item['reviewer_email']:
        conn.close()
        return {'success': False, 'error': 'No reviewer email'}
    
    # Get version info
    version = item['reviewer_response_version'] or 1
    
    # Build email based on action
    if qc_action == 'Approve':
        subject = f"[LEB] {item['type']} {item['identifier']} – Your response (v{version}) was approved"
        body_intro = f"""<p>Good news! Your response <strong>(v{version})</strong> for the following item has been <strong style="color: #059669;">approved</strong> by QC.</p>"""
        action_color = '#059669'
        action_icon = '✅'
    elif qc_action == 'Modify':
        subject = f"[LEB] {item['type']} {item['identifier']} – Your response (v{version}) was modified by QC"
        body_intro = f"""<p>Your response <strong>(v{version})</strong> for the following item has been <strong style="color: #2563eb;">modified</strong> by QC and is now finalized.</p>"""
        action_color = '#2563eb'
        action_icon = '✏️'
    else:  # Send Back
        subject = f"[LEB] {item['type']} {item['identifier']} – Revisions requested on v{version}"
        body_intro = f"""<p>Your response <strong>(v{version})</strong> for the following item has been <strong style="color: #dc2626;">returned</strong> for revision by QC.</p>"""
        action_color = '#dc2626'
        action_icon = '↩️'
    
    # Format QCR notes
    qcr_notes_html = (qcr_notes or 'No notes provided').replace('\n', '<br>')
    
    # Build the response comparison section for Modify action
    response_comparison = ''
    if qc_action == 'Modify' and final_text:
        original_text = item['reviewer_response_text'] or item['reviewer_notes'] or 'No original text'
        response_comparison = f"""
        <div style="margin: 20px 0;">
            <h4 style="margin: 0 0 10px 0;">Your Original Response (v{version}):</h4>
            <div style="background: #fef2f2; border: 1px solid #fecaca; border-radius: 6px; padding: 12px; margin-bottom: 15px;">
                {original_text.replace(chr(10), '<br>')}
            </div>
            
            <h4 style="margin: 0 0 10px 0;">Final Response (QC Modified):</h4>
            <div style="background: #f0fdf4; border: 1px solid #86efac; border-radius: 6px; padding: 12px;">
                {final_text.replace(chr(10), '<br>')}
            </div>
        </div>
        """
    
    # Build revision link for Send Back action
    revision_link = ''
    if qc_action == 'Send Back':
        # Generate new token for revision
        new_token = generate_token()
        cursor.execute('UPDATE item SET email_token_reviewer = ? WHERE id = ?', (new_token, item_id))
        conn.commit()
        
        respond_url = f"{get_app_host()}/respond/reviewer?token={new_token}"
        revision_link = f"""
        <p style="margin: 20px 0;">
            <a href="{respond_url}" style="display: inline-block; background: #dc2626; color: white; padding: 12px 24px; text-decoration: none; border-radius: 8px; font-weight: bold;">📝 Revise Your Response (Submit v{version + 1})</a>
        </p>
        """
    
    conn.close()
    
    html_body = f"""<html><body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Hello {item['reviewer_name'] or 'Reviewer'},</p>

{body_intro}

<table style="border-collapse: collapse; margin: 15px 0;">
<tr><td style="padding: 5px 15px 5px 0; font-weight: bold;">Type:</td><td>{item['type']}</td></tr>
<tr><td style="padding: 5px 15px 5px 0; font-weight: bold;">Identifier:</td><td>{item['identifier']}</td></tr>
<tr><td style="padding: 5px 15px 5px 0; font-weight: bold;">Title:</td><td>{item['title'] or 'N/A'}</td></tr>
<tr><td style="padding: 5px 15px 5px 0; font-weight: bold;">Due Date:</td><td>{item['due_date'] or 'N/A'}</td></tr>
<tr><td style="padding: 5px 15px 5px 0; font-weight: bold;">Priority:</td><td>{item['priority'] or 'Normal'}</td></tr>
<tr><td style="padding: 5px 15px 5px 0; font-weight: bold;">Response Version:</td><td>v{version}</td></tr>
</table>

<div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 15px; margin: 20px 0;">
<h3 style="margin: 0 0 10px 0; color: {action_color};">{action_icon} QC Decision: {qc_action}</h3>
<p style="font-size: 12px; color: #666; margin-bottom: 10px;">Feedback on your v{version} response</p>
{f'<p><strong>Final Category:</strong> {final_category}</p>' if final_category else ''}
<p><strong>QC Notes:</strong></p>
<div style="background: white; border-radius: 4px; padding: 10px; margin-top: 8px;">
{qcr_notes_html}
</div>
</div>

{response_comparison}
{revision_link}

<p>Thank you.</p>
</body></html>"""
    
    try:
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = item['reviewer_email']
            if item['qcr_email']:
                mail.CC = item['qcr_email']
            mail.Subject = subject
            mail.HTMLBody = html_body
            mail.Send()
        finally:
            pythoncom.CoUninitialize()
        
        return {'success': True, 'message': 'Reviewer notification sent'}
        
    except Exception as e:
        return {'success': False, 'error': str(e)}

# =============================================================================
# AUTHENTICATION DECORATOR
# =============================================================================

def login_required(f):
    """Decorator to require login for API endpoints."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'error': 'Authentication required'}), 401
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    """Decorator to require admin role."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'error': 'Authentication required'}), 401
        if session.get('role') != 'admin':
            return jsonify({'error': 'Admin access required'}), 403
        return f(*args, **kwargs)
    return decorated_function

# Global error handler for API routes - return JSON instead of HTML
@app.errorhandler(500)
def internal_error(error):
    return jsonify({'error': 'Internal server error', 'details': str(error)}), 500

@app.errorhandler(404)
def not_found_error(error):
    # Only return JSON for API routes
    if request.path.startswith('/api/'):
        return jsonify({'error': 'Not found'}), 404
    return send_from_directory('static', 'index.html')

# =============================================================================
# API ROUTES - AUTHENTICATION
# =============================================================================

@app.route('/api/auth/login', methods=['POST'])
def api_login():
    """Login endpoint."""
    data = request.get_json()
    email = data.get('email', '').strip()
    password = data.get('password', '')
    
    if not email or not password:
        return jsonify({'error': 'Email and password required'}), 400
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM user WHERE email = ?', (email,))
    user = cursor.fetchone()
    conn.close()
    
    if not user:
        return jsonify({'error': 'Invalid credentials'}), 401
    
    if not bcrypt.checkpw(password.encode('utf-8'), user['password_hash'].encode('utf-8')):
        return jsonify({'error': 'Invalid credentials'}), 401
    
    session.permanent = True
    session['user_id'] = user['id']
    session['email'] = user['email']
    session['display_name'] = user['display_name']
    session['role'] = user['role']
    
    return jsonify({
        'id': user['id'],
        'email': user['email'],
        'display_name': user['display_name'],
        'role': user['role']
    })

@app.route('/api/auth/logout', methods=['POST'])
def api_logout():
    """Logout endpoint."""
    session.clear()
    return jsonify({'success': True})

@app.route('/api/auth/me', methods=['GET'])
def api_me():
    """Get current user info."""
    if 'user_id' not in session:
        return jsonify({'error': 'Not authenticated'}), 401
    
    return jsonify({
        'id': session.get('user_id'),
        'email': session.get('email'),
        'display_name': session.get('display_name'),
        'role': session.get('role')
    })

# =============================================================================
# API ROUTES - USERS
# =============================================================================

@app.route('/api/users', methods=['GET'])
@login_required
def api_get_users():
    """Get all users."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT id, email, display_name, role, created_at FROM user')
    users = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify(users)

@app.route('/api/users', methods=['POST'])
@admin_required
def api_create_user():
    """Create a new user (admin only). Password is optional for team members."""
    data = request.get_json()
    email = (data.get('email') or '').strip()
    password = (data.get('password') or '').strip()
    display_name = (data.get('display_name') or '').strip()
    role = data.get('role', 'user')
    
    if not email:
        return jsonify({'error': 'Email is required'}), 400
    
    if role not in ['admin', 'user']:
        role = 'user'
    
    # Password is optional - if not provided, user cannot log in but can be assigned to items
    password_hash = None
    if password:
        password_hash = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute('''
            INSERT INTO user (email, password_hash, display_name, role)
            VALUES (?, ?, ?, ?)
        ''', (email, password_hash, display_name or email, role))
        conn.commit()
        user_id = cursor.lastrowid
    except sqlite3.IntegrityError as e:
        conn.close()
        error_msg = str(e).lower()
        if 'email' in error_msg or 'unique' in error_msg:
            return jsonify({'error': 'Email already exists'}), 400
        return jsonify({'error': f'Database error: {e}'}), 400
    except Exception as e:
        conn.close()
        print(f"User creation error: {e}")
        return jsonify({'error': f'Error creating user: {e}'}), 500
    
    conn.close()
    return jsonify({'id': user_id, 'email': email, 'display_name': display_name or email, 'role': role})

@app.route('/api/outlook/contacts', methods=['GET'])
@login_required
def api_search_outlook_contacts():
    """Search Outlook contacts/address book for autocomplete."""
    query = request.args.get('q', '').strip()
    
    if not query or len(query) < 2:
        return jsonify([])
    
    if not HAS_WIN32COM:
        return jsonify({'error': 'Outlook not available'}), 503
    
    contacts = []
    
    try:
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Search using Outlook's address book resolution
            # This searches the Global Address List (GAL), contacts, and recent recipients
            recipient = namespace.CreateRecipient(query)
            recipient.Resolve()
            
            # If exact match found
            if recipient.Resolved:
                try:
                    addr_entry = recipient.AddressEntry
                    email = ''
                    display_name = addr_entry.Name or ''
                    
                    # Try to get SMTP address
                    if addr_entry.Type == 'EX':
                        # Exchange user - get SMTP address
                        try:
                            exch_user = addr_entry.GetExchangeUser()
                            if exch_user:
                                email = exch_user.PrimarySmtpAddress
                                display_name = exch_user.Name or display_name
                        except:
                            pass
                    else:
                        email = addr_entry.Address or ''
                    
                    if email:
                        contacts.append({
                            'email': email,
                            'display_name': display_name,
                            'type': 'resolved'
                        })
                except Exception as e:
                    print(f"Error getting resolved contact: {e}")
            
            # Also search the Global Address List for partial matches
            try:
                gal = namespace.AddressLists.Item("Global Address List")
                if gal:
                    query_lower = query.lower()
                    count = 0
                    for entry in gal.AddressEntries:
                        if count >= 10:  # Limit results
                            break
                        try:
                            name = entry.Name or ''
                            if query_lower in name.lower():
                                email = ''
                                if entry.Type == 'EX':
                                    try:
                                        exch_user = entry.GetExchangeUser()
                                        if exch_user:
                                            email = exch_user.PrimarySmtpAddress
                                    except:
                                        pass
                                else:
                                    email = entry.Address or ''
                                
                                if email and not any(c['email'].lower() == email.lower() for c in contacts):
                                    contacts.append({
                                        'email': email,
                                        'display_name': name,
                                        'type': 'gal'
                                    })
                                    count += 1
                        except:
                            continue
            except Exception as e:
                print(f"GAL search error: {e}")
            
            # Also search personal contacts folder
            try:
                contacts_folder = namespace.GetDefaultFolder(10)  # olFolderContacts
                query_lower = query.lower()
                count = 0
                for item in contacts_folder.Items:
                    if count >= 5:
                        break
                    try:
                        full_name = getattr(item, 'FullName', '') or ''
                        email = getattr(item, 'Email1Address', '') or ''
                        
                        if query_lower in full_name.lower() or query_lower in email.lower():
                            if email and not any(c['email'].lower() == email.lower() for c in contacts):
                                contacts.append({
                                    'email': email,
                                    'display_name': full_name or email,
                                    'type': 'contact'
                                })
                                count += 1
                    except:
                        continue
            except Exception as e:
                print(f"Contacts folder search error: {e}")
                
        finally:
            pythoncom.CoUninitialize()
            
    except Exception as e:
        print(f"Outlook contact search error: {e}")
        return jsonify({'error': str(e)}), 500
    
    return jsonify(contacts[:15])  # Limit to 15 results

# =============================================================================
# API ROUTES - ITEMS
# =============================================================================

@app.route('/api/items', methods=['GET'])
@login_required
def api_get_items():
    """Get items with optional filtering."""
    bucket = request.args.get('bucket')
    item_type = request.args.get('type')
    status = request.args.get('status')
    assigned_to = request.args.get('assigned_to')
    
    conn = get_db()
    cursor = conn.cursor()
    
    query = '''
        SELECT i.*, 
               u.display_name as assigned_to_name,
               ir.display_name as initial_reviewer_name,
               qcr.display_name as qcr_name
        FROM item i
        LEFT JOIN user u ON i.assigned_to_user_id = u.id
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE 1=1
    '''
    params = []
    
    if bucket and bucket != 'ALL':
        query += ' AND i.bucket = ?'
        params.append(bucket)
    
    if item_type:
        query += ' AND i.type = ?'
        params.append(item_type)
    
    if status:
        query += ' AND i.status = ?'
        params.append(status)
    
    if assigned_to:
        query += ' AND i.assigned_to_user_id = ?'
        params.append(assigned_to)
    
    # Handle show_closed filter
    show_closed = request.args.get('show_closed', 'true').lower()
    if show_closed != 'true':
        query += " AND i.closed_at IS NULL"
    
    query += ' ORDER BY i.date_received DESC, i.created_at DESC'
    
    cursor.execute(query, params)
    items = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    return jsonify(items)

@app.route('/api/item/<int:item_id>', methods=['GET'])
@login_required
def api_get_item(item_id):
    """Get a single item with all reviewer info."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT i.*, 
               u.display_name as assigned_to_name,
               ir.display_name as initial_reviewer_name,
               qcr.display_name as qcr_name
        FROM item i
        LEFT JOIN user u ON i.assigned_to_user_id = u.id
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE i.id = ?
    ''', (item_id,))
    item = cursor.fetchone()
    conn.close()
    
    if not item:
        return jsonify({'error': 'Item not found'}), 404
    
    item_dict = dict(item)
    
    # Add due date status colors
    item_dict['initial_reviewer_due_status'] = get_due_date_status(item_dict.get('initial_reviewer_due_date'))
    item_dict['qcr_due_status'] = get_due_date_status(item_dict.get('qcr_due_date'))
    item_dict['contractor_due_status'] = get_due_date_status(item_dict.get('due_date'))
    
    return jsonify(item_dict)

@app.route('/api/item/<int:item_id>', methods=['POST', 'PUT'])
@login_required
def api_update_item(item_id):
    """Update an item with validation for reviewers."""
    data = request.get_json()
    
    # Validate that initial_reviewer_id and qcr_id are not the same
    initial_reviewer_id = data.get('initial_reviewer_id')
    qcr_id = data.get('qcr_id')
    
    if initial_reviewer_id and qcr_id and str(initial_reviewer_id) == str(qcr_id):
        return jsonify({'error': 'Initial Reviewer and QCR must be different users'}), 400
    
    # Check if user is manually setting due dates
    manual_initial_due = 'initial_reviewer_due_date' in data
    manual_qcr_due = 'qcr_due_date' in data
    
    # Fields that can be updated
    allowed_fields = [
        'title', 'due_date', 'priority', 'status', 'assigned_to_user_id', 'notes', 'folder_link',
        'initial_reviewer_id', 'qcr_id', 'date_received',
        'initial_reviewer_due_date', 'qcr_due_date'
    ]
    updates = []
    params = []
    
    for field in allowed_fields:
        if field in data:
            updates.append(f'{field} = ?')
            value = data[field] if data[field] != '' else None
            params.append(value)
    
    if not updates:
        return jsonify({'error': 'No fields to update'}), 400
    
    params.append(item_id)
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Get current item data to check for recalculation needs
    cursor.execute('SELECT date_received, due_date, priority FROM item WHERE id = ?', (item_id,))
    current = cursor.fetchone()
    
    cursor.execute(f'''
        UPDATE item SET {', '.join(updates)} WHERE id = ?
    ''', params)
    
    # Only recalculate due dates if priority, due_date, or date_received changed
    # AND the user didn't manually provide the due dates
    recalc_fields = ['priority', 'due_date', 'date_received']
    needs_recalc = any(field in data for field in recalc_fields) and not manual_initial_due and not manual_qcr_due
    
    if needs_recalc:
        # Get updated values
        cursor.execute('SELECT date_received, due_date, priority FROM item WHERE id = ?', (item_id,))
        updated = cursor.fetchone()
        
        if updated['date_received'] and updated['due_date']:
            due_dates = calculate_review_due_dates(
                updated['date_received'],
                updated['due_date'],
                updated['priority']
            )
            
            cursor.execute('''
                UPDATE item SET 
                    initial_reviewer_due_date = ?,
                    qcr_due_date = ?,
                    is_contractor_window_insufficient = ?
                WHERE id = ?
            ''', (
                due_dates['initial_reviewer_due_date'],
                due_dates['qcr_due_date'],
                1 if due_dates['is_contractor_window_insufficient'] else 0,
                item_id
            ))
    
    conn.commit()
    
    # Return updated item with all joins
    cursor.execute('''
        SELECT i.*, 
               u.display_name as assigned_to_name,
               ir.display_name as initial_reviewer_name,
               qcr.display_name as qcr_name
        FROM item i
        LEFT JOIN user u ON i.assigned_to_user_id = u.id
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE i.id = ?
    ''', (item_id,))
    item = cursor.fetchone()
    conn.close()
    
    item_dict = dict(item)
    item_dict['initial_reviewer_due_status'] = get_due_date_status(item_dict.get('initial_reviewer_due_date'))
    item_dict['qcr_due_status'] = get_due_date_status(item_dict.get('qcr_due_date'))
    item_dict['contractor_due_status'] = get_due_date_status(item_dict.get('due_date'))
    
    return jsonify(item_dict)

@app.route('/api/items', methods=['POST'])
@login_required
def api_create_item():
    """Manually create a new item."""
    data = request.get_json()
    
    item_type = data.get('type')
    identifier = data.get('identifier', '').strip()
    bucket = data.get('bucket', 'ALL')
    
    if not item_type or item_type not in ['RFI', 'Submittal']:
        return jsonify({'error': 'Valid type (RFI or Submittal) required'}), 400
    
    if not identifier:
        return jsonify({'error': 'Identifier required'}), 400
    
    # Create folder
    title = data.get('title')
    folder_link = create_item_folder(item_type, identifier, bucket, title)
    
    conn = get_db()
    cursor = conn.cursor()
    
    try:
        cursor.execute('''
            INSERT INTO item (type, bucket, identifier, title, due_date, priority, status, folder_link)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            item_type, bucket, identifier,
            data.get('title'), data.get('due_date'), data.get('priority'),
            data.get('status', 'Unassigned'), folder_link
        ))
        conn.commit()
        item_id = cursor.lastrowid
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'error': 'Item with this identifier and bucket already exists'}), 400
    
    cursor.execute('SELECT * FROM item WHERE id = ?', (item_id,))
    item = dict(cursor.fetchone())
    conn.close()
    
    return jsonify(item), 201

# =============================================================================
# API ROUTES - INBOX
# =============================================================================

@app.route('/api/inbox', methods=['GET'])
@login_required
def api_get_inbox():
    """Get items where user is initial reviewer, QCR, or assigned, sorted by relevant due date."""
    user_id = session['user_id']
    show_closed = request.args.get('show_closed', 'false').lower() == 'true'
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Get items where user is initial reviewer, QCR, or assigned to
    query = '''
        SELECT i.*, 
               u.display_name as assigned_to_name,
               ir.display_name as initial_reviewer_name,
               qcr.display_name as qcr_name,
               CASE 
                   WHEN i.initial_reviewer_id = ? THEN 'initial_reviewer'
                   WHEN i.qcr_id = ? THEN 'qcr'
                   ELSE 'assigned'
               END as user_role
        FROM item i
        LEFT JOIN user u ON i.assigned_to_user_id = u.id
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE (i.initial_reviewer_id = ? OR i.qcr_id = ? OR i.assigned_to_user_id = ?)
    '''
    params = [user_id, user_id, user_id, user_id, user_id]
    
    if not show_closed:
        query += " AND i.closed_at IS NULL"
    
    # Sort by role-specific due date, then priority, then date_received
    query += '''
        ORDER BY 
            CASE 
                WHEN i.initial_reviewer_id = ? THEN COALESCE(i.initial_reviewer_due_date, i.due_date)
                WHEN i.qcr_id = ? THEN COALESCE(i.qcr_due_date, i.due_date)
                ELSE i.due_date
            END ASC NULLS LAST,
            CASE i.priority 
                WHEN 'High' THEN 1 
                WHEN 'Medium' THEN 2 
                WHEN 'Low' THEN 3 
                ELSE 4 
            END,
            i.date_received ASC
    '''
    params.extend([user_id, user_id])
    
    cursor.execute(query, params)
    items = [dict(row) for row in cursor.fetchall()]
    
    # Check read status and add due date status for each item
    user_id_str = str(user_id)
    for item in items:
        read_by = item.get('read_by') or ''
        item['is_unread'] = user_id_str not in read_by.split(',')
        
        # Add role-specific due date status
        role = item.get('user_role')
        if role == 'initial_reviewer':
            item['my_due_date'] = item.get('initial_reviewer_due_date')
            item['my_due_status'] = get_due_date_status(item.get('initial_reviewer_due_date'))
        elif role == 'qcr':
            item['my_due_date'] = item.get('qcr_due_date')
            item['my_due_status'] = get_due_date_status(item.get('qcr_due_date'))
        else:
            item['my_due_date'] = item.get('due_date')
            item['my_due_status'] = get_due_date_status(item.get('due_date'))
    
    conn.close()
    return jsonify(items)

@app.route('/api/item/<int:item_id>/mark-read', methods=['POST'])
@login_required
def api_mark_read(item_id):
    """Mark an item as read by the current user."""
    user_id = str(session['user_id'])
    
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT read_by FROM item WHERE id = ?', (item_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Item not found'}), 404
    
    read_by = row['read_by'] or ''
    read_list = [x for x in read_by.split(',') if x]
    
    if user_id not in read_list:
        read_list.append(user_id)
        cursor.execute('UPDATE item SET read_by = ? WHERE id = ?', (','.join(read_list), item_id))
        conn.commit()
    
    conn.close()
    return jsonify({'success': True})

# =============================================================================
# API ROUTES - RESPONSE & CLOSEOUT
# =============================================================================

@app.route('/api/item/<int:item_id>/response', methods=['POST'])
@login_required
def api_save_response(item_id):
    """Save response details for an item."""
    data = request.get_json()
    user_id = session['user_id']
    user_role = session.get('role')
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Check if user can edit (admin or assigned to item)
    cursor.execute('SELECT assigned_to_user_id, status FROM item WHERE id = ?', (item_id,))
    item = cursor.fetchone()
    
    if not item:
        conn.close()
        return jsonify({'error': 'Item not found'}), 404
    
    if item['status'] == 'Closed':
        conn.close()
        return jsonify({'error': 'Cannot modify closed item'}), 400
    
    if user_role != 'admin' and item['assigned_to_user_id'] != user_id:
        conn.close()
        return jsonify({'error': 'Not authorized to edit this item'}), 403
    
    # Update response fields
    updates = []
    params = []
    
    if 'response_category' in data:
        updates.append('response_category = ?')
        params.append(data['response_category'] or None)
    
    if 'response_text' in data:
        updates.append('response_text = ?')
        params.append(data['response_text'] or None)
    
    if 'response_files' in data:
        import json as json_module
        updates.append('response_files = ?')
        params.append(json_module.dumps(data['response_files']) if data['response_files'] else None)
    
    if updates:
        params.append(item_id)
        cursor.execute(f"UPDATE item SET {', '.join(updates)} WHERE id = ?", params)
        conn.commit()
    
    conn.close()
    return jsonify({'success': True})

@app.route('/api/item/<int:item_id>/close', methods=['POST'])
@login_required
def api_close_item(item_id):
    """Close out an item."""
    user_id = session['user_id']
    user_role = session.get('role')
    
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT * FROM item WHERE id = ?', (item_id,))
    item = cursor.fetchone()
    
    if not item:
        conn.close()
        return jsonify({'error': 'Item not found'}), 404
    
    if item['status'] == 'Closed':
        conn.close()
        return jsonify({'error': 'Item is already closed'}), 400
    
    if user_role != 'admin' and item['assigned_to_user_id'] != user_id:
        conn.close()
        return jsonify({'error': 'Not authorized to close this item'}), 403
    
    # Validate closeout requirements
    if not item['response_category']:
        conn.close()
        return jsonify({'error': 'Response category is required before closing'}), 400
    
    if not item['response_text']:
        conn.close()
        return jsonify({'error': 'Response text is required before closing'}), 400
    
    # Close the item
    cursor.execute('''
        UPDATE item SET status = 'Closed', closed_at = ? WHERE id = ?
    ''', (datetime.now().isoformat(), item_id))
    conn.commit()
    
    cursor.execute('SELECT * FROM item WHERE id = ?', (item_id,))
    updated_item = dict(cursor.fetchone())
    conn.close()
    
    return jsonify(updated_item)

@app.route('/api/item/<int:item_id>/reopen', methods=['POST'])
@admin_required
def api_reopen_item(item_id):
    """Reopen a closed item (admin only)."""
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT status FROM item WHERE id = ?', (item_id,))
    item = cursor.fetchone()
    
    if not item:
        conn.close()
        return jsonify({'error': 'Item not found'}), 404
    
    if item['status'] != 'Closed':
        conn.close()
        return jsonify({'error': 'Item is not closed'}), 400
    
    # Reopen - set to Ready for Response
    cursor.execute('''
        UPDATE item SET status = 'Ready for Response', closed_at = NULL WHERE id = ?
    ''', (item_id,))
    conn.commit()
    
    cursor.execute('SELECT * FROM item WHERE id = ?', (item_id,))
    updated_item = dict(cursor.fetchone())
    conn.close()
    
    return jsonify(updated_item)

@app.route('/api/item/<int:item_id>/files', methods=['GET'])
@login_required
def api_get_item_files(item_id):
    """List files in the item's folder."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT folder_link FROM item WHERE id = ?', (item_id,))
    item = cursor.fetchone()
    conn.close()
    
    if not item:
        return jsonify({'error': 'Item not found'}), 404
    
    folder_path = item['folder_link']
    if not folder_path:
        return jsonify([])
    
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        return jsonify([])
    
    # Allowed file extensions
    allowed_extensions = {'.pdf', '.doc', '.docx', '.xls', '.xlsx', '.dwg', '.dxf', '.png', '.jpg', '.jpeg', '.tif', '.tiff', '.zip'}
    
    files = []
    try:
        for f in folder.iterdir():
            if f.is_file() and f.suffix.lower() in allowed_extensions:
                stat = f.stat()
                files.append({
                    'filename': f.name,
                    'modified': datetime.fromtimestamp(stat.st_mtime).isoformat(),
                    'size': stat.st_size
                })
        # Sort by modified date descending
        files.sort(key=lambda x: x['modified'], reverse=True)
    except Exception as e:
        print(f"Error listing files: {e}")
    
    return jsonify({'files': files, 'folder_path': folder_path})


@app.route('/api/open-folder', methods=['POST'])
@login_required
def api_open_folder():
    """Open a folder in Windows Explorer."""
    data = request.get_json()
    folder_path = data.get('path', '')
    
    if not folder_path:
        return jsonify({'error': 'No folder path provided'}), 400
    
    folder = Path(folder_path)
    
    # Create folder if it doesn't exist
    if not folder.exists():
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            return jsonify({'error': f'Could not create folder: {e}'}), 500
    
    # Open folder in Windows Explorer
    try:
        import subprocess
        subprocess.Popen(['explorer', str(folder)])
        return jsonify({'success': True, 'path': str(folder)})
    except Exception as e:
        return jsonify({'error': f'Could not open folder: {e}'}), 500


@app.route('/api/item/<int:item_id>/open-email', methods=['POST'])
@login_required
def api_open_original_email(item_id):
    """Open the original email in Outlook for an item."""
    if not HAS_WIN32COM:
        return jsonify({'error': 'Outlook integration not available'}), 400
    
    conn = get_db()
    cursor = conn.cursor()
    
    # First try to get email_entry_id from item table
    cursor.execute('SELECT email_entry_id FROM item WHERE id = ?', (item_id,))
    item = cursor.fetchone()
    
    entry_id = None
    if item and item['email_entry_id']:
        entry_id = item['email_entry_id']
    else:
        # Fallback: get from email_log table
        cursor.execute('''
            SELECT entry_id FROM email_log 
            WHERE item_id = ? 
            ORDER BY received_at ASC 
            LIMIT 1
        ''', (item_id,))
        email_log = cursor.fetchone()
        if email_log:
            entry_id = email_log['entry_id']
    
    conn.close()
    
    if not entry_id:
        return jsonify({'error': 'No original email found for this item'}), 404
    
    try:
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Get the email by EntryID
            mail_item = namespace.GetItemFromID(entry_id)
            mail_item.Display()  # Opens the email in Outlook
            
            return jsonify({'success': True, 'message': 'Email opened in Outlook'})
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        return jsonify({'error': f'Could not open email: {str(e)}'}), 500

# =============================================================================
# API ROUTES - COMMENTS
# =============================================================================

@app.route('/api/comments/<int:item_id>', methods=['GET'])
@login_required
def api_get_comments(item_id):
    """Get comments for an item."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT c.*, u.display_name as author_name, u.email as author_email
        FROM comment c
        JOIN user u ON c.user_id = u.id
        WHERE c.item_id = ?
        ORDER BY c.created_at ASC
    ''', (item_id,))
    comments = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify(comments)

@app.route('/api/comments/<int:item_id>', methods=['POST'])
@login_required
def api_add_comment(item_id):
    """Add a comment to an item."""
    data = request.get_json()
    body = data.get('body', '').strip()
    
    if not body:
        return jsonify({'error': 'Comment body required'}), 400
    
    user_id = session['user_id']
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Verify item exists
    cursor.execute('SELECT id FROM item WHERE id = ?', (item_id,))
    if not cursor.fetchone():
        conn.close()
        return jsonify({'error': 'Item not found'}), 404
    
    cursor.execute('''
        INSERT INTO comment (item_id, user_id, body)
        VALUES (?, ?, ?)
    ''', (item_id, user_id, body))
    conn.commit()
    comment_id = cursor.lastrowid
    
    cursor.execute('''
        SELECT c.*, u.display_name as author_name, u.email as author_email
        FROM comment c
        JOIN user u ON c.user_id = u.id
        WHERE c.id = ?
    ''', (comment_id,))
    comment = dict(cursor.fetchone())
    conn.close()
    
    return jsonify(comment), 201

# =============================================================================
# API ROUTES - STATS
# =============================================================================

@app.route('/api/stats', methods=['GET'])
@login_required
def api_get_stats():
    """Get dashboard statistics."""
    user_id = session.get('user_id')
    conn = get_db()
    cursor = conn.cursor()
    
    # Total counts
    cursor.execute('SELECT COUNT(*) FROM item')
    total_items = cursor.fetchone()[0]
    
    # Counts by type
    cursor.execute('SELECT type, COUNT(*) FROM item GROUP BY type')
    by_type = {row[0]: row[1] for row in cursor.fetchall()}
    
    # Counts by bucket
    cursor.execute('SELECT bucket, COUNT(*) FROM item GROUP BY bucket')
    by_bucket = {row[0]: row[1] for row in cursor.fetchall()}
    
    # Counts by status
    cursor.execute('SELECT status, COUNT(*) FROM item GROUP BY status')
    by_status = {row[0]: row[1] for row in cursor.fetchall()}
    
    # Open items (not closed)
    cursor.execute("SELECT COUNT(*) FROM item WHERE closed_at IS NULL")
    open_items = cursor.fetchone()[0]
    
    # Overdue items
    today = datetime.now().strftime('%Y-%m-%d')
    cursor.execute('''
        SELECT COUNT(*) FROM item 
        WHERE due_date < ? AND closed_at IS NULL
    ''', (today,))
    overdue_items = cursor.fetchone()[0]
    
    # Items due this week
    week_end = (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d')
    cursor.execute('''
        SELECT COUNT(*) FROM item 
        WHERE due_date BETWEEN ? AND ? AND closed_at IS NULL
    ''', (today, week_end))
    due_this_week = cursor.fetchone()[0]
    
    # Unassigned items
    cursor.execute("SELECT COUNT(*) FROM item WHERE assigned_to_user_id IS NULL AND closed_at IS NULL")
    unassigned = cursor.fetchone()[0]
    
    # Inbox count for current user
    inbox_count = 0
    if user_id:
        cursor.execute("SELECT COUNT(*) FROM item WHERE assigned_to_user_id = ? AND closed_at IS NULL", (user_id,))
        inbox_count = cursor.fetchone()[0]
    
    conn.close()
    
    return jsonify({
        'total_items': total_items,
        'by_type': by_type,
        'by_bucket': by_bucket,
        'by_status': by_status,
        'open_items': open_items,
        'overdue_items': overdue_items,
        'due_this_week': due_this_week,
        'unassigned': unassigned,
        'inbox_count': inbox_count
    })

# =============================================================================
# API ROUTES - SYSTEM
# =============================================================================

@app.route('/api/poll-status', methods=['GET'])
@login_required
def api_poll_status():
    """Get email polling status."""
    return jsonify(email_poller.get_status())

@app.route('/api/poll-now', methods=['POST'])
@admin_required
def api_poll_now():
    """Trigger immediate email poll."""
    if HAS_WIN32COM:
        try:
            threading.Thread(target=email_poller._poll_emails, daemon=True).start()
            return jsonify({'success': True, 'message': 'Poll triggered'})
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    return jsonify({'error': 'Outlook integration not available'}), 400

@app.route('/api/config', methods=['GET'])
@admin_required
def api_get_config():
    """Get current configuration."""
    return jsonify(CONFIG)

@app.route('/api/config', methods=['POST'])
@admin_required
def api_update_config():
    """Update configuration."""
    global CONFIG
    data = request.get_json()
    
    for key in ['base_folder_path', 'outlook_folder', 'poll_interval_minutes', 'project_name']:
        if key in data:
            CONFIG[key] = data[key]
    
    with open(CONFIG_PATH, 'w') as f:
        json.dump(CONFIG, f, indent=2)
    
    return jsonify(CONFIG)

# =============================================================================
# MAGIC-LINK RESPONSE ROUTES
# =============================================================================

@app.route('/respond/reviewer', methods=['GET'])
def respond_reviewer_form():
    """Show reviewer response form via magic link."""
    token = request.args.get('token')
    if not token:
        return render_template_string(ERROR_PAGE_TEMPLATE, error='Missing token'), 400
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT i.*, 
               ir.display_name as reviewer_name, ir.email as reviewer_email,
               qcr.display_name as qcr_name
        FROM item i
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE i.email_token_reviewer = ?
    ''', (token,))
    item = cursor.fetchone()
    
    if not item:
        conn.close()
        return render_template_string(ERROR_PAGE_TEMPLATE, error='Invalid or expired token'), 404
    
    item_dict = dict(item)
    
    # Check if item is closed
    is_closed = item['status'] == 'Closed'
    
    # Determine if submission is allowed
    # Allowed when: NOT closed AND (QCR hasn't responded yet OR QCR sent it back)
    can_submit = False
    is_resubmit = False
    qcr_feedback = None
    
    if not is_closed:
        qcr_action = item['qcr_action']
        qcr_response_at = item['qcr_response_at']
        
        # Case 1: QCR hasn't responded yet - always allow
        if not qcr_response_at or item['qcr_response_status'] in ['Not Sent', 'Email Sent', 'Waiting for Revision']:
            can_submit = True
            # It's a resubmit if reviewer already responded before
            if item['reviewer_response_at']:
                is_resubmit = True
        
        # Case 2: QCR sent it back - allow resubmit
        if qcr_action == 'Send Back':
            can_submit = True
            is_resubmit = True
            qcr_feedback = item['qcr_notes']
        
        # Case 3: QCR approved/modified - don't allow (finalized)
        if qcr_action in ['Approve', 'Modify'] and item['status'] == 'Ready for Response':
            can_submit = False
    
    # Get current version info
    # First submission is v1, revisions are v2, v3, etc.
    current_version = item['reviewer_response_version'] or 0
    if is_resubmit:
        next_version = current_version + 1
    else:
        next_version = 1  # First submission is always v1
    
    # Get previous response for pre-fill
    previous_response = None
    previous_files = []
    if is_resubmit and item['reviewer_response_at']:
        previous_response = {
            'category': item['reviewer_response_category'],
            'text': item['reviewer_notes'] or item['reviewer_response_text'],
            'files': item['reviewer_selected_files']
        }
        if item['reviewer_selected_files']:
            try:
                previous_files = json.loads(item['reviewer_selected_files'])
            except:
                pass
    
    # Get version history
    version_history = ''
    cursor.execute('''
        SELECT version, submitted_at 
        FROM reviewer_response_history 
        WHERE item_id = ? 
        ORDER BY version DESC
    ''', (item['id'],))
    history = cursor.fetchall()
    if history:
        version_parts = [f"v{h['version']} ({h['submitted_at'][:16].replace('T', ' ')})" for h in history]
        version_history = ', '.join(version_parts)
    
    conn.close()
    
    # Get files in folder
    files = []
    if item['folder_link']:
        try:
            folder_path = Path(item['folder_link'])
            if folder_path.exists():
                files = [f.name for f in folder_path.iterdir() if f.is_file()]
        except:
            pass
    
    return render_template_string(REVIEWER_RESPONSE_TEMPLATE, 
        item=item_dict,
        files=files,
        token=token,
        version=next_version,
        is_closed=is_closed,
        can_submit=can_submit,
        is_resubmit=is_resubmit,
        qcr_feedback=qcr_feedback,
        previous_response=previous_response,
        previous_files=previous_files,
        version_history=version_history
    )

@app.route('/respond/reviewer', methods=['POST'])
def respond_reviewer_submit():
    """Handle reviewer response submission with version tracking."""
    token = request.form.get('token')
    if not token:
        return render_template_string(ERROR_PAGE_TEMPLATE, error='Missing token'), 400
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT i.*, ir.id as reviewer_user_id
        FROM item i
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        WHERE i.email_token_reviewer = ?
    ''', (token,))
    item = cursor.fetchone()
    
    if not item:
        conn.close()
        return render_template_string(ERROR_PAGE_TEMPLATE, error='Invalid or expired token'), 404
    
    item_id = item['id']
    
    # Check if item is closed - block all submissions
    if item['status'] == 'Closed':
        conn.close()
        return render_template_string(ERROR_PAGE_TEMPLATE, 
            error='This item has been closed. No further changes can be submitted. Contact the project administrator if this is unexpected.'), 403
    
    # Check if submission is allowed
    qcr_action = item['qcr_action']
    qcr_response_at = item['qcr_response_at']
    
    # Determine if this is a valid submission scenario
    can_submit = False
    is_resubmit = False
    
    # Case 1: QCR hasn't responded yet - always allow
    if not qcr_response_at or item['qcr_response_status'] in ['Not Sent', 'Email Sent', 'Waiting for Revision']:
        can_submit = True
        if item['reviewer_response_at']:
            is_resubmit = True
    
    # Case 2: QCR sent it back - allow resubmit
    if qcr_action == 'Send Back':
        can_submit = True
        is_resubmit = True
    
    # Case 3: QCR approved/modified - don't allow (finalized)
    if qcr_action in ['Approve', 'Modify'] and item['status'] == 'Ready for Response':
        can_submit = False
    
    if not can_submit:
        conn.close()
        return render_template_string(ERROR_PAGE_TEMPLATE, 
            error='This item has already been finalized in QC. Please contact the project admin if additional changes are required.'), 403
    
    # Get form data
    response_category = request.form.get('response_category')
    notes = request.form.get('notes', '')
    selected_files = request.form.getlist('selected_files')
    
    # Calculate new version
    # First submission is v1, revisions are v2, v3, etc.
    current_version = item['reviewer_response_version'] or 0
    new_version = current_version + 1 if current_version > 0 else 1
    
    # Track if this was a send-back scenario (before we reset qcr_action)
    was_sent_back = qcr_action == 'Send Back'
    
    # Store current response in history before updating (if this is a resubmit)
    if is_resubmit and item['reviewer_response_at']:
        cursor.execute('''
            INSERT INTO reviewer_response_history 
            (item_id, version, submitted_at, response_category, response_text, response_files, submitted_by_user_id)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (
            item_id,
            current_version,
            item['reviewer_response_at'],
            item['reviewer_response_category'],
            item['reviewer_notes'] or item['reviewer_response_text'],
            item['reviewer_selected_files'],
            item['reviewer_user_id']
        ))
    
    # Determine new status based on whether QCR sent it back
    if was_sent_back:
        # Reset QCR state for new review cycle
        cursor.execute('''
            UPDATE item SET
                reviewer_response_at = ?,
                reviewer_response_status = 'Responded',
                reviewer_response_category = ?,
                reviewer_notes = ?,
                reviewer_selected_files = ?,
                reviewer_response_version = ?,
                status = 'In QC',
                qcr_action = NULL,
                qcr_response_at = NULL,
                qcr_response_status = 'Not Sent',
                qcr_notes = NULL,
                qcr_response_category = NULL,
                qcr_response_text = NULL,
                qcr_response_mode = NULL
            WHERE id = ?
        ''', (
            datetime.now().isoformat(),
            response_category,
            notes,
            json.dumps(selected_files),
            new_version,
            item_id
        ))
    else:
        # Standard update
        cursor.execute('''
            UPDATE item SET
                reviewer_response_at = ?,
                reviewer_response_status = 'Responded',
                reviewer_response_category = ?,
                reviewer_notes = ?,
                reviewer_selected_files = ?,
                reviewer_response_version = ?,
                status = 'In QC'
            WHERE id = ?
        ''', (
            datetime.now().isoformat(),
            response_category,
            notes,
            json.dumps(selected_files),
            new_version,
            item_id
        ))
    
    conn.commit()
    conn.close()
    
    # Send appropriate notifications
    if is_resubmit:
        # Send version update notification to QCR
        if was_sent_back:
            # Full new QCR assignment email for revision after send-back
            send_qcr_assignment_email(item_id, is_revision=True, version=new_version)
        else:
            # Just an update notification (QCR hasn't responded yet)
            send_qcr_version_update_email(item_id, new_version)
    else:
        # First submission - send QCR assignment
        send_qcr_assignment_email(item_id)
    
    if is_resubmit:
        return render_template_string(SUCCESS_TEMPLATE, 
            message=f'Your revised response (v{new_version}) has been submitted!',
            details='The QC Reviewer has been notified of your updated response.'
        )
    else:
        return render_template_string(SUCCESS_TEMPLATE, 
            message='Your review has been submitted successfully!',
            details='The QC Reviewer has been notified and will complete the final review.'
        )

@app.route('/respond/qcr', methods=['GET'])
def respond_qcr_form():
    """Show QCR response form via magic link."""
    token = request.args.get('token')
    if not token:
        return render_template_string(ERROR_PAGE_TEMPLATE, error='Missing token'), 400
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT i.*, 
               ir.display_name as reviewer_name, ir.email as reviewer_email,
               qcr.display_name as qcr_name, qcr.email as qcr_email
        FROM item i
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE i.email_token_qcr = ?
    ''', (token,))
    item = cursor.fetchone()
    
    if not item:
        conn.close()
        return render_template_string(ERROR_PAGE_TEMPLATE, error='Invalid or expired token'), 404
    
    # Check if item is closed
    if item['status'] == 'Closed':
        conn.close()
        return render_template_string(ERROR_PAGE_TEMPLATE, 
            error='This item has been closed. No further changes can be submitted. Contact the project administrator if this is unexpected.'), 403
    
    # Check if already responded
    if item['qcr_response_at']:
        conn.close()
        return render_template_string(ALREADY_RESPONDED_TEMPLATE, 
            item=dict(item),
            response_type='qcr'
        )
    
    # Get version history
    cursor.execute('''
        SELECT version, submitted_at, response_category
        FROM reviewer_response_history 
        WHERE item_id = ? 
        ORDER BY version DESC
        LIMIT 5
    ''', (item['id'],))
    version_history = cursor.fetchall()
    conn.close()
    
    # Get files in folder
    files = []
    if item['folder_link']:
        try:
            folder_path = Path(item['folder_link'])
            if folder_path.exists():
                files = [f.name for f in folder_path.iterdir() if f.is_file()]
        except:
            pass
    
    # Parse reviewer selected files
    reviewer_files = []
    if item['reviewer_selected_files']:
        try:
            reviewer_files = json.loads(item['reviewer_selected_files'])
        except:
            pass
    
    # Get version info
    current_version = item['reviewer_response_version'] or 1
    
    return render_template_string(QCR_RESPONSE_TEMPLATE, 
        item=dict(item),
        files=files,
        reviewer_files=reviewer_files,
        token=token,
        version=current_version,
        version_history=[dict(v) for v in version_history] if version_history else []
    )

@app.route('/respond/qcr', methods=['POST'])
def respond_qcr_submit():
    """Handle QCR response submission."""
    token = request.form.get('token')
    if not token:
        return render_template_string(ERROR_PAGE_TEMPLATE, error='Missing token'), 400
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM item WHERE email_token_qcr = ?', (token,))
    item = cursor.fetchone()
    
    if not item:
        conn.close()
        return render_template_string(ERROR_PAGE_TEMPLATE, error='Invalid or expired token'), 404
    
    # Check if item is closed - block all submissions
    if item['status'] == 'Closed':
        conn.close()
        return render_template_string(ERROR_PAGE_TEMPLATE, 
            error='This item has been closed. No further changes can be submitted. Contact the project administrator if this is unexpected.'), 403
    
    # Check if already responded
    if item['qcr_response_at']:
        conn.close()
        return render_template_string(ALREADY_RESPONDED_TEMPLATE, 
            item=dict(item),
            response_type='qcr'
        )
    
    # Get form data
    qc_action = request.form.get('qc_action')  # Approve, Modify, Send Back
    response_mode = request.form.get('response_mode', 'Keep')  # Keep, Tweak, Revise
    response_category = request.form.get('response_category')
    response_text = request.form.get('response_text', '')
    qcr_notes = request.form.get('qcr_notes', '')
    selected_files = request.form.getlist('selected_files')
    
    item_id = item['id']
    item_dict = dict(item)
    
    # Determine final values based on QC action
    if qc_action == 'Send Back':
        # Item goes back to reviewer - set status to indicate revision needed
        new_status = 'In Review'  # Back to review stage
        final_category = None
        final_text = None
        final_files = None
        qcr_response_category = None  # No final category for send back
        qcr_selected_files = None
    else:
        # Approve or Modify - finalize the response
        new_status = 'Ready for Response'
        
        # Determine final response text based on mode
        if response_mode == 'Keep':
            final_text = item['reviewer_response_text'] or item['reviewer_notes'] or ''
        elif response_mode == 'Tweak' or response_mode == 'Revise':
            final_text = response_text
        else:
            final_text = item['reviewer_response_text'] or item['reviewer_notes'] or ''
        
        final_category = response_category
        final_files = json.dumps(selected_files) if selected_files else item['reviewer_selected_files']
        qcr_response_category = response_category
        qcr_selected_files = json.dumps(selected_files) if selected_files else None
    
    # Update item with all QCR response data
    cursor.execute('''
        UPDATE item SET
            qcr_response_at = ?,
            qcr_response_status = 'Responded',
            qcr_action = ?,
            qcr_response_mode = ?,
            qcr_response_category = ?,
            qcr_response_text = ?,
            qcr_notes = ?,
            qcr_selected_files = ?,
            final_response_category = ?,
            final_response_text = ?,
            final_response_files = ?,
            status = ?
        WHERE id = ?
    ''', (
        datetime.now().isoformat(),
        qc_action,
        response_mode if qc_action != 'Send Back' else None,
        qcr_response_category,
        response_text if qc_action != 'Send Back' else None,
        qcr_notes,
        qcr_selected_files,
        final_category,
        final_text,
        final_files,
        new_status,
        item_id
    ))
    
    # If sending back, also clear the reviewer's response so they can resubmit
    if qc_action == 'Send Back':
        cursor.execute('''
            UPDATE item SET
                reviewer_response_at = NULL,
                reviewer_response_status = 'Email Sent',
                qcr_response_status = 'Waiting for Revision'
            WHERE id = ?
        ''', (item_id,))
    
    conn.commit()
    
    # Get item details for notification
    cursor.execute('SELECT type, identifier, title FROM item WHERE id = ?', (item_id,))
    item_info = cursor.fetchone()
    conn.close()
    
    # Send notification email to reviewer
    send_reviewer_notification_email(
        item_id, 
        qc_action, 
        qcr_notes, 
        final_category=final_category,
        final_text=final_text
    )
    
    # Create system notifications based on QC action
    if qc_action == 'Approve' or qc_action == 'Modify':
        # Notification that response is ready to send to contractor
        create_notification(
            'response_ready',
            f'✅ Response Ready: {item_info["type"]} {item_info["identifier"]}',
            f'QC review complete. The response for "{item_info["title"] or item_info["identifier"]}" is ready to be sent to the contractor. Final category: {final_category}',
            item_id=item_id,
            action_url=f'/api/items/{item_id}/complete',
            action_label='Mark Complete'
        )
    elif qc_action == 'Send Back':
        # Notification that item was sent back
        create_notification(
            'sent_back',
            f'↩️ Sent Back: {item_info["type"]} {item_info["identifier"]}',
            f'The item "{item_info["title"] or item_info["identifier"]}" has been sent back to the reviewer for revisions.',
            item_id=item_id
        )
    
    # Return appropriate success message
    if qc_action == 'Approve':
        return render_template_string(SUCCESS_TEMPLATE, 
            message='Response Approved!',
            details='The reviewer has been notified. The item is now ready for closeout.'
        )
    elif qc_action == 'Modify':
        return render_template_string(SUCCESS_TEMPLATE, 
            message='Response Modified and Finalized!',
            details='The reviewer has been notified of your modifications. The item is now ready for closeout.'
        )
    else:  # Send Back
        return render_template_string(SUCCESS_TEMPLATE, 
            message='Item Sent Back to Reviewer',
            details='The reviewer has been notified and will receive a link to revise their response.'
        )

# =============================================================================
# ADMIN WORKFLOW API
# =============================================================================

@app.route('/api/admin/workflow', methods=['GET'])
@admin_required
def api_admin_workflow():
    """Get workflow status for all items."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT i.id, i.type, i.identifier, i.title, i.status, i.priority,
               i.initial_reviewer_due_date, i.qcr_due_date, i.folder_link,
               i.reviewer_email_sent_at, i.reviewer_response_at, i.reviewer_response_status,
               i.reviewer_response_category, i.reviewer_notes, i.reviewer_selected_files,
               i.reviewer_response_text,
               i.qcr_email_sent_at, i.qcr_response_at, i.qcr_response_status,
               i.qcr_response_category, i.qcr_notes, i.qcr_selected_files,
               i.qcr_action, i.qcr_response_mode, i.qcr_response_text,
               i.final_response_category, i.final_response_text, i.final_response_files,
               ir.display_name as reviewer_name, ir.email as reviewer_email,
               qcr.display_name as qcr_name, qcr.email as qcr_email
        FROM item i
        LEFT JOIN user ir ON i.initial_reviewer_id = ir.id
        LEFT JOIN user qcr ON i.qcr_id = qcr.id
        WHERE i.initial_reviewer_id IS NOT NULL OR i.qcr_id IS NOT NULL
        ORDER BY i.created_at DESC
    ''')
    items = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify(items)

@app.route('/api/admin/send_reviewer_email/<int:item_id>', methods=['POST'])
@admin_required
def api_send_reviewer_email(item_id):
    """Send or resend reviewer assignment email."""
    try:
        result = send_reviewer_assignment_email(item_id)
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 400
    except Exception as e:
        return jsonify({'success': False, 'error': f'Server error: {str(e)}'}), 500

@app.route('/api/admin/send_qcr_email/<int:item_id>', methods=['POST'])
@admin_required
def api_send_qcr_email(item_id):
    """Send or resend QCR assignment email."""
    try:
        result = send_qcr_assignment_email(item_id)
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 400
    except Exception as e:
        return jsonify({'success': False, 'error': f'Server error: {str(e)}'}), 500

# =============================================================================
# NOTIFICATIONS API
# =============================================================================

@app.route('/api/notifications', methods=['GET'])
@login_required
def api_get_notifications():
    """Get all notifications."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT n.*, i.type as item_type, i.identifier as item_identifier
        FROM notification n
        LEFT JOIN item i ON n.item_id = i.id
        ORDER BY n.created_at DESC
        LIMIT 100
    ''')
    notifications = [dict(row) for row in cursor.fetchall()]
    
    # Get unread count
    cursor.execute('SELECT COUNT(*) FROM notification WHERE read_at IS NULL')
    unread_count = cursor.fetchone()[0]
    
    conn.close()
    return jsonify({'notifications': notifications, 'unread_count': unread_count})

@app.route('/api/notifications/<int:notification_id>/read', methods=['POST'])
@login_required
def api_mark_notification_read(notification_id):
    """Mark a notification as read."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE notification SET read_at = ? WHERE id = ?
    ''', (datetime.now().isoformat(), notification_id))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/notifications/read-all', methods=['POST'])
@login_required
def api_mark_all_notifications_read():
    """Mark all notifications as read."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE notification SET read_at = ? WHERE read_at IS NULL
    ''', (datetime.now().isoformat(),))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/notifications/<int:notification_id>', methods=['DELETE'])
@login_required
def api_delete_notification(notification_id):
    """Delete a notification."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM notification WHERE id = ?', (notification_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/items/<int:item_id>/complete', methods=['POST'])
@admin_required
def api_mark_item_complete(item_id):
    """Mark an item as complete/closed."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE item SET 
            status = 'Closed',
            closed_at = ?
        WHERE id = ?
    ''', (datetime.now().isoformat(), item_id))
    conn.commit()
    conn.close()
    
    # Create notification
    create_notification(
        'item_closed',
        f'Item marked as complete',
        f'The item has been marked as complete and closed.',
        item_id=item_id
    )
    
    return jsonify({'success': True, 'message': 'Item marked as complete'})

# =============================================================================
# STATIC FILE ROUTES
# =============================================================================

@app.route('/')
def index():
    """Serve the main dashboard."""
    return send_from_directory('static', 'index.html')

@app.route('/<path:filename>')
def serve_static(filename):
    """Serve static files."""
    return send_from_directory('static', filename)

# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    """Main entry point."""
    print("=" * 60)
    print("LEB RFI/Submittal Tracker")
    print("=" * 60)
    
    # Initialize database
    print("Initializing database...")
    init_db()
    
    # Create base folder
    base_folder = Path(CONFIG['base_folder_path'])
    base_folder.mkdir(parents=True, exist_ok=True)
    print(f"Base folder: {base_folder}")
    
    # Start email poller
    print("Starting email poller...")
    email_poller.start()
    
    # Start web server
    port = CONFIG.get('server_port', 5000)
    print(f"\nStarting web server on http://localhost:{port}")
    print("Press Ctrl+C to stop\n")
    print("=" * 60)
    
    try:
        app.run(host='127.0.0.1', port=port, debug=False, threaded=True)
    except KeyboardInterrupt:
        print("\nShutting down...")
        email_poller.stop()

if __name__ == '__main__':
    main()
