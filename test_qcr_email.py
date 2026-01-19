#!/usr/bin/env python3
"""Quick test script to send multi-reviewer QCR email."""

import sys
import os

# Add the app directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the function
from app import send_multi_reviewer_qcr_email

if __name__ == '__main__':
    item_id = 17  # Test item
    print(f"Sending multi-reviewer QCR email for item {item_id}...")
    result = send_multi_reviewer_qcr_email(item_id)
    print(f"Result: {result}")
