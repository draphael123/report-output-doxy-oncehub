"""
Weekly Report Generator - Vercel Serverless Function
This imports the Flask app from the parent directory.
"""

import sys
import os

# Add parent directory to path so we can import app
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

# Import the Flask app from app.py
from app import app

# Required for Vercel - must be named 'application'
application = app
