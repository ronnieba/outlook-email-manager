"""
Configuration file for Outlook Email Manager
"""
import os

# Gemini API Configuration
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY', 'AIzaSyBOUWyZ-Dq2yPopzSZ6oopN7V6oeoB2iNY')

# Email Manager Configuration
MAX_EMAILS = 50
IMPORTANCE_THRESHOLD = 0.7

# Database Configuration
DATABASE_PATH = "email_preferences.db"
