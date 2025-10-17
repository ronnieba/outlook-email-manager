"""
Configuration file for Outlook Email Manager

⚠️ הגדרות חשובות:
   1. העתק את env.example ל-.env
   2. ערוך את .env והוסף את ה-API KEY שלך
   3. הקובץ .env לא מגובה ל-Git (נמצא ב-.gitignore)
"""
import os
from pathlib import Path

# טעינת משתני סביבה מקובץ .env (אם קיים)
def load_env_file():
    """טעינת קובץ .env אם קיים"""
    env_file = Path(__file__).parent / '.env'
    if env_file.exists():
        with open(env_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    os.environ[key.strip()] = value.strip()

# טעינת קובץ .env
load_env_file()

# =============================================================================
# Gemini API Configuration
# =============================================================================
# קבל API Key חינמי מ: https://makersuite.google.com/app/apikey
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY', '')

if not GEMINI_API_KEY or GEMINI_API_KEY == 'your-api-key-here':
    print("⚠️  אזהרה: GEMINI_API_KEY לא מוגדר!")
    print("   1. העתק את env.example ל-.env")
    print("   2. ערוך את .env והוסף את ה-API KEY שלך")
    print("   3. קבל API Key מ: https://makersuite.google.com/app/apikey")

# =============================================================================
# Flask Server Configuration
# =============================================================================
FLASK_PORT = int(os.getenv('FLASK_PORT', 5000))

# =============================================================================
# Email Manager Configuration
# =============================================================================
MAX_EMAILS = os.getenv('MAX_EMAILS', None)  # None = ללא הגבלה
if MAX_EMAILS and MAX_EMAILS != 'None':
    MAX_EMAILS = int(MAX_EMAILS)

IMPORTANCE_THRESHOLD = float(os.getenv('IMPORTANCE_THRESHOLD', 0.7))

# =============================================================================
# Database Configuration
# =============================================================================
DATABASE_PATH = os.getenv('DATABASE_PATH', 'email_manager.db')
PREFERENCES_DB_PATH = os.getenv('PREFERENCES_DB_PATH', 'email_preferences.db')

# =============================================================================
# Logging Configuration
# =============================================================================
LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')
