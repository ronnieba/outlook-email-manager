"""
Outlook Email Manager - With AI Integration
מערכת ניהול מיילים חכמה עם AI + Outlook + Gemini
"""
# השתקת stderr לפני הכל!
import sys
_original_stderr = sys.stderr
try:
    import os
    sys.stderr = open(os.devnull, 'w')
except:
    # אם נכשל, לפחות ננסה עם StringIO
    import io
    sys.stderr = io.StringIO()

import warnings

# חייב להיות לפני כל import אחר!
os.environ['GRPC_VERBOSITY'] = 'NONE'
os.environ['GRPC_TRACE'] = ''
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
os.environ['GLOG_minloglevel'] = '3'
os.environ['ABSL_MIN_LOG_LEVEL'] = '3'

# השתקת warnings
warnings.filterwarnings('ignore')

import logging
logging.basicConfig(level=logging.ERROR)
logging.getLogger('google').setLevel(logging.ERROR)
logging.getLogger('grpc').setLevel(logging.ERROR)
logging.getLogger('absl').setLevel(logging.ERROR)
logging.getLogger('werkzeug').setLevel(logging.ERROR)
logging.getLogger('flask.app').setLevel(logging.ERROR)

from flask import Flask, render_template, request, jsonify, Response, send_file
from flask_cors import CORS
import win32com.client
import json
import subprocess
import os
from datetime import datetime, timedelta
import uuid
import sqlite3
import random
import threading
import pythoncom
from ai_analyzer import EmailAnalyzer
from config import GEMINI_API_KEY
from user_profile_manager import UserProfileManager
from collapsible_logger import logger
import logging
import zipfile

# לא מחזירים את stderr עדיין - יהיו עוד imports של Google בזמן הרצת Flask
import shutil

# כיבוי לוגים של Werkzeug (HTTP requests)
logging.getLogger('werkzeug').setLevel(logging.WARNING)

app = Flask(__name__)
CORS(app)  # הוספת CORS לתמיכה בבקשות cross-origin

# רשימת כל הלוגים (לצורך הצגה בקונסול)
all_console_logs = []

# נתיב מאגר הנתונים
DB_PATH = 'email_manager.db'

# אתחול AI Analyzer (יאותחל בפעם הראשונה שנדרש)
email_analyzer = None

# ---------------------- AI analysis persistence (SQLite) ----------------------
def init_ai_analysis_table():
    try:
        conn = sqlite3.connect('email_manager.db')
        c = conn.cursor()
        c.execute(
            'CREATE TABLE IF NOT EXISTS email_ai_analysis ('
            'email_id TEXT PRIMARY KEY,'
            'ai_score REAL,'
            'score_source TEXT,'
            'summary TEXT,'
            'reason TEXT,'
            'analyzed_at TEXT,'
            'category TEXT,'
            'original_score REAL)'
        )
        
        # יצירת טבלה לניתוח AI של פגישות
        c.execute(
            'CREATE TABLE IF NOT EXISTS meeting_ai_analysis ('
            'meeting_id TEXT PRIMARY KEY,'
            'ai_score REAL,'
            'score_source TEXT,'
            'summary TEXT,'
            'reason TEXT,'
            'analyzed_at TEXT,'
            'category TEXT,'
            'original_score REAL,'
            'ai_processed BOOLEAN DEFAULT FALSE)'
        )
        
        # יצירת טבלה למיילים (לסיכומי AI מלאים)
        c.execute(
            'CREATE TABLE IF NOT EXISTS emails ('
            'id INTEGER PRIMARY KEY AUTOINCREMENT,'
            'outlook_id TEXT UNIQUE,'
            'subject TEXT,'
            'sender TEXT,'
            'ai_summary TEXT,'
            'last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP)'
        )
        
        conn.commit()
    finally:
        try:
            conn.close()
        except Exception:
            pass

def save_ai_analysis_to_db(email_data: dict) -> None:
    try:
        # יצירת מפתח ייחודי על בסיס תוכן המייל (נושא + שולח + תאריך)
        subject = email_data.get('subject', '')
        sender = email_data.get('sender', '')
        received_time = email_data.get('received_time', '')
        
        # יצירת hash ייחודי מהתוכן
        import hashlib
        content_key = f"{subject}|{sender}|{received_time}"
        email_id = hashlib.md5(content_key.encode('utf-8')).hexdigest()
        
        conn = sqlite3.connect('email_manager.db')
        c = conn.cursor()
        c.execute(
            'INSERT OR REPLACE INTO email_ai_analysis (email_id, ai_score, score_source, summary, reason, analyzed_at, category, original_score) '
            'VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
            (
                email_id,
                float(email_data.get('importance_score', email_data.get('ai_importance_score', 0.0)) or 0.0),
                email_data.get('score_source', 'SMART'),
                email_data.get('summary', ''),
                email_data.get('reason', ''),
                email_data.get('ai_analysis_date') or datetime.now().isoformat(),
                email_data.get('category', ''),
                float(email_data.get('original_importance_score', 0.0) or 0.0),
            )
        )
        conn.commit()
        # שמירה הצליחה
    except Exception as e:
        # שגיאה בשמירה - מתעלמים
        pass
    finally:
        try:
            conn.close()
        except Exception:
            pass

def load_ai_analysis_map() -> dict:
    result = {}
    try:
        conn = sqlite3.connect('email_manager.db')
        c = conn.cursor()
        for row in c.execute('SELECT email_id, ai_score, score_source, summary, reason, analyzed_at, category, original_score FROM email_ai_analysis'):
            email_id, ai_score, source, summary, reason, analyzed_at, category, original_score = row
            result[email_id] = {
                'importance_score': ai_score,
                'ai_importance_score': ai_score,
                'score_source': source,
                'summary': summary,
                'reason': reason,
                'ai_analysis_date': analyzed_at,
                'category': category,
                'original_importance_score': original_score,
                'ai_analyzed': source == 'AI',  # רק אם באמת נותח על ידי AI
            }
            # נטען מה-DB בהצלחה
    except Exception:
        return {}
    finally:
        try:
            conn.close()
        except Exception:
            pass
    return result

def save_meeting_ai_analysis_to_db(meeting_data: dict) -> None:
    """שמירת ניתוח AI של פגישה בבסיס הנתונים"""
    try:
        # יצירת מפתח ייחודי על בסיס תוכן הפגישה
        subject = meeting_data.get('subject', '')
        organizer = meeting_data.get('organizer', '')
        start_time = meeting_data.get('start_time', '')
        
        # יצירת hash ייחודי מהתוכן
        import hashlib
        content_key = f"{subject}|{organizer}|{start_time}"
        meeting_id = hashlib.md5(content_key.encode('utf-8')).hexdigest()
        
        conn = sqlite3.connect('email_manager.db')
        c = conn.cursor()
        c.execute(
            'INSERT OR REPLACE INTO meeting_ai_analysis (meeting_id, ai_score, score_source, summary, reason, analyzed_at, category, original_score, ai_processed) '
            'VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)',
            (
                meeting_id,
                float(meeting_data.get('importance_score', meeting_data.get('ai_importance_score', 0.0)) or 0.0),
                meeting_data.get('score_source', 'SMART'),
                meeting_data.get('summary', ''),
                meeting_data.get('reason', ''),
                meeting_data.get('ai_analysis_date') or datetime.now().isoformat(),
                meeting_data.get('category', ''),
                float(meeting_data.get('original_importance_score', 0.0) or 0.0),
                meeting_data.get('ai_processed', False)
            )
        )
        conn.commit()
        # פגישה נשמרה בהצלחה
    except Exception as e:
        # שגיאה בשמירה - מתעלמים
        pass
    finally:
        try:
            conn.close()
        except Exception:
            pass

def load_meeting_ai_analysis_map() -> dict:
    """טעינת מפת ניתוח AI של פגישות מבסיס הנתונים"""
    result = {}
    try:
        conn = sqlite3.connect('email_manager.db')
        c = conn.cursor()
        for row in c.execute('SELECT meeting_id, ai_score, score_source, summary, reason, analyzed_at, category, original_score, ai_processed FROM meeting_ai_analysis'):
            meeting_id, ai_score, source, summary, reason, analyzed_at, category, original_score, ai_processed = row
            result[meeting_id] = {
                'importance_score': ai_score,
                'ai_importance_score': ai_score,
                'score_source': source,
                'summary': summary,
                'reason': reason,
                'ai_analysis_date': analyzed_at,
                'category': category,
                'original_importance_score': original_score,
                'ai_processed': ai_processed,
                'ai_analyzed': source == 'AI',  # רק אם באמת נותח על ידי AI
            }
            # פגישה נטענה מה-DB
    except Exception:
        return {}
    finally:
        try:
            conn.close()
        except Exception:
            pass

    return result

def apply_ai_analysis_from_db(emails: list) -> None:
    """ממזג תוצאות AI שנשמרו בבסיס נתונים לתוך רשימת המיילים הטעונה."""
    try:
        saved = load_ai_analysis_map()
        if not saved:
            return
        
        # יצירת מפתח ייחודי לכל מייל
        import hashlib
        for e in emails:
            subject = e.get('subject', '')
            sender = e.get('sender', '')
            received_time = e.get('received_time', '')
            
            # יצירת hash ייחודי מהתוכן
            content_key = f"{subject}|{sender}|{received_time}"
            email_id = hashlib.md5(content_key.encode('utf-8')).hexdigest()
            
            a = saved.get(email_id)
            if a:
                # נמצא ניתוח שמור
                # עדכון כל השדות הרלוונטיים
                e.update(a)
                # וידוא שהמייל מסומן כנותח על ידי AI רק אם באמת נותח
                if a.get('score_source') == 'AI':
                    e['ai_analyzed'] = True
                else:
                    e['ai_analyzed'] = False
                # שמירת הסיכום וההסבר גם בשדות נפרדים
                if a.get('summary'):
                    e['ai_summary'] = a['summary']
                if a.get('reason'):
                    e['ai_reason'] = a['reason']
    except Exception:
        pass

def apply_meeting_ai_analysis_from_db(meetings: list) -> None:
    """ממזג תוצאות AI שנשמרו בבסיס נתונים לתוך רשימת הפגישות הטעונה."""
    try:
        saved = load_meeting_ai_analysis_map()
        if not saved:
            return
        
        # יצירת מפתח ייחודי לכל פגישה
        import hashlib
        for m in meetings:
            subject = m.get('subject', '')
            organizer = m.get('organizer', '')
            start_time = m.get('start_time', '')
            
            # יצירת hash ייחודי מהתוכן
            content_key = f"{subject}|{organizer}|{start_time}"
            meeting_id = hashlib.md5(content_key.encode('utf-8')).hexdigest()
            
            a = saved.get(meeting_id)
            if a:
                # נמצא ניתוח שמור לפגישה
                # עדכון כל השדות הרלוונטיים
                m.update(a)
                # וידוא שהפגישה מסומנת כנותחת על ידי AI רק אם באמת נותחה
                if a.get('score_source') == 'AI':
                    m['ai_analyzed'] = True
                else:
                    m['ai_analyzed'] = False
                # שמירת הסיכום וההסבר גם בשדות נפרדים
                if a.get('summary'):
                    m['ai_summary'] = a['summary']
                if a.get('reason'):
                    m['ai_reason'] = a['reason']
    except Exception:
        pass
# מזהה ייחודי לשרת (משתנה בכל הפעלה)
server_id = datetime.now().strftime("%Y%m%d_%H%M%S")

# Cache למידע - נטען פעם אחת בהפעלת השרת
cached_data = {
    'emails': None,
    'meetings': None,
    'email_stats': None,
    'meeting_stats': None,
    'last_updated': None,
    'is_loading': False
}

# מצב: מצמצם הדפסות לטרמינל – רק תקלות תשתיתיות חמורות
MINIMAL_TERMINAL_LOG = True
# רמת לוג מינימלית להדפסה לטרמינל (ברירת מחדל: CRITICAL בלבד)
TERMINAL_LOG_LEVEL = os.environ.get('TERMINAL_LOG_LEVEL', 'CRITICAL').upper()
_LEVEL_ORDER = {'DEBUG': 10, 'INFO': 20, 'SUCCESS': 25, 'WARNING': 30, 'ERROR': 40, 'CRITICAL': 50}

def _should_print_to_terminal(level: str) -> bool:
    # כבוי לחלוטין - הכל רק לקונסול Web
    return False

def log_to_console(message, level="INFO"):
    """הוספת הודעה לקונסול (מדפיס לטרמינל רק שגיאות קשות)."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    
    # ניקוי המילים באנגלית מההודעה לפני שמירה
    clean_message = message
    if level == "INFO" and message.startswith("INFO: "):
        clean_message = message[6:]  # הסרת "INFO: "
    elif level == "SUCCESS" and message.startswith("SUCCESS: "):
        clean_message = message[9:]  # הסרת "SUCCESS: "
    elif level == "ERROR" and message.startswith("ERROR: "):
        clean_message = message[7:]  # הסרת "ERROR: "
    elif level == "WARNING" and message.startswith("WARNING: "):
        clean_message = message[9:]  # הסרת "WARNING: "
    
    # ניקוי תווים בעייתיים לפני הדפסה
    safe_message = clean_message.encode('ascii', errors='ignore').decode('ascii')
    
    log_entry = {
        'message': clean_message,  # שמירת ההודעה הנקייה לרשימה
        'level': level,
        'timestamp': timestamp
    }
    all_console_logs.append(log_entry)
    
    # הדפסה לטרמינל – רק במקרי תקלות/קריטיות או אם מצב מינימלי כבוי
    if _should_print_to_terminal(level):
        print(f"[{timestamp}] {safe_message}")

# ===== Server-driven collapsible blocks for UI =====
def ui_block_start(title: str) -> str:
    """יוצר אירוע פתיחת בלוק מובנה לקונסול ומחזיר block_id."""
    block_id = uuid.uuid4().hex[:8]
    all_console_logs.append({
        'type': 'block_start',
        'block_id': block_id,
        'title': title,
        'timestamp': datetime.now().strftime("%H:%M:%S"),
        'level': 'INFO'
    })
    return block_id

def ui_block_add(block_id: str, message: str, level: str = 'INFO') -> None:
    all_console_logs.append({
        'type': 'block_content',
        'block_id': block_id,
        'message': message,
        'timestamp': datetime.now().strftime("%H:%M:%S"),
        'level': level
    })

def ui_block_end(block_id: str, summary: str | None = None, success: bool = True) -> None:
    all_console_logs.append({
        'type': 'block_end',
        'block_id': block_id,
        'summary': summary or ("הושלם" if success else "נכשל"),
        'success': bool(success),
        'timestamp': datetime.now().strftime("%H:%M:%S"),
        'level': 'SUCCESS' if success else 'ERROR'
    })

# הגדרת מערכת הלוגים החדשה להשתמש ב-log_to_console
logger.set_console_logger(log_to_console)

def load_initial_data():
    """טעינת המידע הראשונית לזיכרון"""
    global cached_data
    
    # אם כבר נטענו מיילים – אין צורך לטעון שוב
    try:
        try:
            init_ai_analysis_table()
        except Exception:
            pass
        if cached_data.get('emails'):
            return
    except Exception:
        pass

    if cached_data['is_loading']:
        logger.log_warning("Data loading already in progress...")
        return
    
    cached_data['is_loading'] = True
    
    # התחלת בלוק טעינת נתונים
    block_id = logger.start_block(
        "טעינת נתונים ראשונית", 
        "טוען מיילים ופגישות מ-Outlook"
    )
    
    try:
        # יצירת EmailManager
        logger.add_to_block(block_id, "יוצר מנהל מיילים...")
        email_manager = EmailManager()
        
        # טעינת מיילים
        logger.add_to_block(block_id, "טוען מיילים מ-Outlook...")
        emails = email_manager.get_emails()
        
        # מיזוג נתוני AI שמורים מהבסיס
        logger.add_to_block(block_id, "ממזג נתוני AI שמורים...")
        try:
            apply_ai_analysis_from_db(emails)
            ai_count = sum(1 for email in emails if email.get('ai_analyzed', False))
            logger.add_to_block(block_id, f"נתוני AI הוטענו מהבסיס בהצלחה - {ai_count} מיילים נותחו בעבר")
        except Exception as e:
            logger.add_to_block(block_id, f"שגיאה בטעינת נתוני AI: {e}")
        
        # ניתוח חכם של המיילים (רק מיילים שלא נותחו בעבר)
        logger.add_to_block(block_id, "מנתח מיילים עם ניתוח חכם...")
        emails = email_manager.analyze_emails_smart(emails)
        
        cached_data['emails'] = emails
        logger.add_to_block(block_id, f"נטענו {len(emails)} מיילים")
        
        # טעינת פגישות
        logger.add_to_block(block_id, "טוען פגישות...")
        meetings = email_manager.get_meetings()
        cached_data['meetings'] = meetings
        logger.add_to_block(block_id, f"נטענו {len(meetings)} פגישות")
        
        # חישוב סטטיסטיקות מיילים
        logger.add_to_block(block_id, "מחשב סטטיסטיקות מיילים...")
        email_stats = calculate_email_stats(emails)
        cached_data['email_stats'] = email_stats
        
        # חישוב סטטיסטיקות פגישות
        logger.add_to_block(block_id, "מחשב סטטיסטיקות פגישות...")
        meeting_stats = calculate_meeting_stats(meetings)
        cached_data['meeting_stats'] = meeting_stats
        
        cached_data['last_updated'] = datetime.now()
        cached_data['is_loading'] = False
        
        # סיום הבלוק בהצלחה
        logger.end_block(
            block_id, 
            success=True, 
            summary=f"נטענו {len(emails)} מיילים ו-{len(meetings)} פגישות בהצלחה"
        )
        
    except Exception as e:
        cached_data['is_loading'] = False
        logger.end_block(block_id, success=False, summary=f"שגיאה בטעינת נתונים: {str(e)}")
        logger.log_error(f"Error loading initial data: {str(e)}")

def calculate_email_stats(emails):
    """חישוב סטטיסטיקות מיילים"""
    total_emails = len(emails)
    
    # התפלגות קבועה לפי הדרישות:
    # 10% קריטיים, 25% חשובים, 40% בינוניים, 25% נמוכים
    critical_emails = int(total_emails * 0.10)  # 10%
    important_emails = int(total_emails * 0.25)  # 25%
    medium_emails = int(total_emails * 0.40)     # 40%
    low_emails = int(total_emails * 0.25)        # 25%
    
    # מיילים שלא נקראו בפועל
    actual_unread_emails = len([e for e in emails if not e.get('is_read', True)])
    
    return {
        'total_emails': total_emails,
        'important_emails': important_emails,
        'unread_emails': actual_unread_emails,
        'critical_emails': critical_emails,
        'high_emails': important_emails,
        'medium_emails': medium_emails,
        'low_emails': low_emails
    }

def calculate_meeting_stats(meetings):
    """חישוב סטטיסטיקות פגישות לפי ציונים"""
    total_meetings = len(meetings)
    
    # חישוב קטגוריות לפי ציונים (10% קריטי, 25% חשוב, 35% בינוני, 20% נמוך)
    critical_meetings = 0
    important_meetings = 0
    medium_meetings = 0
    low_meetings = 0
    
    for meeting in meetings:
        score = meeting.get('importance_score', 0.5)
        if score >= 0.8:  # 80% ומעלה = קריטי
            critical_meetings += 1
        elif score >= 0.6:  # 60-79% = חשוב
            important_meetings += 1
        elif score >= 0.4:  # 40-59% = בינוני
            medium_meetings += 1
        else:  # מתחת ל-40% = נמוך
            low_meetings += 1
    
    # פגישות היום
    today_meetings = len([m for m in meetings if m.get('is_today', False)])
    
    # פגישות השבוע
    week_meetings = len([m for m in meetings if m.get('is_this_week', False)])
    
    return {
        'total_meetings': total_meetings,
        'critical_meetings': critical_meetings,
        'important_meetings': important_meetings,
        'medium_meetings': medium_meetings,
        'low_meetings': low_meetings,
        'today_meetings': today_meetings,
        'week_meetings': week_meetings
    }

def refresh_data(data_type=None):
    """רענון המידע בזיכרון"""
    global cached_data
    
    if cached_data['is_loading']:
        log_to_console("Data refresh already in progress...", "WARNING")
        return False
    
    cached_data['is_loading'] = True
    log_to_console(f"Starting data refresh ({data_type or 'all data'})...", "INFO")
    
    try:
        # אתחול טבלאות AI
        init_ai_analysis_table()
        
        # יצירת EmailManager
        email_manager = EmailManager()
        
        if data_type is None or data_type == 'emails':
            # רענון מיילים
            log_to_console("Refreshing emails...", "INFO")
            emails = email_manager.get_emails()
            cached_data['emails'] = emails
            log_to_console(f"Updated {len(emails)} emails", "SUCCESS")
            
            # חישוב סטטיסטיקות מיילים
            log_to_console("Calculating email statistics...", "INFO")
            email_stats = calculate_email_stats(emails)
            cached_data['email_stats'] = email_stats
        
        if data_type is None or data_type == 'meetings':
            # רענון פגישות
            log_to_console("📅 מרענן פגישות...", "INFO")
            meetings = email_manager.get_meetings()
            cached_data['meetings'] = meetings
            log_to_console(f"Updated {len(meetings)} meetings", "SUCCESS")
            
            # חישוב סטטיסטיקות פגישות
            log_to_console("Calculating meeting statistics...", "INFO")
            meeting_stats = calculate_meeting_stats(meetings)
            cached_data['meeting_stats'] = meeting_stats
        
        cached_data['last_updated'] = datetime.now()
        cached_data['is_loading'] = False
        
        log_to_console("🎉 רענון נתונים הושלם!", "SUCCESS")
        return True
        
    except Exception as e:
        cached_data['is_loading'] = False
        log_to_console(f"Error in data refresh: {str(e)}", "ERROR")
        return False

class EmailManager:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.inbox = None
        self.user_preferences = {}
        self.db_path = "email_preferences.db"
        self.ai_analyzer = EmailAnalyzer()
        self.profile_manager = UserProfileManager(self.db_path)
        self.use_ai = True
        self.use_learning = True
        self.init_database()
        self.load_user_preferences()
        self.outlook_connected = False
    
    def init_database(self):
        """יצירת מסד נתונים לניהול העדפות"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # טבלת העדפות משתמש
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_preferences (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                preference_type TEXT NOT NULL,
                preference_value TEXT NOT NULL,
                weight REAL DEFAULT 1.0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # טבלת מיילים שסומנו כחשובים
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS important_emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                subject TEXT,
                sender TEXT,
                received_time TIMESTAMP,
                importance_score REAL,
                user_feedback TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # טבלת ניתוחי AI
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS ai_analysis (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email_id INTEGER,
                importance_score REAL,
                category TEXT,
                summary TEXT,
                action_items TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # טבלת העדפות משתמש מתקדמות
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_preferences_advanced (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                preference_type TEXT NOT NULL,
                preference_key TEXT NOT NULL,
                preference_value TEXT NOT NULL,
                confidence_score REAL DEFAULT 0.5,
                usage_count INTEGER DEFAULT 1,
                last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def connect_to_outlook(self):
        """חיבור ל-Outlook"""
        try:
            # אם כבר מחובר – אל תחבר שוב ואל תדפיס לוגים מיותרים
            if getattr(self, 'outlook_connected', False) and getattr(self, 'namespace', None) is not None:
                return True
            # אתחול COM רק אם לא מאותחל כבר
            try:
                pythoncom.CoInitialize()
            except:
                pass  # כבר מאותחל
            
            log_to_console("🔌 מנסה להתחבר ל-Outlook...", "INFO")
            log_to_console("Trying to connect to Outlook...", "INFO")
            
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            log_to_console("Outlook Application connection successful!", "SUCCESS")
            log_to_console("Outlook Application connection successful!", "SUCCESS")
            
            # חיפוש בכל התיקיות, לא רק Inbox
            self.inbox = self.namespace.GetDefaultFolder(6)  # Inbox הראשי
            
            log_to_console("Inbox folder connection successful!", "SUCCESS")
            log_to_console("Inbox folder connection successful!", "SUCCESS")
            
            # בדיקת מספר המיילים ב-Inbox
            try:
                messages = self.inbox.Items
                # print(f"Found {messages.Count} emails in Inbox")
                log_to_console(f"Found {messages.Count} emails in Inbox", "INFO")
            except Exception as e:
                log_to_console(f"Cannot count emails: {e}", "ERROR")
                log_to_console(f"Cannot count emails: {e}", "WARNING")
            
            # נסה לקבל גישה לכל המיילים בחשבון
            try:
                # קבלת החשבון הראשי
                self.account = self.namespace.Accounts.Item(1)
                # קבלת תיקיית הרכיבים הראשית
                self.root_folder = self.account.DeliveryStore.GetRootFolder()
                log_to_console(f"Found account: {self.account.DisplayName}", "INFO")
                log_to_console(f"Found account: {self.account.DisplayName}", "INFO")
            except:
                # fallback לתיקיית Inbox הרגילה
                log_to_console("Using regular Inbox folder", "INFO")
                log_to_console("Using regular Inbox folder", "WARNING")
            
            self.outlook_connected = True
            log_to_console("Outlook connection successful!", "SUCCESS")
            log_to_console("Outlook connection successful!", "SUCCESS")
            return True
        except Exception as e:
            log_to_console(f"Error connecting to Outlook: {e}", "ERROR")
            log_to_console(f"Error connecting to Outlook: {e}", "ERROR")
            self.outlook_connected = False
            return False
    
    def get_emails(self, limit=None):  # ללא הגבלה - יטען את כל המיילים
        """קבלת מיילים - מועדפת קריאה מהקאש בזיכרון למניעת טעינות חוזרות."""
        try:
            # שימוש בנתונים מהקאש הגלובלי אם קיימים
            global cached_data
            if cached_data.get('emails'):
                return cached_data['emails'][:limit] if limit else cached_data['emails']

            # אחרת נטען מ-Outlook פעם אחת ונשמור בקאש
            emails = self.get_emails_from_outlook(limit)
            # מיזוג ניתוחי AI ששמורים בבסיס נתונים
            try:
                init_ai_analysis_table()
                apply_ai_analysis_from_db(emails)
            except Exception:
                pass
            if emails and len(emails) > 0:
                cached_data['emails'] = emails
                log_to_console(f"Loaded {len(emails)} real emails from Outlook", "INFO")
                return emails
            else:
                # fallback לנתונים דמה
                log_to_console("Using demo data", "WARNING")
                sample = self.get_sample_emails()
                cached_data['emails'] = sample
                return sample
        except Exception as e:
            log_to_console(f"Error getting emails: {e}", "ERROR")
            sample = self.get_sample_emails()
            try:
                cached_data['emails'] = sample
            except Exception:
                pass
            return sample
    
    def _clean_email_body(self, body):
        """ניקוי ופענוח תוכן מייל מ-Outlook"""
        if not body:
            return ""
        
        try:
            # המרה למחרוזת
            body_str = str(body)
            
            # ניסיון פענוח URL encoding
            import urllib.parse
            try:
                # פענוח URL encoding (עד 3 רמות)
                for _ in range(3):
                    decoded = urllib.parse.unquote(body_str)
                    if decoded == body_str:
                        break
                    body_str = decoded
            except:
                pass
            
            # ניקוי HTML tags
            import re
            body_str = re.sub(r'<[^>]+>', '', body_str)
            
            # ניקוי HTML entities
            html_entities = {
                '&amp;': '&',
                '&lt;': '<',
                '&gt;': '>',
                '&quot;': '"',
                '&#39;': "'",
                '&nbsp;': ' ',
                '&copy;': '©',
                '&reg;': '®',
                '&trade;': '™'
            }
            for entity, char in html_entities.items():
                body_str = body_str.replace(entity, char)
            
            # ניקוי תווים מיוחדים אבל שמירה על עברית
            body_str = re.sub(r'[^\w\s\u0590-\u05FF\u2000-\u206F\u2E00-\u2E7F\s\.,!?;:()\[\]{}"\'@#$%^&*+=<>/\\|`~-]', '', body_str)
            
            # ניקוי רווחים מיותרים
            body_str = re.sub(r'\s+', ' ', body_str).strip()
            
            return body_str
            
        except Exception as e:
            # fallback - החזרת התוכן המקורי
            return str(body) if body else ""

    def get_emails_from_outlook(self, limit=None):  # ללא הגבלה - יטען את כל המיילים
        """קבלת מיילים אמיתיים מ-Outlook"""
        try:
            # התחל בלוק UI עבור טעינת מיילים
            block_id = ui_block_start("📧 טעינת מיילים מ-Outlook")
            ui_block_add(block_id, "מתחיל טעינת מיילים מ-Outlook...", "INFO")
            # אתחול COM רק אם לא מאותחל כבר
            try:
                pythoncom.CoInitialize()
            except:
                pass  # כבר מאותחל
            
            # יצירת חיבור חדש בכל קריאה כדי למנוע בעיות threading
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            ui_block_add(block_id, "Searching all emails in Inbox...", "INFO")
            
            # גישה ישירה לתיקיית Inbox
            inbox_folder = namespace.GetDefaultFolder(6)  # Inbox
            messages = inbox_folder.Items
            
            ui_block_add(block_id, f"Found {messages.Count} emails in Inbox", "INFO")
            
            # מיון לפי תאריך - חדשים קודם. פעולה זו יכולה "להכריח" את Outlook לטעון את כל המיילים.
            messages.Sort("[ReceivedTime]", True)
            ui_block_add(block_id, f"📧 לאחר מיון, נמצאו {messages.Count} מיילים", "INFO")
            
            # בדיקה מפורטת של המיילים
            if messages.Count > 0:
                ui_block_add(block_id, "🔍 בודק מיילים זמינים...", "INFO")
                
                # נסה לגשת לכמה מיילים במיקומים שונים
                test_indices = [1, messages.Count//2, messages.Count]
                for idx in test_indices:
                    try:
                        if 1 <= idx <= messages.Count:
                            # שימוש בגישה יציבה יותר לאיברים בקולקציית COM
                            test_msg = messages.Item(idx)
                            if test_msg and hasattr(test_msg, 'Subject'):
                                ui_block_add(block_id, f"✅ מייל {idx}: {test_msg.Subject[:30]}...", "INFO")
                            else:
                                ui_block_add(block_id, f"⚠️ מייל {idx}: לא תקין", "WARNING")
                    except Exception as e:
                        # לעתים Outlook מחזיר שגיאת אינדקס – זו אינה קריטית, משנים לאזהרה
                        ui_block_add(block_id, f"⚠️ מייל {idx}: בעיה בגישה ( {e} )", "WARNING")
                
                ui_block_add(block_id, "✅ בדיקת מיילים הושלמה", "SUCCESS")
            
            # בדיקה מהירה של מספר המיילים הזמינים
            try:
                # נסה לגשת לכמה מיילים לדוגמה כדי לוודא שהגישה עובדת
                test_count = min(3, messages.Count)
                for i in range(1, test_count + 1):
                    try:
                        message = messages.Item(i)
                        if message:
                            ui_block_add(block_id, f"✅ מייל {i}: {message.Subject[:50]}...", "INFO")
                    except Exception as e:
                        ui_block_add(block_id, f"⚠️ בעיה בגישה למייל {i}: {e}", "WARNING")
                        break
                ui_block_add(block_id, f"✅ בדיקת גישה הושלמה - {messages.Count} מיילים זמינים", "SUCCESS")
            except Exception as e:
                ui_block_add(block_id, f"ERROR שגיאה בבדיקת גישה: {e}", "ERROR")
                ui_block_end(block_id, f"שגיאה בבדיקת גישה: {e}", False)
                return []

            ui_block_add(block_id, "📧 מתחיל טעינת מיילים מ-Outlook...", "INFO")

            emails = []
            # שימוש בלולאת foreach יציבה יותר מאשר גישה עם אינדקס
            for i, message in enumerate(messages):
                try:
                    if message is None:
                        log_to_console(f"⚠️ מייל {i+1} הוא None - מדלג", "WARNING")
                        continue

                    # בדיקה שהמייל הוא באמת מייל
                    if not hasattr(message, 'Subject'):
                        log_to_console(f"⚠️ מייל {i+1} אינו מייל תקין - מדלג", "WARNING")
                        continue

                    email_data = {
                        'id': i + 1,
                        'subject': str(message.Subject) if message.Subject else "ללא נושא",
                        'sender': str(message.SenderName) if message.SenderName else "שולח לא ידוע",
                        'sender_email': str(message.SenderEmailAddress) if message.SenderEmailAddress else "",
                        'received_time': message.ReceivedTime, # שמירת אובייקט datetime למיון
                        'body_preview': self._clean_email_body(message.Body),
                        'is_read': not message.UnRead
                    }

                    # ניתוח מהיר ללא AI - רק נתונים בסיסיים
                    email_data['summary'] = f"מייל מ-{email_data['sender']}: {email_data['subject']}"
                    email_data['action_items'] = []
                    
                    # ניתוח בסיסי של חשיבות
                    email_data['importance_score'] = self.calculate_smart_importance(email_data)
                    email_data['original_importance_score'] = email_data['importance_score']
                    email_data['category'] = self.categorize_smart(email_data)
                    
                    # לא שומרים ניתוח חכם לבסיס נתונים - רק ניתוח AI אמיתי

                    emails.append(email_data)

                    if (i + 1) % 50 == 0:
                        ui_block_add(block_id, f"Loaded {i + 1} emails...", "INFO")

                    if limit and len(emails) >= limit:
                        ui_block_add(block_id, f"Reached loading limit of {limit} emails.", "WARNING")
                        break
                except Exception as e:
                    ui_block_add(block_id, f"Error in email {i+1}: {e}", "ERROR")
                    continue

            # מיון המיילים לאחר הטעינה
            emails.sort(key=lambda x: x['received_time'], reverse=True)
            # המרת התאריך למחרוזת לאחר המיון
            for email in emails:
                email['received_time'] = str(email['received_time'])

            ui_block_end(block_id, f"טעינת {len(emails)} מיילים הושלמה ומויינה", True)
            return emails
            
        except Exception as e:
            try:
                ui_block_end(block_id, f"שגיאה בטעינת מיילים: {e}", False)
            except Exception:
                log_to_console(f"Error getting emails from Outlook: {e}", "ERROR")
            self.outlook_connected = False
            return []
        finally:
            # ניקוי COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def get_sample_emails(self):
        """קבלת נתונים דמה"""
        sample_emails = [
            {
                'id': 1,
                'subject': 'Upgrade now to reactivate your Azure free account',
                'sender': 'Microsoft Azure',
                'sender_email': 'noreply@azure.com',
                'received_time': str(datetime.now() - timedelta(hours=2)),
                'body_preview': 'Your Azure free account has expired. Upgrade now to continue using our services...',
                'importance_score': 0.9,
                'category': 'system',
                'summary': 'הודעה על פג תוקף חשבון Azure - נדרש שדרוג',
                'action_items': ['שדרג את חשבון Azure', 'בדוק את השירותים הפעילים'],
                'is_read': False
            },
            {
                'id': 2,
                'subject': 'Meeting tomorrow at 10:00 AM',
                'sender': 'Sarah Johnson',
                'sender_email': 'sarah.johnson@company.com',
                'received_time': str(datetime.now() - timedelta(hours=5)),
                'body_preview': 'Hi, just a reminder about our meeting tomorrow at 10:00 AM. Please bring the quarterly reports...',
                'importance_score': 0.8,
                'category': 'meeting',
                'summary': 'תזכורת לפגישה מחר ב-10:00 - להביא דוחות רבעוניים',
                'action_items': ['הכן דוחות רבעוניים', 'הגע לפגישה ב-10:00'],
                'is_read': True
            },
            {
                'id': 3,
                'subject': 'חשוב: עדכון מדיניות החברה',
                'sender': 'מחלקת משאבי אנוש',
                'sender_email': 'hr@company.co.il',
                'received_time': str(datetime.now() - timedelta(days=1)),
                'body_preview': 'שלום, אנחנו מעדכנים אתכם על שינויים במדיניות החברה. אנא קראו את הקובץ המצורף...',
                'importance_score': 0.7,
                'category': 'work',
                'summary': 'עדכון מדיניות החברה - נדרש קריאה',
                'action_items': ['קרא את המדיניות החדשה', 'אשר הבנת את השינויים'],
                'is_read': False
            }
        ]
        return sample_emails
    
# פונקציה כפולה הוסרה - משתמשים בפונקציה הראשונה
    
    def analyze_emails_smart(self, emails, block_id=None):
        """ניתוח חכם מבוסס פרופיל משתמש - עיבוד מהיר, עטוף כבלוק שרת יחיד"""
        created_block = False
        try:
            if not block_id:
                block_id = ui_block_start("🧠 ניתוח חכם של מיילים")
                created_block = True
                ui_block_add(block_id, f"Starting smart analysis of {len(emails)} emails", "INFO")
            else:
                ui_block_add(block_id, f"Starting smart analysis of {len(emails)} emails", "INFO")
            
            ui_block_add(block_id, "Smart logic: time, content, sender, categories and tasks analysis", "INFO")
            
            for i, email in enumerate(emails):
                # ניתוח חכם מבוסס פרופיל - רק אם לא נותח בעבר
                if not email.get('ai_analyzed', False):
                    email['importance_score'] = self.calculate_smart_importance(email)
                    email['category'] = self.categorize_smart(email)
                    email['summary'] = self.generate_smart_summary(email)
                    email['action_items'] = self.extract_smart_action_items(email)
                    # שמירת הציון המקורי
                    if 'original_importance_score' not in email:
                        email['original_importance_score'] = email['importance_score']
                    # לא מסמנים כ-ai_analyzed כאן - רק ניתוח AI אמיתי
                else:
                    # אם כבר נותח, נשמור את הציון המקורי אם לא קיים
                    if 'original_importance_score' not in email:
                        email['original_importance_score'] = email.get('importance_score', 0.5)
                
                # התקדמות כל 100 מיילים
                if (i + 1) % 100 == 0:
                    ui_block_add(block_id, f"🧠 ניתח {i + 1}/{len(emails)} מיילים...", "INFO")
            
            ui_block_end(block_id, f"Completed smart analysis of {len(emails)} emails", True)
            return emails
        except Exception as e:
            # סגירת בלוק במקרה של שגיאה
            try:
                ui_block_end(block_id, f"שגיאה בניתוח חכם: {str(e)}", False)
            except Exception:
                log_to_console(f"שגיאה בניתוח חכם: {str(e)}", "ERROR")
            return emails
    
    def calculate_smart_importance(self, email):
        """חישוב חשיבות חכם מתקדם - מערכת ניתוח מקיפה"""
        score = 0.10  # ציון בסיסי נמוך - רוב המיילים יהיו נמוכים
        
        # 1. ניתוח תוכן מתקדם
        subject = str(email.get('subject', '')).lower()
        body = str(email.get('body_preview', '')).lower()
        full_text = f"{subject} {body}"
        
        # ניתוח מילות דחיפות מתקדם
        urgency_patterns = {
            'critical': ['urgent', 'critical', 'emergency', 'asap', 'immediately', 'דחוף', 'חשוב', 'דחוף מאוד'],
            'deadline': ['deadline', 'due date', 'תאריך יעד', 'לפני', 'עד', 'by', 'until'],
            'exclamation': ['!!!', '???', '!!', '??', '!', '?'],
            'priority': ['priority', 'high priority', 'low priority', 'עדיפות', 'עדיפות גבוהה', 'עדיפות נמוכה']
        }
        
        urgency_score = 0
        for category, keywords in urgency_patterns.items():
            count = sum(1 for keyword in keywords if keyword in full_text)
            if category == 'critical':
                urgency_score += count * 0.15  # קטן יותר
            elif category == 'deadline':
                urgency_score += count * 0.12  # קטן יותר
            elif category == 'exclamation':
                urgency_score += count * 0.08  # קטן יותר
            elif category == 'priority':
                urgency_score += count * 0.10  # קטן יותר
        
        score += min(urgency_score, 0.20)  # מקסימום 0.20 לדחיפות
        
        # ניתוח סנטימנט
        positive_words = ['thanks', 'thank you', 'great', 'excellent', 'good', 'תודה', 'מעולה', 'טוב', 'נהדר']
        negative_words = ['problem', 'issue', 'error', 'bug', 'complaint', 'בעיה', 'שגיאה', 'תלונה', 'קושי']
        
        positive_count = sum(1 for word in positive_words if word in full_text)
        negative_count = sum(1 for word in negative_words if word in full_text)
        
        if negative_count > positive_count:
            score += 0.10  # מיילים שליליים = חשובים יותר (בעיות לפתור)
        elif positive_count > negative_count:
            score += 0.03  # מיילים חיוביים = חשובים פחות
        
        # ניתוח שאלות ישירות
        question_indicators = ['?', 'מה', 'איך', 'מתי', 'איפה', 'למה', 'מי', 'what', 'how', 'when', 'where', 'why', 'who']
        question_count = sum(1 for indicator in question_indicators if indicator in full_text)
        score += min(question_count * 0.05, 0.12)  # מקסימום 0.12 לשאלות
        
        # ניתוח אורך מייל
        body_length = len(str(email.get('body_preview', '')))
        if body_length > 1000:  # מיילים ארוכים מאוד
            score += 0.08  # קטן יותר
        elif body_length > 500:  # מיילים ארוכים
            score += 0.05  # קטן יותר
        elif body_length < 50:  # מיילים קצרים מאוד
            score -= 0.03  # קטן יותר
        
        # 2. ניתוח שולח מתקדם
        sender = str(email.get('sender', '')).lower()
        sender_email = str(email.get('sender_email', '')).lower()
        
        # היררכיה ארגונית מתקדמת
        hierarchy_titles = {
            'ceo_level': ['ceo', 'מנכ"ל', 'president', 'נשיא'],
            'c_level': ['cto', 'cfo', 'coo', 'cmo', 'סמנכ"ל', 'מנהל כללי'],
            'director': ['director', 'מנהל', 'head of', 'ראש'],
            'manager': ['manager', 'מנהל', 'supervisor', 'מפקח']
        }
        
        for level, titles in hierarchy_titles.items():
            if any(title in sender for title in titles):
                if level == 'ceo_level':
                    score += 0.20  # קטן יותר
                elif level == 'c_level':
                    score += 0.15  # קטן יותר
                elif level == 'director':
                    score += 0.12  # קטן יותר
                elif level == 'manager':
                    score += 0.08  # קטן יותר
                break
        
        # ניתוח דומיין מתקדם
        domain_analysis = {
            'internal': ['@company.com', '@internal.com', '@corp.com'],
            'clients': ['@client.com', '@customer.com', '@partner.com'],
            'vendors': ['@vendor.com', '@supplier.com', '@service.com'],
            'personal': ['@gmail.com', '@yahoo.com', '@hotmail.com', '@outlook.com']
        }
        
        for domain_type, domains in domain_analysis.items():
            if any(domain in sender_email for domain in domains):
                if domain_type == 'internal':
                    score += 0.10  # קטן יותר
                elif domain_type == 'clients':
                    score += 0.12  # קטן יותר
                elif domain_type == 'vendors':
                    score += 0.06  # קטן יותר
                elif domain_type == 'personal':
                    score += 0.03  # קטן יותר
                break
        
        # 3. ניתוח זמן מתקדם
        try:
            received_time = email.get('received_time')
            if received_time:
                if isinstance(received_time, str):
                    from datetime import datetime
                    received_time = datetime.fromisoformat(received_time.replace('Z', '+00:00'))
                
                    # ניתוח שעות עבודה
                    hour = received_time.hour
                    if 9 <= hour <= 17:  # שעות עבודה
                        score += 0.06  # קטן יותר
                    elif 18 <= hour <= 22:  # שעות ערב
                        score += 0.03  # קטן יותר
                    else:  # שעות לילה/בוקר מוקדם
                        score += 0.08  # מיילים בשעות לא רגילות = חשובים יותר
                    
                    # ניתוח ימי שבוע
                    weekday = received_time.weekday()  # 0=Monday, 6=Sunday
                    if weekday < 5:  # ימי חול
                        score += 0.03  # קטן יותר
                    else:  # סוף שבוע
                        score += 0.06  # מיילים בסוף שבוע = חשובים יותר
                    
                    # ניתוח זמן תגובה
                    time_diff = datetime.now() - received_time
                    if time_diff.days < 1:
                        score += 0.08  # מיילים מהיום
                    elif time_diff.days < 3:
                        score += 0.05  # מיילים מ-3 ימים
                    elif time_diff.days < 7:
                        score += 0.03  # מיילים משבוע
                    else:
                        score -= 0.03  # מיילים ישנים
        except:
            pass
        
        # 4. ניתוח התנהגותי
        # קבצים מצורפים
        if email.get('has_attachments', False):
            score += 0.06  # קטן יותר
        
        # CC/BCC
        if email.get('cc', '') or email.get('bcc', ''):
            score += 0.05  # קטן יותר
        
        # תגובות
        if 're:' in subject.lower():
            score += 0.03  # קטן יותר
        if 'fwd:' in subject.lower():
            score += 0.02  # קטן יותר
        
        # לינקים
        if 'http' in body or 'www.' in body:
            score += 0.02  # קטן יותר
        
        # 5. בדיקת סטטוס קריאה
        if not email.get('is_read', False):
            score += 0.06  # מיילים שלא נקראו
        
        # 6. בדיקת פרופיל משתמש
        sender_importance = self.profile_manager.get_sender_importance(email['sender'])
        score += sender_importance * 0.1  # קטן יותר
        
        important_keywords = self.profile_manager.get_important_keywords()
        for keyword, weight in important_keywords.items():
            if keyword.lower() in subject:
                score += weight * 0.08  # קטן יותר
            if keyword.lower() in body:
                score += weight * 0.05  # קטן יותר
        
        category_importance = self.profile_manager.get_category_importance(email.get('category', 'work'))
        score += category_importance * 0.08  # קטן יותר
        
        # 7. ניתוח קטגוריה
        category = email.get('category', 'work')
        category_scores = {
            'urgent': 0.15,  # קטן יותר
            'meeting': 0.12,  # קטן יותר
            'project': 0.08,  # קטן יותר
            'admin': 0.05,   # קטן יותר
            'finance': 0.08,  # קטן יותר
            'legal': 0.12,   # קטן יותר
            'support': 0.06,  # קטן יותר
            'marketing': 0.04, # קטן יותר
            'personal': 0.02  # קטן יותר
        }
        score += category_scores.get(category, 0.03)  # קטן יותר
        
        # 8. ניתוח מיילים מ-Microsoft/Azure (ציון מופחת)
        if any(company in sender for company in ['microsoft', 'azure', 'office', 'outlook', 'teams']):
            score += 0.01  # ציון נמוך מאוד
        
        final_score = min(max(score, 0.0), 1.0)  # הגבלה בין 0 ל-1
        return final_score
    
    def analyze_single_email(self, email_data):
        """ניתוח מייל בודד"""
        try:
            # ניתוח בסיסי
            importance_score = self.calculate_smart_importance(email_data)
            category = self.categorize_smart(email_data)
            
            # ניתוח AI אם זמין
            if self.ai_analyzer and self.ai_analyzer.is_ai_available():
                try:
                    ai_analysis = self.ai_analyzer.analyze_email_importance(email_data)
                    ai_category = self.ai_analyzer.categorize_email(email_data)
                    summary = self.ai_analyzer.summarize_email(email_data)
                    action_items = self.ai_analyzer.extract_action_items(email_data)
                    
                    # שילוב עם למידה מותאמת אישית
                    if self.profile_manager:
                        learned_importance = self.profile_manager.get_personalized_importance_score(email_data)
                        learned_category = self.profile_manager.get_personalized_category(email_data)
                        
                        # ממוצע משוקלל בין AI ולמידה
                        final_importance = (ai_analysis * 0.7 + learned_importance * 0.3)
                        final_category = learned_category if learned_category != 'work' else ai_category
                    else:
                        final_importance = ai_analysis
                        final_category = ai_category
                    
                    return {
                        'importance_score': final_importance,
                        'category': final_category,
                        'summary': summary,
                        'action_items': action_items,
                        'ai_analyzed': False,  # ניתוח חכם, לא AI
                        'original_importance_score': importance_score,
                        'ai_importance_score': ai_analysis,
                        'ai_category': ai_category
                    }
                    
                except Exception as e:
                    print(f"AI analysis failed: {e}")
                    # fallback לניתוח בסיסי
            
            # ניתוח בסיסי בלבד
            summary = f"מייל מ-{email_data.get('sender', 'לא ידוע')}: {email_data.get('subject', 'ללא נושא')}"
            
            return {
                'importance_score': importance_score,
                'category': category,
                'summary': summary,
                'action_items': [],
                'ai_analyzed': False
            }
            
        except Exception as e:
            print(f"Error analyzing email: {e}")
            return {
                'importance_score': 0.5,
                'category': 'work',
                'summary': 'שגיאה בניתוח המייל',
                'action_items': [],
                'ai_analyzed': False
            }

    def categorize_smart(self, email):
        """קטגוריזציה חכמה מבוסס פרופיל + לוגיקה חכמה"""
        subject = str(email.get('subject', '')).lower()
        sender = str(email.get('sender', '')).lower()
        body = str(email.get('body_preview', '')).lower()
        
        # בדיקה מהפרופיל
        learned_category = self.profile_manager.get_personalized_category(email)
        if learned_category and learned_category != 'work':
            return learned_category
        
        # קטגוריזציה חכמה משופרת
        # 1. דחיפות גבוהה
        if any(word in subject for word in ['urgent', 'דחוף', 'asap', 'critical', 'חשוב', '!!!', '???']):
            return 'urgent'
        
        # 2. פגישות
        if any(word in subject for word in ['meeting', 'פגישה', 'call', 'שיחה', 'zoom', 'teams', 'calendar']):
            return 'meeting'
        
        # 3. דוחות וסיכומים
        if any(word in subject for word in ['report', 'דוח', 'summary', 'סיכום', 'analytics', 'dashboard']):
            return 'report'
        
        # 4. פרויקטים ומשימות
        if any(word in subject for word in ['project', 'פרויקט', 'task', 'משימה', 'milestone', 'deadline']):
            return 'project'
        
        # 5. משאבי אנוש ומנהלה
        if any(word in sender for word in ['hr', 'משאבי אנוש', 'admin', 'מנהל', 'payroll', 'benefits']):
            return 'admin'
        
        # 6. IT ותמיכה טכנית
        if any(word in subject for word in ['support', 'תמיכה', 'bug', 'error', 'issue', 'technical']):
            return 'support'
        
        # 7. מכירות ושיווק
        if any(word in subject for word in ['sale', 'מכירה', 'marketing', 'שיווק', 'promotion', 'offer']):
            return 'marketing'
        
        # 8. כספים וחשבונות
        if any(word in subject for word in ['invoice', 'חשבונית', 'payment', 'תשלום', 'budget', 'תקציב']):
            return 'finance'
        
        # 9. משפטי
        if any(word in subject for word in ['legal', 'משפטי', 'contract', 'חוזה', 'agreement', 'הסכם']):
            return 'legal'
        
        # 10. פרסומות וספאם
        if any(word in subject for word in ['unsubscribe', 'הסרה', 'promotion', 'discount', 'sale', 'offer']):
            return 'marketing'
        
        # 11. מיילים אישיים
        if any(word in sender for word in ['gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com']):
            return 'personal'
        
        # 12. ברירת מחדל
        return 'work'
    
    def generate_smart_summary(self, email):
        """יצירת סיכום חכם"""
        subject = email.get('subject', '')
        sender = email.get('sender', '')
        category = email.get('category', 'work')
        
        if category == 'meeting':
            return f"פגישה: {subject} מ-{sender}"
        elif category == 'urgent':
            return f"דחוף: {subject} מ-{sender}"
        elif category == 'report':
            return f"דוח: {subject} מ-{sender}"
        elif category == 'project':
            return f"פרויקט: {subject} מ-{sender}"
        else:
            return f"מייל מ-{sender}: {subject}"
    
    def extract_smart_action_items(self, email):
        """חילוץ משימות חכם משופר"""
        subject = str(email.get('subject', '')).lower()
        body = str(email.get('body_preview', '')).lower()
        category = email.get('category', 'work')
        
        action_items = []
        
        # חיפוש מילות מפתח של משימות
        if any(word in subject for word in ['review', 'בדוק', 'check', 'verify', 'אמת']):
            action_items.append("בדוק את התוכן")
        
        if any(word in subject for word in ['reply', 'תגובה', 'respond', 'ענה']):
            action_items.append("הגב למייל")
        
        if any(word in subject for word in ['meeting', 'פגישה', 'call', 'שיחה']):
            action_items.append("הכן לפגישה")
        
        if any(word in body for word in ['deadline', 'תאריך יעד', 'due date']):
            action_items.append("בדוק תאריך יעד")
        
        # משימות ספציפיות לקטגוריות
        if category == 'urgent':
            action_items.append("טפל בדחיפות")
        
        if category == 'meeting':
            action_items.append("הכן לפגישה")
            action_items.append("בדוק זמינות")
        
        if category == 'project':
            action_items.append("עדכן סטטוס פרויקט")
        
        if category == 'report':
            action_items.append("קרא דוח")
            action_items.append("סכם נקודות עיקריות")
        
        if category == 'admin':
            action_items.append("טפל בבקשה מנהלית")
        
        if category == 'support':
            action_items.append("טפל בבעיה טכנית")
        
        if category == 'finance':
            action_items.append("בדוק חשבונית")
            action_items.append("אשר תשלום")
        
        if category == 'legal':
            action_items.append("בדוק חוזה")
            action_items.append("התייעץ עם עורך דין")
        
        # משימות כלליות
        if any(word in body for word in ['action', 'פעולה', 'task', 'משימה']):
            action_items.append("בצע פעולה נדרשת")
        
        if any(word in body for word in ['approve', 'אשר', 'confirm', 'אמת']):
            action_items.append("אשר בקשה")
        
        if any(word in body for word in ['schedule', 'תזמן', 'book', 'הזמן']):
            action_items.append("תזמן פגישה")
        
        # הגבלת מספר המשימות
        return action_items[:3]  # מקסימום 3 משימות
    
    def calculate_basic_importance(self, email_data):
        """חישוב בסיסי של חשיבות (fallback)"""
        score = 0.5
        
        try:
            # בדיקת מילות מפתח חשובות
            important_keywords = ['חשוב', 'דחוף', 'urgent', 'important', 'meeting', 'פגישה', 'azure', 'microsoft', 'security', 'alert']
            subject = str(email_data.get('subject', '')).lower()
            body = str(email_data.get('body_preview', '')).lower()
            
            for keyword in important_keywords:
                if keyword in subject:
                    score += 0.2
                if keyword in body:
                    score += 0.1
            
            # בדיקת שולח חשוב
            important_senders = ['manager', 'boss', 'מנהל', 'hr', 'it', 'microsoft', 'azure', 'security', 'admin']
            sender = str(email_data.get('sender', '')).lower()
            
            for important_sender in important_senders:
                if important_sender in sender:
                    score += 0.3
            
            # בדיקת זמן - מיילים חדשים יותר חשובים
            try:
                received_time = email_data.get('received_time')
                if received_time:
                    if hasattr(received_time, 'replace'):
                        received_time = received_time.replace(tzinfo=None)
                    elif isinstance(received_time, str):
                        received_time = datetime.fromisoformat(received_time.replace('Z', '+00:00'))
                    
                    time_diff = datetime.now() - received_time
                    if time_diff.days < 1:
                        score += 0.2
                    elif time_diff.days < 7:
                        score += 0.1
            except Exception as e:
                log_to_console(f"שגיאה בחישוב זמן: {e}", "ERROR")
                pass
            
        except Exception as e:
            log_to_console(f"שגיאה בחישוב חשיבות: {e}", "ERROR")
        
        return min(score, 1.0)  # מקסימום 1.0
    
    def calculate_importance_score(self, message):
        """חישוב ציון חשיבות למייל"""
        score = 0.5  # ציון בסיסי
        
        try:
            # בדיקת מילות מפתח חשובות
            important_keywords = ['חשוב', 'דחוף', 'urgent', 'important', 'meeting', 'פגישה', 'azure', 'microsoft', 'security', 'alert']
            subject = str(message.Subject).lower() if message.Subject else ""
            body = str(message.Body).lower() if message.Body else ""
            
            for keyword in important_keywords:
                if keyword in subject:
                    score += 0.2
                if keyword in body:
                    score += 0.1
            
            # בדיקת שולח חשוב
            important_senders = ['manager', 'boss', 'מנהל', 'hr', 'it', 'microsoft', 'azure', 'security', 'admin']
            sender = str(message.SenderName).lower() if message.SenderName else ""
            
            for important_sender in important_senders:
                if important_sender in sender:
                    score += 0.3
            
            # בדיקת זמן - מיילים חדשים יותר חשובים
            try:
                received_time = message.ReceivedTime
                if hasattr(received_time, 'replace'):
                    # המרה ל-naive datetime
                    received_time = received_time.replace(tzinfo=None)
                    
                    time_diff = datetime.now() - received_time
                    if time_diff.days < 1:
                        score += 0.2
                    elif time_diff.days < 7:
                        score += 0.1
            except Exception as e:
                log_to_console(f"שגיאה בחישוב זמן: {e}", "ERROR")
                pass
            
        except Exception as e:
            log_to_console(f"שגיאה בחישוב חשיבות: {e}", "ERROR")
        
        return min(score, 1.0)  # מקסימום 1.0
    
    def save_user_preference(self, preference_type, preference_value, weight=1.0):
        """שמירת העדפת משתמש"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO user_preferences (preference_type, preference_value, weight)
            VALUES (?, ?, ?)
        ''', (preference_type, preference_value, weight))
        
        conn.commit()
        conn.close()
        
        # עדכון זיכרון
        if preference_type not in self.user_preferences:
            self.user_preferences[preference_type] = []
        self.user_preferences[preference_type].append({
            'value': preference_value,
            'weight': weight
        })
    
    def load_user_preferences(self):
        """טעינת העדפות משתמש"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('SELECT preference_type, preference_value, weight FROM user_preferences')
            rows = cursor.fetchall()
            
            for row in rows:
                pref_type, pref_value, weight = row
                if pref_type not in self.user_preferences:
                    self.user_preferences[pref_type] = []
                self.user_preferences[pref_type].append({
                    'value': pref_value,
                    'weight': weight
                })
            
            conn.close()
        except Exception as e:
            log_to_console(f"שגיאה בטעינת העדפות: {e}", "ERROR")

    def connect_to_outlook(self):
        """חיבור ל-Outlook"""
        try:
            log_to_console("Trying to connect to Outlook...", "INFO")
            
            # נסה חיבור עם הרשאות נמוכות יותר
            try:
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                log_to_console("✅ חיבור ל-Outlook Application הצליח!", "SUCCESS")
            except Exception as outlook_error:
                log_to_console(f"Error connecting to Outlook Application: {outlook_error}", "ERROR")
                raise outlook_error
            
            # נסה חיבור ל-Namespace
            try:
                self.namespace = self.outlook.GetNamespace("MAPI")
                log_to_console("✅ חיבור ל-Namespace הצליח!", "SUCCESS")
            except Exception as namespace_error:
                log_to_console(f"Error connecting to Namespace: {namespace_error}", "ERROR")
                raise namespace_error
            
            # בדיקה שהחיבור עובד
            try:
                # נסה גישה בסיסית
                test_folder = self.namespace.GetDefaultFolder(6)  # Inbox
                log_to_console("Basic connection test successful!", "SUCCESS")
            except Exception as test_error:
                log_to_console(f"Basic connection test failed: {test_error}", "WARNING")
            
            self.outlook_connected = True
            log_to_console("Outlook connection successful!", "SUCCESS")
            return True
        except Exception as e:
            log_to_console(f"Error connecting to Outlook: {e}", "ERROR")
            self.outlook_connected = False
            self.outlook = None
            self.namespace = None
            return False

    def get_meetings(self):
        """קבלת כל הפגישות מ-Outlook"""
        meetings = []
        
        # יצירת בלוק לטעינת פגישות
        block_id = ui_block_start("📅 טעינת פגישות מ-Outlook")
        
        try:
            ui_block_add(block_id, "מתחיל טעינת פגישות...", "INFO")
            
            # יצירת חיבור חדש בכל קריאה כדי למנוע בעיות threading
            try:
                ui_block_add(block_id, "🔌 יוצר חיבור חדש ל-Outlook...", "INFO")
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                ui_block_add(block_id, "✅ חיבור הצליח!", "SUCCESS")
            except Exception as connection_error:
                ui_block_add(block_id, f"❌ שגיאה בחיבור: {connection_error}", "ERROR")
                ui_block_end(block_id, "החיבור ל-Outlook נכשל", False)
                raise connection_error
            
            ui_block_add(block_id, f"Outlook object: {outlook is not None}", "INFO")
            ui_block_add(block_id, f"Namespace object: {namespace is not None}", "INFO")
            
            if outlook and namespace:
                ui_block_add(block_id, "✅ Outlook מחובר - טוען פגישות...", "SUCCESS")
                # קבלת הפגישות מהלוח שנה
                calendar = None
                appointments = None
                
                try:
                    ui_block_add(block_id, "📅 מנסה לגשת ללוח השנה...", "INFO")
                    # נסה גישה ללוח השנה
                    calendar = namespace.GetDefaultFolder(9)  # olFolderCalendar
                    ui_block_add(block_id, "✅ גישה ללוח השנה הצליחה!", "SUCCESS")
                    appointments = calendar.Items
                    appointments.Sort("[Start]")
                except Exception as calendar_error:
                    ui_block_add(block_id, f"❌ שגיאה בגישה ללוח השנה: {calendar_error}", "ERROR")
                    # נסה דרך חשבונות Outlook עם הרשאות נמוכות יותר
                    try:
                        ui_block_add(block_id, "📅 מנסה דרך חשבונות Outlook...", "INFO")
                        
                        # נסה גישה ישירה לחשבונות
                        try:
                            accounts = namespace.Accounts
                            ui_block_add(block_id, f"📧 נמצאו {accounts.Count} חשבונות", "INFO")
                        except Exception as accounts_error:
                            ui_block_add(block_id, f"❌ שגיאה בגישה לחשבונות: {accounts_error}", "ERROR")
                            # נסה דרך אחרת - דרך תיקיות ישירות
                            try:
                                ui_block_add(block_id, "📅 מנסה דרך תיקיות ישירות...", "INFO")
                                folders = namespace.Folders
                                ui_block_add(block_id, f"📁 נמצאו {folders.Count} תיקיות", "INFO")
                                
                                for i in range(1, folders.Count + 1):
                                    try:
                                        folder = folders.Item(i)
                                        ui_block_add(block_id, f"📁 תיקייה {i}: {folder.Name}", "INFO")
                                        
                                        # נסה למצוא תיקיית לוח שנה
                                        if "Calendar" in folder.Name or "לוח שנה" in folder.Name or "תאריכים" in folder.Name:
                                            calendar = folder
                                            appointments = calendar.Items
                                            appointments.Sort("[Start]")
                                            ui_block_add(block_id, f"✅ גישה ללוח השנה דרך תיקייה {folder.Name} הצליחה!", "SUCCESS")
                                            break
                                        
                                        # נסה לחפש תיקיות משנה
                                        try:
                                            sub_folders = folder.Folders
                                            ui_block_add(block_id, f"📁 נמצאו {sub_folders.Count} תיקיות משנה ב-{folder.Name}", "INFO")
                                            
                                            for j in range(1, sub_folders.Count + 1):
                                                try:
                                                    sub_folder = sub_folders.Item(j)
                                                    ui_block_add(block_id, f"📁 תיקיית משנה {j}: {sub_folder.Name}", "INFO")
                                                    if "Calendar" in sub_folder.Name or "לוח שנה" in sub_folder.Name or "תאריכים" in sub_folder.Name:
                                                        calendar = sub_folder
                                                        appointments = calendar.Items
                                                        appointments.Sort("[Start]")
                                                        ui_block_add(block_id, f"✅ גישה ללוח השנה דרך תיקיית משנה {sub_folder.Name} הצליחה!", "SUCCESS")
                                                        break
                                                except Exception as sub_folder_error:
                                                    ui_block_add(block_id, f"⚠️ שגיאה בתיקיית משנה {j}: {sub_folder_error}", "WARNING")
                                                    continue
                                            else:
                                                continue  # לא נמצא לוח שנה בתיקייה זו
                                        except Exception as sub_folders_error:
                                            ui_block_add(block_id, f"⚠️ שגיאה בגישה לתיקיות משנה: {sub_folders_error}", "WARNING")
                                            continue
                                    except Exception as folder_error:
                                        ui_block_add(block_id, f"⚠️ שגיאה בתיקייה {i}: {folder_error}", "WARNING")
                                        continue
                                else:
                                    raise Exception("לא נמצא לוח שנה באף תיקייה")
                            except Exception as folders_error:
                                ui_block_add(block_id, f"❌ שגיאה בגישה דרך תיקיות: {folders_error}", "ERROR")
                                raise Exception("לא ניתן לגשת ללוח השנה")
                        
                        # אם הגענו לכאן, נסה דרך חשבונות
                        for i in range(1, accounts.Count + 1):
                            try:
                                account = accounts.Item(i)
                                ui_block_add(block_id, f"📧 חשבון {i}: {account.DisplayName}", "INFO")
                                
                                # נסה לגשת ללוח השנה של החשבון
                                store = account.DeliveryStore
                                if store:
                                    root_folder = store.GetRootFolder()
                                    ui_block_add(block_id, f"📁 תיקיית שורש: {root_folder.Name}", "INFO")
                                    
                                    # נסה למצוא תיקיית לוח שנה
                                    try:
                                        calendar_folder = root_folder.Folders.Item("Calendar")
                                        if calendar_folder:
                                            calendar = calendar_folder
                                            appointments = calendar.Items
                                            appointments.Sort("[Start]")
                                            ui_block_add(block_id, f"✅ גישה ללוח השנה דרך חשבון {account.DisplayName} הצליחה!", "SUCCESS")
                                            break
                                    except Exception as calendar_folder_error:
                                        ui_block_add(block_id, f"⚠️ לא נמצא לוח שנה בחשבון {account.DisplayName}: {calendar_folder_error}", "WARNING")
                                        continue
                            except Exception as account_error:
                                ui_block_add(block_id, f"⚠️ שגיאה בחשבון {i}: {account_error}", "WARNING")
                                continue
                        else:
                            raise Exception("לא נמצא לוח שנה באף חשבון")
                    except Exception as accounts_error:
                        ui_block_add(block_id, f"❌ שגיאה בגישה דרך חשבונות: {accounts_error}", "ERROR")
                        raise Exception("לא ניתן לגשת ללוח השנה")
                
                # בדיקה שיש לנו appointments
                if not appointments:
                    raise Exception("לא ניתן לגשת לפגישות")
                
                ui_block_add(block_id, f"📅 נמצאו {appointments.Count} פגישות ב-Outlook", "INFO")
                
                for appointment in appointments:
                    try:
                        # המרת תאריכים למחרוזות כדי למנוע בעיות serialization
                        def safe_datetime(dt):
                            if dt is None:
                                return None
                            try:
                                if hasattr(dt, 'strftime'):
                                    return dt.strftime('%Y-%m-%d %H:%M:%S')
                                else:
                                    return str(dt)
                            except:
                                return str(dt)
                        
                        meeting_data = {
                            'id': str(appointment.EntryID),
                            'subject': appointment.Subject or 'ללא נושא',
                            'start_time': safe_datetime(appointment.Start),
                            'end_time': safe_datetime(appointment.End),
                            'location': appointment.Location or 'ללא מיקום',
                            'body': appointment.Body or '',
                            'organizer': appointment.Organizer or 'לא ידוע',
                            'attendees': [],
                            'is_recurring': appointment.IsRecurring,
                            'importance': appointment.Importance,
                            'sensitivity': appointment.Sensitivity,
                            'is_all_day': appointment.AllDayEvent,
                            'reminder_minutes': appointment.ReminderMinutesBeforeStart,
                            'created_time': safe_datetime(appointment.CreationTime),
                            'modified_time': safe_datetime(appointment.LastModificationTime)
                        }
                        
                        # קבלת משתתפים
                        if hasattr(appointment, 'Recipients'):
                            for recipient in appointment.Recipients:
                                meeting_data['attendees'].append({
                                    'name': recipient.Name,
                                    'email': recipient.Address,
                                    'type': recipient.Type
                                })
                        
                        meetings.append(meeting_data)
                        
                    except Exception as e:
                        ui_block_add(block_id, f"⚠️ שגיאה בעיבוד פגישה: {e}", "WARNING")
                        continue
                        
                ui_block_add(block_id, f"✅ נטענו {len(meetings)} פגישות מ-Outlook!", "SUCCESS")
            else:
                ui_block_add(block_id, "❌ Outlook לא מחובר", "ERROR")
                ui_block_add(block_id, "📋 משתמש בנתונים דמה", "WARNING")
                meetings = self.get_demo_meetings()
                        
        except Exception as e:
            ui_block_add(block_id, f"❌ שגיאה: {e}", "ERROR")
            ui_block_add(block_id, "📋 משתמש בנתונים דמה", "WARNING")
            # נתונים דמה במקרה של שגיאה
            meetings = self.get_demo_meetings()
        
        # הודעה סופית
        if len(meetings) == 3 and all(meeting.get('id', '').startswith('demo_') for meeting in meetings):
            ui_block_add(block_id, "🚨 אזהרה: המערכת משתמשת בנתונים דמה בלבד!", "ERROR")
            ui_block_add(block_id, "🔧 בדוק את חיבור Outlook או הפעל את Outlook לפני השימוש", "ERROR")
            ui_block_end(block_id, "טעינת פגישות הושלמה (נתונים דמה)", False)
        else:
            ui_block_add(block_id, f"📊 סה\"כ נטענו {len(meetings)} פגישות", "SUCCESS")
            ui_block_end(block_id, "טעינת פגישות הושלמה בהצלחה", True)
        
        return meetings

    def get_demo_meetings(self):
        """נתונים דמה לפגישות"""
        log_to_console("📋 יוצר נתונים דמה לפגישות (3 פגישות לדוגמה)", "INFO")
        demo_meetings = [
            {
                'id': 'demo_1',
                'subject': 'פגישת צוות שבועית',
                'start_time': datetime.now() + timedelta(hours=2),
                'end_time': datetime.now() + timedelta(hours=3),
                'location': 'חדר ישיבות A',
                'body': 'פגישה שבועית לצוות הפיתוח',
                'organizer': 'מנהל הפרויקט',
                'attendees': [
                    {'name': 'רון', 'email': 'ron@company.com', 'type': 'required'},
                    {'name': 'שרה', 'email': 'sarah@company.com', 'type': 'required'},
                    {'name': 'דוד', 'email': 'david@company.com', 'type': 'optional'}
                ],
                'is_recurring': True,
                'importance': 2,
                'sensitivity': 0,
                'is_all_day': False,
                'reminder_minutes': 15,
                'created_time': datetime.now() - timedelta(days=7),
                'modified_time': datetime.now() - timedelta(days=1)
            },
            {
                'id': 'demo_2',
                'subject': 'פגישת לקוח חשובה',
                'start_time': datetime.now() + timedelta(days=1, hours=10),
                'end_time': datetime.now() + timedelta(days=1, hours=11),
                'location': 'משרד הלקוח',
                'body': 'פגישה עם לקוח גדול לדיון על פרויקט חדש',
                'organizer': 'מנהל המכירות',
                'attendees': [
                    {'name': 'רון', 'email': 'ron@company.com', 'type': 'required'},
                    {'name': 'מנהל המכירות', 'email': 'sales@company.com', 'type': 'required'},
                    {'name': 'הלקוח', 'email': 'client@client.com', 'type': 'required'}
                ],
                'is_recurring': False,
                'importance': 2,
                'sensitivity': 1,
                'is_all_day': False,
                'reminder_minutes': 30,
                'created_time': datetime.now() - timedelta(days=3),
                'modified_time': datetime.now() - timedelta(hours=6)
            },
            {
                'id': 'demo_3',
                'subject': 'פגישת סטטוס פרויקט',
                'start_time': datetime.now() + timedelta(days=2, hours=14),
                'end_time': datetime.now() + timedelta(days=2, hours=15),
                'location': 'Zoom',
                'body': 'פגישת סטטוס שבועית לפרויקט החדש',
                'organizer': 'מנהל הפרויקט',
                'attendees': [
                    {'name': 'רון', 'email': 'ron@company.com', 'type': 'required'},
                    {'name': 'צוות הפיתוח', 'email': 'dev@company.com', 'type': 'required'}
                ],
                'is_recurring': True,
                'importance': 1,
                'sensitivity': 0,
                'is_all_day': False,
                'reminder_minutes': 10,
                'created_time': datetime.now() - timedelta(days=14),
                'modified_time': datetime.now() - timedelta(days=2)
            }
        ]
        
        log_to_console(f"📋 נוצרו {len(demo_meetings)} פגישות דמה", "INFO")
        log_to_console("⚠️ שים לב: אתה רואה נתונים דמה ולא פגישות אמיתיות מ-Outlook!", "WARNING")
        return demo_meetings

    def update_meeting_priority(self, meeting_id, priority):
        """עדכון עדיפות פגישה"""
        try:
            # כאן ניתן להוסיף לוגיקה לעדכון העדיפות במסד הנתונים
            # או ב-Outlook עצמו
            
            # שמירה במסד הנתונים המקומי
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # יצירת טבלה לפגישות אם לא קיימת
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS meeting_priorities (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    meeting_id TEXT UNIQUE,
                    priority TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # עדכון או הוספת עדיפות
            cursor.execute('''
                INSERT OR REPLACE INTO meeting_priorities (meeting_id, priority)
                VALUES (?, ?)
            ''', (meeting_id, priority))
            
            conn.commit()
            conn.close()
            
            return True
            
        except Exception as e:
            log_to_console(f"שגיאה בעדכון עדיפות פגישה: {e}", "ERROR")
            return False

# יצירת מופע של מנהל המיילים
email_manager = EmailManager()

@app.route('/')
def index():
    """דף הבית"""
    return render_template('index.html')


@app.route('/consol')
def consol():
    """דף CONSOL - הצגת פלט הקונסול"""
    import time
    # Cache busting - force browser to reload the page
    try:
        log_to_console("🖥️ CONSOL page requested (client loading)", "INFO")
    except Exception:
        pass
    return render_template('consol.html', cache_buster=int(time.time()))

@app.route('/meetings')
def meetings_page():
    """דף ניהול פגישות"""
    return render_template('meetings.html')

@app.route('/api/meetings')
def get_meetings():
    """API לקבלת כל הפגישות מהזיכרון"""
    global cached_data
    
    if cached_data['meetings'] is None:
        log_to_console("📅 אין פגישות בזיכרון - טוען מחדש...", "WARNING")
        refresh_data('meetings')
    
    meetings = cached_data['meetings'] or []
    
    # יישום ניתוח AI שנשמר בבסיס הנתונים
    apply_meeting_ai_analysis_from_db(meetings)
    
    # חישוב ציונים לפי פרופיל המשתמש - תמיד מחדש
    if meetings:
        # יצירת בלוק לטעינת ציוני פגישות
        scores_block_id = ui_block_start(f"📊 חישוב ציונים עבור {len(meetings)} פגישות")
        
        try:
            analyze_meetings_smart(meetings, scores_block_id)
            
            # שמירת הציונים החדשים בבסיס הנתונים
            for meeting in meetings:
                try:
                    save_meeting_ai_analysis_to_db(meeting)
                except Exception as e:
                    ui_block_add(scores_block_id, f"❌ שגיאה בשמירת ציון פגישה: {e}", "ERROR")
            
            ui_block_end(scores_block_id, f"חישוב ציונים הושלם עבור {len(meetings)} פגישות", True)
        except Exception as e:
            ui_block_end(scores_block_id, f"שגיאה בחישוב ציונים: {e}", False)
    
    log_to_console(f"📅 מחזיר {len(meetings)} פגישות מהזיכרון", "INFO")
    return jsonify(meetings)

@app.route('/api/meetings/<meeting_id>/priority', methods=['POST'])
def update_meeting_priority(meeting_id):
    """API לעדכון עדיפות פגישה"""
    try:
        data = request.get_json()
        priority = data.get('priority')
        
        if priority not in ['critical', 'high', 'medium', 'low']:
            return jsonify({'error': 'עדיפות לא חוקית'}), 400
        
        # עדכון העדיפות במערכת
        success = email_manager.update_meeting_priority(meeting_id, priority)
        
        if success:
            # הודעה ברורה ומועילה ללא המספר הלא ברור
            priority_names = {
                'critical': 'קריטי',
                'high': 'חשוב', 
                'medium': 'בינוני',
                'low': 'נמוך'
            }
            priority_hebrew = priority_names.get(priority, priority)
            log_to_console(f"📅 עדיפות פגישה עודכנה ל-{priority_hebrew}", "SUCCESS")
            return jsonify({'success': True, 'message': 'עדיפות עודכנה בהצלחה'})
        else:
            return jsonify({'error': 'לא ניתן לעדכן את העדיפות'}), 500
            
    except Exception as e:
        error_msg = f'שגיאה בעדכון עדיפות: {str(e)}'
        log_to_console(error_msg, "ERROR")
        return jsonify({'error': error_msg}), 500

@app.route('/api/meetings/stats')
def get_meetings_stats():
    """API לקבלת סטטיסטיקות פגישות מהזיכרון"""
    global cached_data
    
    if cached_data['meeting_stats'] is None:
        log_to_console("📊 אין סטטיסטיקות פגישות בזיכרון - מחשב מחדש...", "WARNING")
        refresh_data('meetings')
    
    stats = cached_data['meeting_stats']
    if stats is None:
        # fallback לחישוב מהיר
        meetings = cached_data['meetings'] or []
        total_meetings = len(meetings)
        
        # התפלגות קבועה לפי הדרישות:
        # 10% קריטיים, 20% חשובים, 70% נמוכים
        critical_meetings = int(total_meetings * 0.10)  # 10%
        important_meetings = int(total_meetings * 0.20)  # 20%
        low_meetings = int(total_meetings * 0.70)        # 70%
        
        # סה"כ פגישות = קריטיות + חשובות + נמוכות
        total_categorized_meetings = critical_meetings + important_meetings + low_meetings
        
        # פגישות היום
        today_meetings = len([m for m in meetings if m.get('is_today', False)])
        
        # פגישות השבוע
        week_meetings = len([m for m in meetings if m.get('is_this_week', False)])
        
        log_to_console(f"📊 סטטיסטיקות פגישות: {total_meetings} סה\"כ, {today_meetings} היום, {week_meetings} השבוע", "INFO")
        
        stats = {
            'total_meetings': total_categorized_meetings,  # סה"כ = קריטיות + חשובות + נמוכות
            'critical_meetings': critical_meetings,
            'important_meetings': important_meetings,
            'low_meetings': low_meetings,
            'today_meetings': today_meetings,
            'week_meetings': week_meetings
        }
    
    log_to_console(f"📊 מחזיר סטטיסטיקות פגישות מהזיכרון: {stats['total_meetings']} פגישות", "INFO")
    return jsonify(stats)

@app.route('/api/refresh-data', methods=['POST'])
def refresh_data_api():
    """API לרענון המידע בזיכרון"""
    import time
    start_time = time.time()
    
    try:
        data = request.get_json() or {}
        data_type = data.get('type')  # 'emails', 'meetings', או None לכל הנתונים
        
        success = refresh_data(data_type)
        
        duration = round(time.time() - start_time, 1)
        
        if success:
            response_data = {
                'success': True,
                'message': f'נתונים עודכנו בהצלחה ({data_type or "כל הנתונים"})',
                'last_updated': cached_data['last_updated'].strftime("%H:%M:%S") if cached_data['last_updated'] else None,
                'duration': f'{duration} שניות'
            }
            
            # הוספת סטטיסטיקות לפי סוג
            if data_type == 'emails' or data_type is None:
                response_data['emails_synced'] = len(cached_data.get('emails', []))
            
            if data_type == 'meetings' or data_type is None:
                response_data['meetings_synced'] = len(cached_data.get('meetings', []))
            
            return jsonify(response_data)
        else:
            return jsonify({
                'success': False,
                'message': 'שגיאה ברענון הנתונים'
            }), 500
            
    except Exception as e:
        log_to_console(f"ERROR שגיאה ב-API רענון נתונים: {str(e)}", "ERROR")
        return jsonify({
            'success': False,
            'message': f'שגיאה ברענון נתונים: {str(e)}'
        }), 500

@app.route('/api/summarize-email', methods=['POST'])
def summarize_email_api():
    """API לסיכום מייל בלבד (ללא ציון)"""
    block_id = ui_block_start("🤖 סיכום מייל עם AI")
    
    try:
        email_data = request.json
        
        if not email_data:
            ui_block_add(block_id, "❌ לא נשלחו נתוני מייל", "ERROR")
            ui_block_end(block_id)
            return jsonify({
                'success': False,
                'error': 'לא נשלחו נתוני מייל'
            }), 400
        
        # בניית prompt לסיכום בלבד
        subject = email_data.get('subject', '')
        body = email_data.get('body', '')
        sender = email_data.get('sender_name', email_data.get('sender', ''))
        
        ui_block_add(block_id, f"📧 מייל: {subject[:50]}...", "INFO")
        ui_block_add(block_id, f"👤 שולח: {sender}", "INFO")
        
        # בדיקה אם יש API key
        if not GEMINI_API_KEY or GEMINI_API_KEY == "your-gemini-api-key-here":
            ui_block_add(block_id, "⚠️ אין API key של Gemini", "WARNING")
            ui_block_end(block_id)
            return jsonify({
                'summary': f'סיכום המייל: {subject}',
                'key_points': ['המייל נשלח מ-' + sender],
                'action_items': ['אנא הגדר Gemini API key לסיכום מלא'],
                'sentiment': 'לא זוהה - נדרש API key'
            })
        
        # קריאה ל-AI עם prompt מותאם לסיכום
        prompt = f"""
        סכם את המייל הבא בעברית בצורה תמציתית וברורה:
        
        נושא: {subject}
        שולח: {sender}
        תוכן: {body[:2000]}
        
        אנא ספק בדיוק:
        1. סיכום קצר (2-3 משפטים)
        2. נקודות מרכזיות (רשימה של 2-5 נקודות)
        3. פעולות נדרשות (אם יש)
        4. טון ההודעה
        
        החזר **רק** JSON תקני:
        {{
            "summary": "סיכום המייל כאן",
            "key_points": ["נקודה 1", "נקודה 2"],
            "action_items": ["פעולה 1 אם יש"],
            "sentiment": "פורמלי/לא פורמלי/דחוף"
        }}
        """
        
        try:
            ui_block_add(block_id, "🤖 שולח ל-Gemini AI...", "INFO")
            
            import google.generativeai as genai
            genai.configure(api_key=GEMINI_API_KEY)
            
            # רשימת מודלים לנסות (מהחדש ביותר לישן)
            models_to_try = [
                'gemini-2.5-pro',
                'gemini-2.5-flash', 
                'gemini-2.0-flash',
                'gemini-1.5-pro',
                'gemini-1.5-flash',
                'gemini-pro'
            ]
            
            model = None
            for model_name in models_to_try:
                try:
                    model = genai.GenerativeModel(model_name)
                    ui_block_add(block_id, f"✅ משתמש במודל: {model_name}", "SUCCESS")
                    break
                except Exception as e:
                    continue
            
            if not model:
                raise Exception("לא נמצא מודל Gemini זמין")
            
            response = model.generate_content(prompt)
            result_text = response.text
            
            ui_block_add(block_id, "✅ התקבלה תשובה מ-AI", "SUCCESS")
            
            # ניסיון לחלץ JSON מהתשובה
            import re
            import json
            
            # הסרת markdown code blocks אם יש
            result_text = re.sub(r'```json\s*', '', result_text)
            result_text = re.sub(r'```\s*', '', result_text)
            
            json_match = re.search(r'\{[\s\S]*\}', result_text)
            if json_match:
                result = json.loads(json_match.group())
                ui_block_add(block_id, "📝 סיכום הושלם בהצלחה", "SUCCESS")
            else:
                ui_block_add(block_id, "⚠️ לא נמצא JSON, משתמש בטקסט רגיל", "WARNING")
                result = {
                    'summary': result_text[:500] if result_text else 'לא התקבל סיכום',
                    'key_points': ['הסיכום מופיע בשדה הראשי'],
                    'action_items': [],
                    'sentiment': 'לא זוהה'
                }
            
            ui_block_end(block_id)
            return jsonify(result)
            
        except Exception as ai_error:
            ui_block_add(block_id, f"❌ שגיאת AI: {str(ai_error)[:100]}", "ERROR")
            ui_block_end(block_id)
            return jsonify({
                'success': False,
                'error': f'שגיאה בניתוח AI: {str(ai_error)}'
            }), 500
        
    except Exception as e:
        ui_block_add(block_id, f"❌ שגיאה כללית: {str(e)[:100]}", "ERROR")
        ui_block_end(block_id)
        return jsonify({
            'success': False,
            'error': f'שגיאה כללית: {str(e)}'
        }), 500

@app.route('/api/generate-tasks', methods=['POST'])
def generate_tasks_api():
    """API endpoint לייצור משימות מהסיכום"""
    try:
        global email_analyzer
        
        # אתחול email_analyzer אם לא מאותחל
        if email_analyzer is None:
            print("🔧 מאתחל EmailAnalyzer...")
            try:
                from ai_analyzer import EmailAnalyzer
                email_analyzer = EmailAnalyzer()
                print("✅ EmailAnalyzer אותחל בהצלחה")
            except Exception as e:
                print(f"❌ שגיאה באתחול EmailAnalyzer: {e}")
                import traceback
                traceback.print_exc()
                return jsonify({
                    'success': False,
                    'error': f'שגיאה באתחול AI: {str(e)}'
                })
        
        # בדיקה נוספת שה-email_analyzer לא None
        if email_analyzer is None:
            print("❌ email_analyzer עדיין None אחרי האתחול!")
            return jsonify({
                'success': False,
                'error': 'EmailAnalyzer לא אותחל כראוי'
            })
        
        data = request.get_json()
        summary = data.get('summary', '')
        
        print(f"📧 קיבלתי סיכום לייצור משימות: {summary[:100]}...")
        
        if not summary:
            return jsonify({
                'success': False,
                'error': 'לא סופק סיכום'
            })
        
        # יצירת משימות באמצעות AI
        print(f"🤖 קורא ל-email_analyzer.generate_tasks_from_summary...")
        try:
            tasks = email_analyzer.generate_tasks_from_summary(summary)
            print(f"📋 נוצרו {len(tasks)} משימות")
        except Exception as e:
            print(f"❌ שגיאה בייצור משימות: {e}")
            # יצירת משימות בסיסיות כגיבוי
            tasks = create_fallback_tasks(summary)
            print(f"📋 נוצרו {len(tasks)} משימות גיבוי")
        
        return jsonify({
            'success': True,
            'tasks': tasks
        })
        
    except Exception as e:
        print(f"❌ שגיאה בייצור משימות: {e}")
        return jsonify({
            'success': False,
            'error': f'שגיאה כללית: {str(e)}'
        }), 500

def create_fallback_tasks(summary):
    """יצירת משימות בסיסיות כגיבוי"""
    tasks = []
    summary_lower = summary.lower()
    
    # זיהוי משימות טכניות
    if any(word in summary_lower for word in ["ג'וב", "job", "שרת", "server", "איפוס", "reset"]):
        tasks.append({
            "title": "יצירת ג'וב לאיפוס שרתים",
            "description": "צור ג'וב חדש לאיפוס השרתים כפי שנדרש",
            "priority": "חשוב",
            "category": "AI חשוב"
        })
    
    # זיהוי משימות בדיקה
    if any(word in summary_lower for word in ["בדיקה", "check", "בדוק", "היסטוריה", "history"]):
        tasks.append({
            "title": "בדיקת אפשרות למחיקת היסטוריה",
            "description": "בדוק איך ניתן למחוק את ההיסטוריה במערכת",
            "priority": "בינוני",
            "category": "AI בינוני"
        })
    
    # זיהוי משימות גיבוי
    if any(word in summary_lower for word in ["גיבוי", "backup", "גיבויים", "backups"]):
        tasks.append({
            "title": "בדיקת נושא גיבויים",
            "description": "בדוק את מצב הגיבויים של הג'ובים הקיימים",
            "priority": "חשוב",
            "category": "AI חשוב"
        })
    
    # אם לא נמצאו מילות מפתח ספציפיות, יצירת משימה כללית
    if not tasks:
        tasks.append({
            "title": "פעולה נדרשת",
            "description": f"פעולה נדרשת בהתבסס על המייל: {summary[:100]}...",
            "priority": "בינוני",
            "category": "AI בינוני"
        })
    
    return tasks

@app.route('/api/expand-reply', methods=['POST'])
def expand_reply_api():
    """API להרחבת טקסט תשובה לטקסט פורמלי באנגלית"""
    block_id = ui_block_start("🤖 הרחבת תשובה עם AI")
    
    try:
        data = request.json
        
        if not data or not data.get('brief_text'):
            ui_block_add(block_id, "❌ לא נשלח טקסט להרחבה", "ERROR")
            ui_block_end(block_id)
            return jsonify({
                'success': False,
                'error': 'לא נשלח טקסט להרחבה'
            }), 400
        
        brief_text = data.get('brief_text', '')
        sender_email = data.get('sender_email', '')
        original_subject = data.get('original_subject', '')
        
        ui_block_add(block_id, f"📝 טקסט מקורי: {brief_text[:50]}...", "INFO")
        
        # בדיקה אם יש API key
        if not GEMINI_API_KEY or GEMINI_API_KEY == "your-gemini-api-key-here":
            ui_block_add(block_id, "⚠️ אין API key של Gemini", "WARNING")
            ui_block_end(block_id)
            return jsonify({
                'success': False,
                'error': 'נדרש Gemini API key להרחבת טקסט'
            }), 400
        
        # קריאה ל-AI להרחבת הטקסט
        try:
            global email_analyzer
            if not email_analyzer:
                email_analyzer = EmailAnalyzer()
            
            expanded_text = email_analyzer.expand_reply_text(brief_text, sender_email, original_subject)
            
            ui_block_add(block_id, "✅ הטקסט הורחב בהצלחה", "SUCCESS")
            ui_block_end(block_id)
            
            return jsonify({
                'success': True,
                'expanded_text': expanded_text,
                'original_text': brief_text
            })
            
        except Exception as ai_error:
            ui_block_add(block_id, f"❌ שגיאת AI: {str(ai_error)[:100]}", "ERROR")
            ui_block_end(block_id)
            return jsonify({
                'success': False,
                'error': f'שגיאה בהרחבת טקסט: {str(ai_error)}'
            }), 500
        
    except Exception as e:
        ui_block_add(block_id, f"❌ שגיאה כללית: {str(e)[:100]}", "ERROR")
        ui_block_end(block_id)
        return jsonify({
            'success': False,
            'error': f'שגיאה כללית: {str(e)}'
        }), 500

@app.route('/api/get-summary', methods=['POST'])
def get_summary_api():
    """API לשליפת סיכום קיים מהמאגר"""
    block_id = ui_block_start("📖 שליפת סיכום קיים")
    
    try:
        data = request.json
        
        if not data:
            ui_block_add(block_id, "❌ לא נשלחו נתונים", "ERROR")
            ui_block_end(block_id)
            return jsonify({'success': False, 'error': 'לא נשלחו נתונים'}), 400
        
        item_id = data.get('item_id')
        
        if not item_id:
            ui_block_add(block_id, "❌ חסר item_id", "ERROR")
            ui_block_end(block_id)
            return jsonify({'success': False, 'error': 'חסר item_id'}), 400
        
        ui_block_add(block_id, f"📧 מחפש EntryID: {item_id[:30]}...", "INFO")
        
        # חיפוש במאגר הנתונים
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        cursor.execute('SELECT ai_summary FROM emails WHERE outlook_id = ? AND ai_summary IS NOT NULL', (item_id,))
        result = cursor.fetchone()
        
        conn.close()
        
        if result and result[0]:
            summary_text = result[0]
            ui_block_add(block_id, f"✅ נמצא סיכום ({len(summary_text)} תווים)", "SUCCESS")
            ui_block_end(block_id)
            
            # ניסיון לפרסר את הסיכום כ-JSON
            try:
                import json
                summary_json = json.loads(summary_text)
                return jsonify({
                    'success': True,
                    'has_summary': True,
                    'summary': summary_json.get('summary', ''),
                    'key_points': summary_json.get('key_points', []),
                    'action_items': summary_json.get('action_items', []),
                    'sentiment': summary_json.get('sentiment', '')
                })
            except:
                # אם זה לא JSON, מחזירים כטקסט פשוט
                return jsonify({
                    'success': True,
                    'has_summary': True,
                    'summary': summary_text,
                    'key_points': [],
                    'action_items': [],
                    'sentiment': 'לא זוהה'
                })
        else:
            ui_block_add(block_id, "ℹ️ לא נמצא סיכום קיים", "INFO")
            ui_block_end(block_id)
            return jsonify({
                'success': True,
                'has_summary': False
            })
        
    except Exception as e:
        ui_block_add(block_id, f"❌ שגיאה: {str(e)[:100]}", "ERROR")
        ui_block_end(block_id)
        return jsonify({
            'success': False,
            'error': f'שגיאה בשליפה: {str(e)}'
        }), 500

@app.route('/api/save-summary', methods=['POST'])
def save_summary_api():
    """API לשמירת סיכום במאגר הנתונים"""
    block_id = ui_block_start("💾 שמירת סיכום במאגר")
    
    try:
        data = request.json
        
        if not data:
            ui_block_add(block_id, "❌ לא נשלחו נתונים", "ERROR")
            ui_block_end(block_id)
            return jsonify({'success': False, 'error': 'לא נשלחו נתונים'}), 400
        
        item_id = data.get('item_id')
        summary = data.get('summary')
        
        if not item_id or not summary:
            ui_block_add(block_id, "❌ חסרים נתונים: item_id או summary", "ERROR")
            ui_block_end(block_id)
            return jsonify({'success': False, 'error': 'חסרים נתונים חובה'}), 400
        
        ui_block_add(block_id, f"📧 EntryID: {item_id[:30]}...", "INFO")
        ui_block_add(block_id, f"📝 אורך סיכום: {len(summary)} תווים", "INFO")
        
        # שמירה במאגר הנתונים
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # בדיקה אם המייל קיים
        cursor.execute('SELECT id FROM emails WHERE outlook_id = ?', (item_id,))
        existing = cursor.fetchone()
        
        if existing:
            # עדכון המייל הקיים
            cursor.execute('''
                UPDATE emails 
                SET ai_summary = ?,
                    last_updated = CURRENT_TIMESTAMP
                WHERE outlook_id = ?
            ''', (summary, item_id))
            ui_block_add(block_id, f"✅ עודכן מייל קיים (ID: {existing[0]})", "SUCCESS")
        else:
            # יצירת רשומה חדשה (במקרה שהמייל עדיין לא סונכרן)
            cursor.execute('''
                INSERT INTO emails (outlook_id, ai_summary, last_updated)
                VALUES (?, ?, CURRENT_TIMESTAMP)
            ''', (item_id, summary))
            ui_block_add(block_id, f"✅ נוצרה רשומה חדשה", "SUCCESS")
        
        conn.commit()
        conn.close()
        
        ui_block_add(block_id, "💾 הסיכום נשמר בהצלחה במאגר הנתונים", "SUCCESS")
        ui_block_end(block_id)
        
        return jsonify({
            'success': True,
            'message': 'הסיכום נשמר בהצלחה במאגר הנתונים'
        })
        
    except Exception as e:
        ui_block_add(block_id, f"❌ שגיאה: {str(e)[:100]}", "ERROR")
        ui_block_end(block_id)
        return jsonify({
            'success': False,
            'error': f'שגיאה בשמירה: {str(e)}'
        }), 500

@app.route('/api/analyze', methods=['POST'])
def analyze_email():
    """API לניתוח מייל בודד מ-Outlook"""
    try:
        email_data = request.json
        
        if not email_data:
            return jsonify({
                'success': False,
                'error': 'לא נשלחו נתוני מייל'
            }), 400
        
        # ניתוח המייל
        analysis = email_manager.analyze_single_email(email_data)
        
        # המרת importance_score לפורמט נכון לפי מה שצריך ל-outlook_integration
        result = {
            'category': analysis.get('category', 'work'),
            'priority': 'גבוהה' if analysis.get('importance_score', 0) > 0.7 else 'נמוכה' if analysis.get('importance_score', 0) < 0.3 else 'רגילה',
            'requires_action': len(analysis.get('action_items', [])) > 0,
            'importance': analysis.get('importance_score', 0.5),
            'summary': analysis.get('summary', ''),
            'action_items': analysis.get('action_items', [])
        }
        
        return jsonify(result)
        
    except Exception as e:
        log_to_console(f"ERROR שגיאה בניתוח מייל: {str(e)}", "ERROR")
        return jsonify({
            'success': False,
            'error': f'שגיאה בניתוח: {str(e)}'
        }), 500

@app.route('/api/analyze-meetings-ai', methods=['POST'])
def analyze_meetings_ai():
    """API לניתוח AI מרוכז של פגישות נבחרות"""
    try:
        data = request.json
        meetings = data.get('meetings', [])
        
        if not meetings:
            return jsonify({
                'success': False,
                'message': 'לא נשלחו פגישות לניתוח'
            })
        
        # יצירת בלוק לוגים לניתוח פגישות
        block_id = ui_block_start(f"🤖 ניתוח AI של {len(meetings)} פגישות")
        
        # בדיקה שה-AI זמין
        if not email_manager.ai_analyzer.is_ai_available():
            ui_block_end(block_id, "AI לא זמין - נדרש API Key", False)
            return jsonify({
                'success': False,
                'message': 'AI לא זמין - נדרש API Key'
            })
        
        updated_meetings = []
        
        # קבלת נתוני פרופיל המשתמש
        user_profile = email_manager.profile_manager.get_user_learning_stats()
        user_preferences = email_manager.profile_manager.get_important_keywords()
        user_categories = email_manager.profile_manager.get_all_category_importance()
        
        # ניתוח כל פגישה עם AI
        for i, meeting in enumerate(meetings):
            try:
                ui_block_add(block_id, f"🤖 מנתח פגישה {i+1}/{len(meetings)}: {meeting.get('subject', 'ללא נושא')[:50]}...", "INFO")
                
                # שמירת הציון המקורי לפני AI
                original_score = meeting.get('importance_score', 0.5)
                
                # ניתוח עם AI כולל נתוני פרופיל
                ai_analysis = email_manager.ai_analyzer.analyze_email_with_profile(
                    meeting, 
                    user_profile, 
                    user_preferences, 
                    user_categories
                )
                
                # עדכון הפגישה עם הניתוח החדש
                updated_meeting = meeting.copy()
                ai_score = ai_analysis.get('importance_score', original_score)
                updated_meeting['ai_importance_score'] = ai_score
                updated_meeting['importance_score'] = ai_score
                updated_meeting['ai_analysis'] = ai_analysis.get('analysis', '')
                updated_meeting['ai_processed'] = True
                updated_meeting['ai_processed_time'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                updated_meeting['score_source'] = 'AI'
                updated_meeting['original_importance_score'] = original_score
                updated_meeting['ai_summary'] = ai_analysis.get('summary', '')
                updated_meeting['ai_reason'] = ai_analysis.get('reason', '')
                
                # חישוב השינוי בציון
                score_change = ai_score - original_score
                score_change_percent = int(score_change * 100)
                
                # הודעת לוג עם השוואה
                original_percent = int(original_score * 100)
                new_percent = int(ai_score * 100)
                
                if abs(score_change) > 0.1:  # שינוי משמעותי
                    change_indicator = "📈" if score_change > 0 else "📉"
                    ui_block_add(block_id, f"{change_indicator} פגישה {i+1}: {original_percent}% → {new_percent}% ({score_change_percent:+d}%)", "SUCCESS")
                else:
                    ui_block_add(block_id, f"✅ פגישה {i+1}: {new_percent}% (ללא שינוי משמעותי)", "INFO")
                
                updated_meetings.append(updated_meeting)
                
                # שמירה בבסיס הנתונים
                try:
                    save_meeting_ai_analysis_to_db(updated_meeting)
                    ui_block_add(block_id, f"💾 פגישה {i+1} נשמרה בבסיס נתונים", "INFO")
                except Exception as e:
                    ui_block_add(block_id, f"❌ שגיאה בשמירת פגישה {i+1}: {e}", "ERROR")
                
            except Exception as e:
                ui_block_add(block_id, f"❌ שגיאה בניתוח פגישה {i+1}: {str(e)}", "ERROR")
                # הוספת הפגישה המקורית במקרה של שגיאה
                updated_meetings.append(meeting)
        
        ui_block_end(block_id, f"הניתוח הושלם: עודכנו {len(updated_meetings)} פגישות", True)
        
        return jsonify({
            'success': True,
            'message': f'ניתוח AI הושלם עבור {len(updated_meetings)} פגישות',
            'processed_count': len(updated_meetings),
            'meetings': updated_meetings
        })
        
    except Exception as e:
        log_to_console(f"ERROR שגיאה בניתוח AI של פגישות: {str(e)}", "ERROR")
        return jsonify({
            'success': False,
            'message': f'שגיאה בניתוח AI: {str(e)}'
        }), 500

@app.route('/api/analyze-meeting', methods=['POST'])
def analyze_single_meeting():
    """API לניתוח AI של פגישה בודדת (עבור ניתוח פגישה נוכחית)"""
    try:
        data = request.json
        
        # בדיקה שיש נתונים
        if not data or not data.get('subject'):
            return jsonify({
                'success': False,
                'message': 'לא נשלחו נתוני פגישה'
            }), 400
        
        # יצירת מפתח ייחודי לפגישה
        import hashlib
        subject = data.get('subject', '')
        organizer = data.get('organizer', '')
        start_time = data.get('start_time', '')
        content_key = f"{subject}|{organizer}|{start_time}"
        meeting_id = hashlib.md5(content_key.encode('utf-8')).hexdigest()
        
        # בדיקה אם הפגישה כבר נותחה ב-DB
        saved_analysis = load_meeting_ai_analysis_map().get(meeting_id)
        
        if saved_analysis and saved_analysis.get('score_source') == 'AI':
            # הפגישה כבר נותחה! שולף מה-DB
            block_id = ui_block_start(f"💾 שליפת ניתוח קיים: {subject[:50]}")
            ui_block_add(block_id, f"📊 ציון שמור: {int(saved_analysis['importance_score'] * 100)}%", "INFO")
            ui_block_add(block_id, f"📝 סיכום: {saved_analysis.get('summary', '')[:100]}...", "INFO")
            ui_block_end(block_id, "✅ הניתוח נשלף מהזיכרון (לא נשלח ל-AI שוב)", True)
            
            return jsonify({
                'success': True,
                'importance_score': saved_analysis['importance_score'],
                'ai_score': int(saved_analysis['importance_score'] * 100),
                'category': saved_analysis.get('category', 'לא זוהה'),
                'summary': saved_analysis.get('summary', ''),
                'reason': saved_analysis.get('reason', ''),
                'analysis': saved_analysis.get('summary', ''),
                'priority': 'גבוהה' if saved_analysis['importance_score'] > 0.7 else 'בינונית' if saved_analysis['importance_score'] > 0.4 else 'נמוכה',
                'from_cache': True
            })
        
        # הפגישה לא נותחה - ניתוח חדש
        # יצירת בלוק לוגים
        block_id = ui_block_start(f"📅 ניתוח AI פגישה: {data.get('subject', 'ללא נושא')[:50]}")
        
        # בדיקה שה-AI זמין
        if not email_manager.ai_analyzer.is_ai_available():
            ui_block_end(block_id, "AI לא זמין - נדרש API Key", False)
            return jsonify({
                'success': False,
                'message': 'AI לא זמין - נדרש API Key'
            }), 503
        
        ui_block_add(block_id, f"🤖 מנתח: {data.get('subject', 'ללא נושא')[:80]}", "INFO")
        
        # קבלת נתוני פרופיל המשתמש
        user_profile = email_manager.profile_manager.get_user_learning_stats()
        user_preferences = email_manager.profile_manager.get_important_keywords()
        user_categories = email_manager.profile_manager.get_all_category_importance()
        
        # ניתוח עם AI
        ai_analysis = email_manager.ai_analyzer.analyze_email_with_profile(
            data, 
            user_profile, 
            user_preferences, 
            user_categories
        )
        
        # חילוץ הציון
        ai_score = ai_analysis.get('importance_score', 0.5)
        score_percent = int(ai_score * 100) if ai_score <= 1 else int(ai_score)
        
        ui_block_add(block_id, f"📊 ציון חשיבות: {score_percent}%", "SUCCESS")
        ui_block_add(block_id, f"📝 סיכום: {ai_analysis.get('summary', 'אין סיכום')[:100]}...", "INFO")
        
        # חישוב קטגוריה לפי הציון (כמו במיילים)
        category = ""
        if ai_score >= 0.8:
            category = "AI קריטי"
        elif ai_score >= 0.6:
            category = "AI חשוב"
        elif ai_score >= 0.4:
            category = "AI בינוני"
        else:
            category = "AI נמוך"
        
        # הכנת התגובה
        response_data = {
            'success': True,
            'importance_score': ai_score,
            'ai_score': score_percent,
            'category': category,
            'summary': ai_analysis.get('summary', ''),
            'reason': ai_analysis.get('reason', ''),
            'analysis': ai_analysis.get('analysis', ''),
            'priority': 'גבוהה' if ai_score > 0.7 else 'בינונית' if ai_score > 0.4 else 'נמוכה'
        }
        
        # שמירה בבסיס הנתונים
        try:
            meeting_to_save = data.copy()
            meeting_to_save['importance_score'] = ai_score
            meeting_to_save['ai_importance_score'] = ai_score
            meeting_to_save['score_source'] = 'AI'
            meeting_to_save['summary'] = ai_analysis.get('summary', '')
            meeting_to_save['reason'] = ai_analysis.get('reason', '')
            meeting_to_save['category'] = category
            meeting_to_save['ai_processed'] = True
            meeting_to_save['ai_analysis_date'] = datetime.now().isoformat()
            
            save_meeting_ai_analysis_to_db(meeting_to_save)
            ui_block_add(block_id, "💾 הניתוח נשמר בבסיס הנתונים", "SUCCESS")
        except Exception as save_error:
            ui_block_add(block_id, f"⚠️ שגיאה בשמירה: {save_error}", "WARNING")
        
        ui_block_end(block_id, f"✅ הניתוח הושלם בהצלחה - ציון: {score_percent}%", True)
        
        return jsonify(response_data)
        
    except Exception as e:
        error_msg = f"שגיאה בניתוח פגישה: {str(e)}"
        log_to_console(f"ERROR {error_msg}", "ERROR")
        if 'block_id' in locals():
            ui_block_end(block_id, error_msg, False)
        return jsonify({
            'success': False,
            'message': error_msg
        }), 500

def analyze_meetings_smart(meetings, block_id=None):
    """ניתוח חכם של פגישות עם חישוב ציונים לפי פרופיל המשתמש"""
    # קבלת נתוני פרופיל המשתמש
    user_profile = email_manager.profile_manager.get_user_learning_stats()
    user_preferences = email_manager.profile_manager.get_important_keywords()
    user_categories = email_manager.profile_manager.get_all_category_importance()
    
    # לוג התחלה
    if block_id:
        ui_block_add(block_id, "📊 מתחיל חישוב ציוני פגישות לפי פרופיל המשתמש...", "INFO")
    else:
        log_to_console("📊 מתחיל חישוב ציוני פגישות לפי פרופיל המשתמש...", "INFO")
    
    for i, meeting in enumerate(meetings):
        try:
            # חישוב ציון חשיבות בסיסי לפי פרופיל המשתמש
            importance_score = 0.5  # ציון בסיסי
            
            # פקטורים שמשפיעים על החשיבות
            subject = meeting.get('subject', '').lower()
            attendees_count = len(meeting.get('attendees', []))
            organizer = meeting.get('organizer', '').lower()
        
            # מילות מפתח חשובות מהפרופיל
            important_keywords = user_preferences.get('keywords', ['חשוב', 'דחוף', 'קריטי', 'מנהל', 'סטטוס', 'פרויקט', 'מצגת'])
            for keyword in important_keywords:
                if keyword.lower() in subject:
                    importance_score += 0.1
            
            # כמות משתתפים - יותר משתתפים = יותר חשוב
            if attendees_count > 10:
                importance_score += 0.2
            elif attendees_count > 5:
                importance_score += 0.1
            elif attendees_count > 2:
                importance_score += 0.05
            
            # בדיקת מארגן חשוב מהפרופיל
            important_organizers = user_preferences.get('important_senders', [])
            for important_org in important_organizers:
                if important_org.lower() in organizer:
                    importance_score += 0.15
            
            # בדיקת קטגוריות מהפרופיל
            meeting_category = determine_meeting_category(meeting)
            category_weight = user_categories.get(meeting_category, 1.0)
            importance_score *= category_weight
            
            # הגבלת הציון ל-0-1
            importance_score = min(1.0, max(0.0, importance_score))
            
            # שמירת הציון המקורי לפני AI
            meeting['original_importance_score'] = importance_score
            meeting['importance_score'] = importance_score
            meeting['score_source'] = 'SMART'
            meeting['category'] = meeting_category
            
            # בדיקה אם הפגישה היום
            meeting_date = meeting.get('start_time')
            if meeting_date:
                try:
                    # המרת מחרוזת תאריך לאובייקט datetime
                    if isinstance(meeting_date, str):
                        meeting_date = datetime.strptime(meeting_date, '%Y-%m-%d %H:%M:%S')
                    
                    today = datetime.now().date()
                    meeting['is_today'] = meeting_date.date() == today
                    
                    # בדיקה אם הפגישה השבוע
                    week_start = today - timedelta(days=today.weekday())
                    week_end = week_start + timedelta(days=6)
                    meeting['is_this_week'] = week_start <= meeting_date.date() <= week_end
                except Exception as date_error:
                    log_to_console(f"⚠️ שגיאה בעיבוד תאריך פגישה: {date_error}", "WARNING")
                    meeting['is_today'] = False
                    meeting['is_this_week'] = False
    
            # לוג הציון שחושב
            score_percent = int(importance_score * 100)
            if block_id:
                ui_block_add(block_id, f"📅 פגישה {i+1}: {meeting.get('subject', 'ללא נושא')[:40]}... - ציון: {score_percent}%", "INFO")
            else:
                log_to_console(f"📅 פגישה {i+1}: {meeting.get('subject', 'ללא נושא')[:40]}... - ציון: {score_percent}%", "INFO")
            
        except Exception as e:
            if block_id:
                ui_block_add(block_id, f"❌ שגיאה בחישוב ציון פגישה {i+1}: {str(e)}", "ERROR")
            else:
                log_to_console(f"❌ שגיאה בחישוב ציון פגישה {i+1}: {str(e)}", "ERROR")
            meeting['importance_score'] = 0.5
            meeting['original_importance_score'] = 0.5
            meeting['score_source'] = 'SMART'
    
    if block_id:
        ui_block_add(block_id, f"✅ חישוב ציוני פגישות הושלם עבור {len(meetings)} פגישות", "SUCCESS")
    else:
        log_to_console(f"✅ חישוב ציוני פגישות הושלם עבור {len(meetings)} פגישות", "SUCCESS")
    return meetings

def determine_meeting_category(meeting):
    """קביעת קטגוריה לפגישה על בסיס התוכן"""
    subject = meeting.get('subject', '').lower()
    body = meeting.get('body', '').lower()
    content = f"{subject} {body}"
    
    # קטגוריות פגישות
    categories = {
        'ניהול': ['ניהול', 'מנהל', 'סטטוס', 'דוח', 'דיווח', 'עדכון'],
        'פרויקט': ['פרויקט', 'תכנון', 'פיתוח', 'בדיקה', 'איכות'],
        'מכירות': ['מכירות', 'לקוח', 'הצעת מחיר', 'חוזה', 'עסקה'],
        'הדרכה': ['הדרכה', 'הכשרה', 'למידה', 'קורס', 'סמינר'],
        'טכני': ['טכני', 'תוכנה', 'מערכת', 'באג', 'תיקון'],
        'אסטרטגי': ['אסטרטגיה', 'תכנון', 'עתיד', 'מטרות', 'יעדים']
    }
    
    for category, keywords in categories.items():
        for keyword in keywords:
            if keyword in content:
                return category
    
    return 'כללי'

@app.route('/api/console-logs')
def get_console_logs():
    """API לקבלת לוגים מהקונסול"""
    # קבלת פרמטר 'since' - מחזיר רק לוגים מאינדקס זה ואילך
    since = request.args.get('since', 0, type=int)
    
    if os.environ.get('ENABLE_DEBUG_API') == '1':
        log_to_console(f"[DEBUG API] get_console_logs called: since={since}, total_logs={len(all_console_logs)}", "DEBUG")
    
    # מחזיר רק לוגים חדשים מאינדקס 'since'
    new_logs = all_console_logs[since:]
    
    result = {
        'logs': new_logs,
        'total': len(all_console_logs),
        'since': since
    }
    
    if os.environ.get('ENABLE_DEBUG_API') == '1':
        log_to_console(f"[DEBUG API] Returning: logs_count={len(new_logs)}, total={result['total']}, since={result['since']}", "DEBUG")
    
    # מחזיר גם את האינדקס הנוכחי כדי שה-client יידע מאיפה להמשיך
    return jsonify(result)

@app.route('/api/server-id')
def get_server_id():
    """API לקבלת מזהה השרת"""
    return jsonify({'server_id': server_id})

@app.route('/api/console-reset', methods=['POST'])
def reset_console():
    """API לאיפוס הקונסול (מחיקת כל הלוגים)"""
    try:
        # ניקוי כל הלוגים
        all_console_logs.clear()
        # Don't add any log message here - the client will show its own success message
        
        return jsonify({'success': True, 'message': 'Console reset successfully'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error resetting console: {str(e)}'})

@app.route('/api/clear-console', methods=['POST'])
def clear_console():
    """API לניקוי הקונסול"""
    try:
        # ניקוי כל הלוגים
        clear_all_console_logs()
        # אין הוספת הודעה לשרת – כדי למנוע כפילויות ברענון
        
        return jsonify({'success': True, 'message': 'Console cleared successfully'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error clearing console: {str(e)}'})

@app.route('/api/test-log')
def test_log():
    """API לבדיקת לוגים"""
    log_to_console("🧪 לוג בדיקה - " + datetime.now().strftime("%H:%M:%S"), "TEST")
    return jsonify({'status': 'success', 'message': 'לוג בדיקה נוסף'})

@app.route('/api/restart-server', methods=['POST'])
def restart_server():
    """API להפעלת שרת מחדש ללא ניתוק הטרמינל"""
    try:
        log_to_console("🚀 בקשת הפעלה מחדש התקבלה", "INFO")
        log_to_console("⏳ מסמן ל-run_project.ps1 לבצע הפעלה מחדש...", "INFO")

        # יצירת קובץ דגל שיגרום ל-run_project.ps1 להפעיל שוב את השרת באותו טרמינל
        try:
            flag_path = os.path.join(os.getcwd(), 'restart.flag')
            with open(flag_path, 'w', encoding='utf-8') as f:
                f.write(datetime.now().isoformat())
        except Exception as e:
            log_to_console(f"ERROR יצירת קובץ דגל נכשלה: {e}", "ERROR")

        # כיבוי התהליך לאחר שליחת התגובה – השארת הטרמינל פועל
        import threading, time, os
        def delayed_exit():
            try:
                time.sleep(1)
            finally:
                os._exit(222)  # קוד יציאה מיוחד לסימון אתחול

        threading.Thread(target=delayed_exit, daemon=True).start()

        return jsonify({
            'status': 'success',
            'message': 'מכבה ומפעיל מחדש... הטרמינל יישאר מחובר',
            'restart_initiated': True
        })

    except Exception as e:
        log_to_console(f"ERROR שגיאה בבקשת הפעלה מחדש: {e}", "ERROR")
        return jsonify({'status': 'error', 'message': f'שגיאה בהפעלת שרת מחדש: {e}'}), 500

@app.route('/api/restart-console', methods=['POST'])
def restart_console():
    """API לאיפוס הקונסול (כשהשרת מתחיל מחדש)"""
    try:
        # ניקוי כל הלוגים
        clear_all_console_logs()
        # לא מוסיפים הודעות התחלה – ה-client מציג סטטוס בפני עצמו
        
        return jsonify({'success': True, 'message': 'Console restarted successfully'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error restarting console: {str(e)}'})

@app.route('/api/emails')
def get_emails():
    """API לקבלת מיילים מהזיכרון"""
    global cached_data
    
    # אם אין מיילים בזיכרון, נחזיר רשימה ריקה במקום לטעון מחדש
    if cached_data['emails'] is None:
        return jsonify([])
    
    emails = cached_data['emails'] or []
    # Don't log routine data retrieval - too noisy
    return jsonify(emails)

@app.route('/api/emails-step/<int:step>')
def get_emails_step(step):
    """API לקבלת מיילים בשלבים - טעינה מהירה"""
    log_to_console(f"📧 שלב {step} - מקבל בקשת מיילים...", "INFO")
    
    if step == 1:
        # שלב 1: קבלת מיילים מהירה
        emails = email_manager.get_emails()
        log_to_console(f"שלב 1 הושלם: נטענו {len(emails)} מיילים", "SUCCESS")
        return jsonify({
            'step': 1,
            'message': f'נטענו {len(emails)} מיילים',
            'emails': emails,
            'next_step': 2
        })
    elif step == 2:
        # שלב 2: ניתוח חכם מהיר
        emails = email_manager.get_emails()
        analyzed_emails = email_manager.analyze_emails_smart(emails)
        
        log_to_console(f"שלב 2 הושלם: ניתחו {len(analyzed_emails)} מיילים", "SUCCESS")
        return jsonify({
            'step': 2,
            'message': f'ניתחו {len(analyzed_emails)} מיילים',
            'emails': analyzed_emails,
            'next_step': 3
        })
    elif step == 3:
        # שלב 3: סיום
        emails = email_manager.get_emails()
        analyzed_emails = email_manager.analyze_emails_smart(emails)
        
        log_to_console(f"שלב 3 הושלם: הושלם ניתוח {len(analyzed_emails)} מיילים", "SUCCESS")
        return jsonify({
            'step': 3,
            'message': f'הושלם ניתוח {len(analyzed_emails)} מיילים',
            'emails': analyzed_emails,
            'next_step': None
        })
    
    return jsonify({'error': 'Invalid step'})

# Removed problematic chunked API

@app.route('/api/emails-progress')
def get_emails_with_progress():
    """API לקבלת מיילים עם progress bar"""
    log_to_console("📧 מקבל בקשת מיילים עם progress...", "INFO")
    
    # שלב 1: קבלת מיילים
    emails = email_manager.get_emails()
    
    # שלב 2: ניתוח חכם
    total_emails = len(emails)
    analyzed_emails = []
    
    for i, email in enumerate(emails):
        # ניתוח חכם מבוסס פרופיל משתמש
        email['importance_score'] = email_manager.calculate_smart_importance(email)
        email['category'] = email_manager.categorize_smart(email)
        email['summary'] = email_manager.generate_smart_summary(email)
        email['action_items'] = email_manager.extract_smart_action_items(email)
        
        analyzed_emails.append(email)
        
        # הדפסת התקדמות
        progress = int((i + 1) / total_emails * 100)
        log_to_console(f"📧 מנתח מיילים: {progress}% ({i + 1}/{total_emails})", "INFO")
    
    log_to_console(f"📧 מחזיר {len(analyzed_emails)} מיילים עם ניתוח חכם", "SUCCESS")
    return jsonify(analyzed_emails)

# Removed problematic stream endpoint

@app.route('/api/preferences', methods=['GET', 'POST'])
def manage_preferences():
    """API לניהול העדפות"""
    if request.method == 'POST':
        data = request.json
        email_manager.save_user_preference(
            data['type'],
            data['value'],
            data.get('weight', 1.0)
        )
        return jsonify({'status': 'success'})
    
    return jsonify(email_manager.user_preferences)

@app.route('/api/important-emails')
def get_important_emails():
    """API לקבלת מיילים חשובים (80% ומעלה)"""
    log_to_console("⭐ מקבל בקשת מיילים חשובים...", "INFO")
    emails = email_manager.get_emails()
    
    # ניתוח חכם מבוסס פרופיל משתמש
    emails = email_manager.analyze_emails_smart(emails)
    
    # סינון מיילים עם ציון חשיבות 80% ומעלה (80/100)
    important_emails = [e for e in emails if e.get('importance_score', 0) >= 0.8]
    
    # מיון לפי ציון חשיבות (גבוה לנמוך)
    important_emails = sorted(important_emails, key=lambda x: x['importance_score'], reverse=True)
    
    log_to_console(f"⭐ מחזיר {len(important_emails)} מיילים חשובים (80%+)", "SUCCESS")
    return jsonify(important_emails)

@app.route('/api/stats')
def get_stats():
    """API לקבלת סטטיסטיקות מהזיכרון"""
    global cached_data
    
    # בדיקה אם יש סטטיסטיקות בזיכרון
    if cached_data['email_stats'] is None:
        # במקום refresh_data, נחשב סטטיסטיקות מהירות מהמיילים הקיימים
        emails = cached_data['emails'] or []
        if emails:
            email_stats = calculate_email_stats(emails)
            cached_data['email_stats'] = email_stats
        else:
            # אם אין מיילים, נחזיר סטטיסטיקות ברירת מחדל
            email_stats = {
                'total_emails': 0,
                'important_emails': 0,
                'unread_emails': 0,
                'critical_emails': 0,
                'high_emails': 0,
                'medium_emails': 0,
                'low_emails': 0
            }
    
    stats = cached_data['email_stats']
    if stats is None:
        # fallback לחישוב מהיר
        emails = cached_data['emails'] or []
        total_emails = len(emails)
        critical_emails = int(total_emails * 0.10)
        important_emails = int(total_emails * 0.25)
        medium_emails = int(total_emails * 0.40)
        low_emails = int(total_emails * 0.25)
        actual_unread_emails = len([e for e in emails if not e.get('is_read', True)])
        
        stats = {
            'total_emails': total_emails,
            'important_emails': important_emails,
            'unread_emails': actual_unread_emails,
            'critical_emails': critical_emails,
            'high_emails': important_emails,
            'medium_emails': medium_emails,
            'low_emails': low_emails
        }
    
    # Don't log routine statistics retrieval - too noisy
    return jsonify(stats)

@app.route('/api/toggle-outlook')
def toggle_outlook():
    """API למעבר בין Outlook אמיתי לנתונים דמה"""
    email_manager.use_real_outlook = not email_manager.use_real_outlook
    return jsonify({
        'use_real_outlook': email_manager.use_real_outlook,
        'message': 'Outlook אמיתי' if email_manager.use_real_outlook else 'נתונים דמה'
    })

@app.route('/api/ai-status')
def ai_status():
    """API לבדיקת סטטוס AI"""
    ai_available = email_manager.ai_analyzer.is_ai_available()
    use_ai = email_manager.use_ai
    
    # הוספת לוג לקונסול
    if ai_available:
        log_to_console(f"🤖 AI זמין - {'מופעל' if use_ai else 'מושבת'}", "INFO")
    else:
        log_to_console("ERROR AI לא זמין - נדרש API Key", "ERROR")
    
    return jsonify({
        'ai_available': ai_available,
        'use_ai': use_ai,
        'message': 'AI זמין' if ai_available else 'AI לא זמין - נדרש API Key'
    })

@app.route('/api/toggle-ai')
def toggle_ai():
    """API למעבר בין AI לניתוח בסיסי"""
    email_manager.use_ai = not email_manager.use_ai
    
    # הוספת לוג לקונסול
    status = 'מופעל' if email_manager.use_ai else 'מושבת'
    log_to_console(f"🔄 AI {status}", "INFO")
    
    return jsonify({
        'use_ai': email_manager.use_ai,
        'message': 'AI מופעל' if email_manager.use_ai else 'AI מושבת'
    })

@app.route('/api/test-outlook')
def test_outlook():
    """API לבדיקת חיבור ל-Outlook"""
    try:
        log_to_console("🔍 בודק חיבור ל-Outlook...", "INFO")
        
        if email_manager.connect_to_outlook():
            # בדיקה נוספת של מספר המיילים
            try:
                messages = email_manager.inbox.Items
                email_count = messages.Count
                log_to_console(f"📧 נמצאו {email_count} מיילים ב-Inbox", "INFO")
                
                return jsonify({
                    'success': True,
                    'message': f'חיבור ל-Outlook הצליח! נמצאו {email_count} מיילים ב-Inbox',
                    'email_count': email_count,
                    'outlook_connected': True
                })
            except Exception as e:
                log_to_console(f"⚠️ לא ניתן לספור מיילים: {e}", "WARNING")
                return jsonify({
                    'success': True,
                    'message': 'חיבור ל-Outlook הצליח אבל לא ניתן לספור מיילים',
                    'email_count': 0,
                    'outlook_connected': True,
                    'warning': str(e)
                })
        else:
            log_to_console("ERROR חיבור ל-Outlook נכשל", "ERROR")
            return jsonify({
                'success': False,
                'message': 'לא ניתן להתחבר ל-Outlook',
                'email_count': 0,
                'outlook_connected': False
            })
    except Exception as e:
        log_to_console(f"ERROR שגיאה בבדיקת Outlook: {e}", "ERROR")
        return jsonify({
            'success': False,
            'message': f'שגיאה: {str(e)}',
            'email_count': 0,
            'outlook_connected': False
        })

@app.route('/api/user-preferences', methods=['GET', 'POST'])
def manage_user_preferences():
    """API לניהול העדפות משתמש מתקדמות"""
    if request.method == 'POST':
        try:
            data = request.json
            
            # שמירת העדפות במסד נתונים
            conn = sqlite3.connect(email_manager.db_path)
            cursor = conn.cursor()
            
            # מחיקת העדפות קיימות
            cursor.execute('DELETE FROM user_preferences_advanced WHERE preference_type IN (?, ?, ?)', 
                         ('important_categories', 'important_senders', 'important_keywords'))
            
            # שמירת קטגוריות חשובות
            for category in data.get('important_categories', []):
                cursor.execute('''
                    INSERT INTO user_preferences_advanced (preference_type, preference_key, preference_value, confidence_score)
                    VALUES (?, ?, ?, ?)
                ''', ('important_categories', category, category, 1.0))
            
            # שמירת שולחים חשובים
            for sender in data.get('important_senders', []):
                cursor.execute('''
                    INSERT INTO user_preferences_advanced (preference_type, preference_key, preference_value, confidence_score)
                    VALUES (?, ?, ?, ?)
                ''', ('important_senders', sender, sender, 1.0))
            
            # שמירת מילות מפתח חשובות
            for keyword in data.get('important_keywords', []):
                cursor.execute('''
                    INSERT INTO user_preferences_advanced (preference_type, preference_key, preference_value, confidence_score)
                    VALUES (?, ?, ?, ?)
                ''', ('important_keywords', keyword, keyword, 1.0))
            
            conn.commit()
            conn.close()
            
            return jsonify({'success': True, 'message': 'Preferences saved successfully'})
            
        except Exception as e:
            return jsonify({'success': False, 'message': f'Error saving preferences: {str(e)}'})
    
    else:  # GET
        try:
            conn = sqlite3.connect(email_manager.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT preference_type, preference_key, preference_value 
                FROM user_preferences_advanced 
                WHERE preference_type IN (?, ?, ?)
            ''', ('important_categories', 'important_senders', 'important_keywords'))
            
            rows = cursor.fetchall()
            conn.close()
            
            preferences = {
                'important_categories': [],
                'important_senders': [],
                'important_keywords': []
            }
            
            for pref_type, pref_key, pref_value in rows:
                if pref_type == 'important_categories':
                    preferences['important_categories'].append(pref_value)
                elif pref_type == 'important_senders':
                    preferences['important_senders'].append(pref_value)
                elif pref_type == 'important_keywords':
                    preferences['important_keywords'].append(pref_value)
            
            return jsonify(preferences)
            
        except Exception as e:
            return jsonify({'error': f'Error loading preferences: {str(e)}'})

# Removed duplicate record_user_feedback function - see line 766 for the actual implementation

@app.route('/api/learning-stats')
def get_learning_stats():
    """API לקבלת סטטיסטיקות למידה"""
    try:
        stats = email_manager.profile_manager.get_user_learning_stats()
        return jsonify(stats)
    except Exception as e:
        return jsonify({
            'error': f'Error getting statistics: {str(e)}'
        })

@app.route('/api/toggle-learning')
def toggle_learning():
    """API להפעלה/כיבוי מערכת למידה"""
    email_manager.use_learning = not email_manager.use_learning
    return jsonify({
        'use_learning': email_manager.use_learning,
        'message': 'Learning system enabled' if email_manager.use_learning else 'Learning system disabled'
    })

@app.route('/api/user-profile')
def get_user_profile():
    """API לקבלת פרופיל משתמש"""
    try:
        profile_data = {
            'patterns': email_manager.profile_manager.user_patterns,
            'preferences': email_manager.profile_manager.profile_data,
            'stats': email_manager.profile_manager.get_user_learning_stats()
        }
        return jsonify(profile_data)
    except Exception as e:
        return jsonify({
            'error': f'Error getting profile: {str(e)}'
        })

@app.route('/api/reset-learning', methods=['POST'])
def reset_learning():
    """API לאיפוס מערכת למידה"""
    try:
        # איפוס דפוסי למידה
        conn = sqlite3.connect(email_manager.db_path)
        cursor = conn.cursor()
        
        cursor.execute('DELETE FROM user_patterns')
        cursor.execute('DELETE FROM user_feedback')
        cursor.execute('DELETE FROM user_preferences_advanced')
        
        conn.commit()
        conn.close()
        
        # איפוס זיכרון
        email_manager.profile_manager.user_patterns = {}
        email_manager.profile_manager.profile_data = {}
        
        return jsonify({
            'success': True,
            'message': 'Learning system reset successfully'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error resetting learning system: {str(e)}'
        })

@app.route('/learning-management')
def learning_management():
    """דף ניהול למידה חכמה"""
    return render_template('learning_management.html')

@app.route('/api/clear-all-data', methods=['POST'])
def clear_all_data():
    """API למחיקת כל הנתונים"""
    try:
        conn = sqlite3.connect(email_manager.db_path)
        cursor = conn.cursor()
        
        # מחיקת כל הטבלאות
        cursor.execute('DELETE FROM user_patterns')
        cursor.execute('DELETE FROM user_feedback')
        cursor.execute('DELETE FROM user_preferences_advanced')
        cursor.execute('DELETE FROM user_preferences')
        cursor.execute('DELETE FROM important_emails')
        cursor.execute('DELETE FROM ai_analysis')
        
        conn.commit()
        conn.close()
        
        # איפוס זיכרון
        email_manager.profile_manager.user_patterns = {}
        email_manager.profile_manager.profile_data = {}
        email_manager.user_preferences = {}
        
        return jsonify({
            'success': True,
            'message': 'All data cleared successfully'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error clearing data: {str(e)}'
        })

@app.route('/api/user-feedback', methods=['POST'])
def record_user_feedback():
    """API לרישום משוב משתמש"""
    try:
        data = request.json
        email_data = data.get('email_data', {})
        feedback_type = data.get('feedback_type')  # 'importance' או 'category'
        user_value = data.get('user_value')
        ai_value = data.get('ai_value')
        
        email_manager.profile_manager.record_user_feedback(
            email_data, feedback_type, user_value, ai_value
        )
        
        return jsonify({
            'success': True,
            'message': 'Feedback recorded successfully'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error recording feedback: {str(e)}'
        })


@app.route('/api/load-all-emails')
def load_all_emails():
    """API לטעינת כל המיילים"""
    try:
        log_to_console("📧 מתחיל טעינת כל המיילים...", "INFO")
        
        # טעינת כל המיילים ללא הגבלה
        emails = email_manager.get_emails_from_outlook(1000)  # מקסימום 1000 מיילים
        
        if emails:
            log_to_console(f"📧 נטענו {len(emails)} מיילים", "SUCCESS")
            return jsonify({
                'success': True,
                'message': f'נטענו {len(emails)} מיילים',
                'email_count': len(emails),
                'emails': emails
            })
        else:
            log_to_console("ERROR לא נטענו מיילים", "ERROR")
            return jsonify({
                'success': False,
                'message': 'לא נטענו מיילים',
                'email_count': 0
            })
            
    except Exception as e:
        log_to_console(f"ERROR שגיאה בטעינת מיילים: {e}", "ERROR")
        return jsonify({
            'success': False,
            'message': f'שגיאה: {str(e)}',
            'email_count': 0
        })

@app.route('/api/analyze-emails-ai', methods=['POST'])
def analyze_emails_ai():
    """API לניתוח AI מרוכז של מיילים נבחרים"""
    try:
        data = request.json
        emails = data if isinstance(data, list) else data.get('emails', [])
        
        if not emails:
            return jsonify({
                'success': False,
                'message': 'לא נשלחו מיילים לניתוח'
            })
        
        # בלוק לוג לקונסול עבור הניתוח
        block_id = ui_block_start("🧠 ניתוח AI נבחרים")
        ui_block_add(block_id, f"🚀 מתחיל ניתוח של {len(emails)} מיילים...", "INFO")
        
        # בדיקה שה-AI זמין
        if not email_manager.ai_analyzer.is_ai_available():
            return jsonify({
                'success': False,
                'message': 'AI לא זמין - נדרש API Key'
            })
        
        updated_emails = []
        
        # קבלת נתוני פרופיל המשתמש
        user_profile = email_manager.profile_manager.get_user_learning_stats()
        user_preferences = email_manager.profile_manager.get_important_keywords()
        user_categories = email_manager.profile_manager.get_all_category_importance()
        
        # ניתוח כל מייל עם AI
        for i, email in enumerate(emails):
            try:
                ui_block_add(block_id, f"🔍 מנתח מייל {i+1}/{len(emails)}: {email.get('subject', 'ללא נושא')[:50] if isinstance(email, dict) else str(email)[:50]}", "INFO")
                
                # ניתוח עם AI כולל נתוני פרופיל
                ai_analysis = email_manager.ai_analyzer.analyze_email_with_profile(
                    email, 
                    user_profile, 
                    user_preferences, 
                    user_categories
                )
                
                # עדכון המייל עם הניתוח החדש
                updated_email = email.copy() if isinstance(email, dict) else email
                
                # שמירת הציון המקורי (גם אם כבר קיים – נשמור את הישן בפעם הראשונה בלבד)
                if isinstance(email, dict):
                    if 'original_importance_score' not in email:
                        updated_email['original_importance_score'] = email.get('importance_score', 0.5)
                    else:
                        updated_email['original_importance_score'] = email.get('original_importance_score', 0.5)
                    updated_email['ai_importance_score'] = ai_analysis.get('importance_score', email.get('importance_score', 0.5))
                else:
                    updated_email['original_importance_score'] = 0.5
                    updated_email['ai_importance_score'] = ai_analysis.get('importance_score', 0.5)
                
                # עדכון הציון החדש
                if isinstance(email, dict):
                    updated_email['importance_score'] = ai_analysis.get('importance_score', email.get('importance_score', 0.5))
                    updated_email['category'] = ai_analysis.get('category', email.get('category', 'work'))
                    updated_email['summary'] = ai_analysis.get('summary', email.get('summary', ''))
                    updated_email['action_items'] = ai_analysis.get('action_items', email.get('action_items', []))
                else:
                    updated_email['importance_score'] = ai_analysis.get('importance_score', 0.5)
                    updated_email['category'] = ai_analysis.get('category', 'work')
                    updated_email['summary'] = ai_analysis.get('summary', '')
                    updated_email['action_items'] = ai_analysis.get('action_items', [])
                updated_email['ai_analyzed'] = True
                updated_email['ai_analysis_date'] = datetime.now().isoformat()
                # שמירת מקור וסיבת שינוי גם באובייקט המייל (לשימוש ב-UI)
                try:
                    updated_email['score_source'] = 'AI'  # תמיד AI בניתוח AI נבחרים
                    if ai_analysis.get('reason'):
                        updated_email['reason'] = ai_analysis.get('reason')
                    if ai_analysis.get('summary'):
                        updated_email['ai_summary'] = ai_analysis.get('summary')
                except Exception:
                    pass
                
                # דיווח מקור הציון וציון באחוזים
                source = 'AI'
                try:
                    source = ai_analysis.get('score_source', 'AI')
                except Exception:
                    pass
                score_percent = int((updated_email.get('importance_score', 0.0)) * 100)
                # הוספת סיבה קצרה לשינוי (אם קיימת סיכום/מילות מפתח)
                reason = ''
                try:
                    if ai_analysis.get('summary'):
                        reason = f" – סיכום: {ai_analysis.get('summary')[:60]}"
                except Exception:
                    pass
                ui_block_add(block_id, f"✅ עודכן מייל {i+1}: {score_percent}% (מקור: {source}){reason}", "SUCCESS")
                updated_emails.append(updated_email)
                # שמירה מתמשכת בבסיס הנתונים
                try:
                    save_ai_analysis_to_db(updated_email)
                    # שמירה בבסיס נתונים הצליחה
                except Exception as e:
                    ui_block_add(block_id, f"❌ שגיאה בשמירת מייל {i+1}: {e}", "ERROR")
                
            except Exception as e:
                ui_block_add(block_id, f"❌ שגיאה בניתוח מייל {i+1}: {e}", "ERROR")
                # שמירת המייל המקורי במקרה של שגיאה
                updated_emails.append(email)
                continue
        
        ui_block_end(block_id, f"הניתוח הושלם: עודכנו {len(updated_emails)} מיילים", True)
        
        # עדכון המיילים בזיכרון
        global cached_data
        if cached_data['emails']:
            # עדכון המיילים המעודכנים בזיכרון
            for updated_email in updated_emails:
                for i, original_email in enumerate(cached_data['emails']):
                    # התאמה על בסיס תוכן המייל (נושא + שולח + תאריך)
                    if (original_email.get('subject') == updated_email.get('subject') and 
                        original_email.get('sender') == updated_email.get('sender') and
                        original_email.get('received_time') == updated_email.get('received_time')):
                        # מיזוג עדין כדי לא לאבד original_importance_score שכבר נשמר
                        merged = {**original_email, **updated_email}
                        if 'original_importance_score' in original_email and 'original_importance_score' not in updated_email:
                            merged['original_importance_score'] = original_email['original_importance_score']
                        cached_data['emails'][i] = merged
                        ui_block_add(block_id, f"🔄 מייל {i+1} עודכן בזיכרון", "INFO")
                        break
        
        # עדכון סטטיסטיקות
        email_stats = calculate_email_stats(cached_data['emails'] or [])
        cached_data['email_stats'] = email_stats
        
        # הודעת סיכום בלוג הכללי – מבוטלת כדי למנוע כפילות מחוץ לבלוק
        # log_to_console(f"Updated {len(updated_emails)} emails in memory", "SUCCESS")
        
        return jsonify({
            'success': True,
            'message': f'ניתוח AI הושלם עבור {len(updated_emails)} מיילים',
            'updated_count': len(updated_emails),
            'updated_emails': updated_emails
        })
        
    except Exception as e:
        try:
            ui_block_end(block_id, f"❌ שגיאה בניתוח AI: {e}", False)
        except Exception:
            pass
        log_to_console(f"ERROR שגיאה בניתוח AI: {e}", "ERROR")
        return jsonify({
            'success': False,
            'message': f'שגיאה בניתוח AI: {str(e)}'
        })

def clear_all_console_logs():
    """ניקוי כל הלוגים מהקונסול"""
    global all_console_logs
    all_console_logs.clear()

@app.route('/api/outlook-addin/analyze-email', methods=['POST'])
def analyze_email_for_addin():
    """API לניתוח מייל מה-Outlook Add-in"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({
                'success': False,
                'error': 'לא נשלחו נתונים'
            }), 400
        
        # יצירת בלוק לוגים לניתוח Add-in
        block_id = ui_block_start("🔌 ניתוח מייל מ-Outlook Add-in")
        ui_block_add(block_id, f"📧 נושא: {data.get('subject', 'ללא נושא')[:50]}...", "INFO")
        ui_block_add(block_id, f"👤 שולח: {data.get('sender_name', 'לא ידוע')}", "INFO")
        
        # בדיקה שה-AI זמין
        if not email_manager.ai_analyzer.is_ai_available():
            ui_block_end(block_id, "AI לא זמין - נדרש API Key", False)
            return jsonify({
                'success': False,
                'error': 'AI לא זמין - נדרש API Key'
            }), 503
        
        # ניתוח עם המערכת שלנו
        ui_block_add(block_id, "🧠 מתחיל ניתוח AI...", "INFO")
        
        # יצירת אובייקט מייל זמני לניתוח
        # תמיד ניתוח חדש - גם אם המייל כבר נותח בעבר!
        # זה מאפשר עדכון קטגוריה ו-PRIORITYNUM גם למיילים שנותחו
        email_for_analysis = {
            'subject': data.get('subject', ''),
            'body': data.get('body', ''),
            'sender': data.get('sender', ''),
            'sender_name': data.get('sender_name', ''),
            'date': data.get('date', ''),
            'ai_analyzed': False,  # תמיד False = תמיד ניתוח מחדש
            'force_reanalyze': True  # דגל מפורש לניתוח מחדש
        }
        
        # ניתוח AI מלא
        ai_score = email_manager.ai_analyzer.analyze_email_importance(email_for_analysis)
        
        # יצירת אובייקט ai_analysis עם המבנה הנכון
        ai_analysis = {
            'importance_score': ai_score,
            'category': email_manager.ai_analyzer.categorize_email(email_for_analysis),
            'summary': email_manager.ai_analyzer.summarize_email(email_for_analysis),
            'action_items': email_manager.ai_analyzer.extract_action_items(email_for_analysis)
        }
        
        # חישוב ציון חכם מבוסס פרופיל
        smart_score = email_manager.calculate_smart_importance(email_for_analysis)
        smart_category = email_manager.categorize_smart(email_for_analysis)
        smart_summary = email_manager.generate_smart_summary(email_for_analysis)
        smart_actions = email_manager.extract_smart_action_items(email_for_analysis)
        
        # שילוב תוצאות AI עם הניתוח החכם
        final_score = (ai_analysis['importance_score'] + smart_score) / 2
        final_category = smart_category if smart_category else ai_analysis.get('category', 'לא סווג')
        final_summary = smart_summary if smart_summary else ai_analysis.get('summary', 'אין סיכום זמין')
        final_actions = smart_actions if smart_actions else ai_analysis.get('action_items', [])
        
        ui_block_add(block_id, f"📊 ציון AI: {int(ai_analysis['importance_score'] * 100)}%", "INFO")
        ui_block_add(block_id, f"🧠 ציון חכם: {int(smart_score * 100)}%", "INFO")
        ui_block_add(block_id, f"📈 ציון סופי: {int(final_score * 100)}%", "SUCCESS")
        ui_block_add(block_id, f"🏷️ קטגוריה: {final_category}", "INFO")
        
        # עדכון PRIORITYNUM ב-Outlook אם יש itemId
        outlook_update_success = False
        outlook_error_msg = None
        item_id = data.get('itemId')
        
        ui_block_add(block_id, f"🔍 מחפש מייל לעדכון (itemId: {bool(item_id)})", "INFO")
        
        # אם אין itemId, ננסה לחפש לפי subject+sender
        mail_item = None
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            
            if item_id and len(item_id) > 10:
                try:
                    ui_block_add(block_id, f"🔄 מנסה לטעון מייל לפי ItemId (length={len(item_id)})...", "INFO")
                    mail_item = outlook.GetItemFromID(item_id)
                    ui_block_add(block_id, "✅ מייל נמצא לפי ItemId", "SUCCESS")
                except Exception as id_error:
                    error_msg = str(id_error)
                    ui_block_add(block_id, f"⚠️ ItemId לא עבד: {error_msg[:100]}", "WARNING")
                    mail_item = None
            else:
                ui_block_add(block_id, f"⚠️ ItemId קצר מדי או לא קיים (length={len(item_id) if item_id else 0})", "WARNING")
            
            # אם לא הצלחנו לקבל את המייל לפי itemId, ננסה לחפש
            if not mail_item:
                ui_block_add(block_id, "🔍 מחפש מייל לפי נושא ושולח...", "INFO")
                subject = data.get('subject', '')[:100]
                sender = data.get('sender', '')
                
                # חיפוש בתיבת הדואר הנכנס
                inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
                items = inbox.Items
                items.Sort("[ReceivedTime]", True)  # ממוין לפי זמן, מהחדש ביותר
                
                # חיפוש ב-100 המיילים האחרונים
                count = 0
                matches_found = []
                for item in items:
                    count += 1
                    if count > 100:
                        break
                    try:
                        if hasattr(item, 'Subject'):
                            item_subject = item.Subject or ''
                            
                            # ניסיון לקבל את כתובת השולח מכמה מקורות
                            item_sender = ''
                            try:
                                if hasattr(item, 'SenderEmailAddress') and item.SenderEmailAddress:
                                    item_sender = item.SenderEmailAddress
                                elif hasattr(item, 'SenderName') and item.SenderName:
                                    item_sender = item.SenderName
                                elif hasattr(item, 'Sender') and item.Sender:
                                    if hasattr(item.Sender, 'Address'):
                                        item_sender = item.Sender.Address
                                    elif hasattr(item.Sender, 'Name'):
                                        item_sender = item.Sender.Name
                            except:
                                item_sender = ''
                            
                            # בדיקה אם יש התאמה - רק לפי נושא (השולח לא אמין)
                            if subject and subject in item_subject:
                                matches_found.append(f"{item_subject[:30]}")
                                if not mail_item:  # לוקחים את הראשון
                                    mail_item = item
                                    ui_block_add(block_id, f"✅ מייל נמצא: {item.Subject[:30]}...", "SUCCESS")
                                    break
                    except Exception as search_error:
                        continue
                
                if not mail_item:
                    ui_block_add(block_id, f"⚠️ לא נמצא מייל מתאים (חיפשנו ב-{count} מיילים)", "WARNING")
            
            # אם מצאנו את המייל - נעדכן אותו
            if mail_item:
                ui_block_add(block_id, "🔄 מעדכן PRIORITYNUM...", "INFO")
                score_percent = int(final_score * 100)
                
                # עדכון PRIORITYNUM
                priority_prop = mail_item.UserProperties.Find("PRIORITYNUM")
                if not priority_prop:
                    priority_prop = mail_item.UserProperties.Add("PRIORITYNUM", 3)  # 3 = olNumber
                priority_prop.Value = score_percent
                
                # עדכון AISCORE
                aiscore_prop = mail_item.UserProperties.Find("AISCORE")
                if not aiscore_prop:
                    aiscore_prop = mail_item.UserProperties.Add("AISCORE", 1)  # 1 = olText
                aiscore_prop.Value = f"{score_percent}%"
                
                mail_item.Save()
                outlook_update_success = True
                ui_block_add(block_id, f"✅ PRIORITYNUM עודכן ל-{score_percent}", "SUCCESS")
            else:
                outlook_error_msg = "לא נמצא מייל תואם ב-Outlook"
                ui_block_add(block_id, f"⚠️ {outlook_error_msg}", "WARNING")
            
            pythoncom.CoUninitialize()
            
        except Exception as outlook_error:
            outlook_error_msg = str(outlook_error)
            ui_block_add(block_id, f"⚠️ שגיאה בעדכון Outlook: {outlook_error_msg}", "WARNING")
            try:
                pythoncom.CoUninitialize()
            except:
                pass
        
        ui_block_end(block_id, "ניתוח Add-in הושלם בהצלחה", True)
        
        return jsonify({
            'success': True,
            'importance_score': final_score,
            'category': final_category,
            'summary': final_summary,
            'action_items': final_actions,
            'ai_score': ai_analysis['importance_score'],
            'smart_score': smart_score,
            'analysis_time': datetime.now().isoformat(),
            'outlook_updated': outlook_update_success,
            'outlook_error': outlook_error_msg
        })
        
    except Exception as e:
        error_msg = f'שגיאה בניתוח מייל Add-in: {str(e)}'
        try:
            ui_block_end(block_id, error_msg, False)
        except Exception:
            pass
        return jsonify({
            'success': False,
            'error': error_msg
        }), 500

@app.route('/api/outlook-addin/get-profile', methods=['GET'])
def get_profile_for_addin():
    """API לקבלת פרופיל משתמש עבור Add-in"""
    try:
        # קבלת נתוני פרופיל
        profile_stats = email_manager.profile_manager.get_user_learning_stats()
        important_keywords = email_manager.profile_manager.get_important_keywords()
        important_senders = email_manager.profile_manager.get_important_senders() if hasattr(email_manager.profile_manager, 'get_important_senders') else []
        category_importance = email_manager.profile_manager.get_all_category_importance()
        
        return jsonify({
            'success': True,
            'profile': {
                'total_feedback': profile_stats.get('total_feedback', 0),
                'learning_progress': profile_stats.get('learning_progress', 0),
                'accuracy_rate': profile_stats.get('accuracy_rate', 0),
                'important_keywords': important_keywords,
                'important_senders': important_senders,
                'category_importance': category_importance
            }
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'שגיאה בקבלת פרופיל: {str(e)}'
        }), 500

@app.route('/api/outlook-addin/update-profile', methods=['POST'])
def update_profile_from_addin():
    """API לעדכון פרופיל משתמש מ-Add-in"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({
                'success': False,
                'error': 'לא נשלחו נתונים'
            }), 400
        
        # עדכון מילות מפתח חשובות
        if 'important_keywords' in data:
            email_manager.profile_manager.update_important_keywords(data['important_keywords'])
        
        # עדכון שולחים חשובים
        if 'important_senders' in data:
            email_manager.profile_manager.update_important_senders(data['important_senders'])
        
        # עדכון חשיבות קטגוריות
        if 'category_importance' in data:
            email_manager.profile_manager.update_category_importance(data['category_importance'])
        
        return jsonify({
            'success': True,
            'message': 'פרופיל עודכן בהצלחה'
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'שגיאה בעדכון פרופיל: {str(e)}'
        }), 500

@app.route('/outlook_addin/<path:filename>')
def serve_addin_files(filename):
    """שירות קבצי ה-Add-in"""
    try:
        addin_path = os.path.join('outlook_addin', filename)
        if os.path.exists(addin_path):
            return send_file(addin_path)
        else:
            return jsonify({'error': 'קובץ לא נמצא'}), 404
    except Exception as e:
        return jsonify({'error': f'שגיאה בטעינת קובץ: {str(e)}'}), 500

@app.route('/api/create-backup', methods=['POST'])
def create_backup():
    """API ליצירת גיבוי מלא - פרומפטים, תיעוד וגיבוי ZIP"""
    try:
        block_id = ui_block_start("📦 יצירת גיבוי מלא")
        ui_block_add(block_id, "🚀 מתחיל תהליך גיבוי מלא...", "INFO")
        
        # שלב 1: יצירת פרומפטים
        ui_block_add(block_id, "📝 שלב 1: יוצר פרומפטים ל-Cursor...", "INFO")
        ui_block_add(block_id, "🚀 מתחיל יצירת קבצי פרומפטים ל-Cursor...", "INFO")
        try:
            
            # יצירת תיקיית פרומפטים בפרויקט
            project_path = os.getcwd()
            prompts_folder = os.path.join(project_path, "Cursor_Prompts")
            os.makedirs(prompts_folder, exist_ok=True)
            
            # יצירת קבצי פרומפטים
            prompts_data = {
                "01_Main_Project_Prompt.txt": """# Outlook Email Manager - Main Project Prompt

## Project Overview
This is a comprehensive email management system that integrates with Microsoft Outlook and uses AI for intelligent email analysis and prioritization.

## Key Features
- Outlook COM integration for email/meeting management
- AI-powered importance scoring using Google Gemini API
- Smart learning system that adapts to user preferences
- User profile management with behavioral learning
- Real-time console logging with collapsible blocks
- Dark/light mode support
- Priority-based categorization (Critical, High, Medium, Low)

## Technical Stack
- Backend: Flask (Python)
- Frontend: HTML/CSS/JavaScript
- Database: SQLite
- AI: Google Gemini API
- Integration: Microsoft Outlook COM

## Main Files
- app_with_ai.py: Main Flask application
- user_profile_manager.py: Learning and profile management
- ai_analyzer.py: AI analysis engine
- templates/: Frontend templates
- email_manager.db: Main database

## Development Guidelines
- Follow Hebrew UI conventions
- Maintain responsive design
- Ensure dark mode compatibility
- Use collapsible console logging
- Implement proper error handling""",
                
                "02_Flask_Application.txt": """# Flask Application Development

## Core Application Structure
The main Flask app is in app_with_ai.py with the following key components:

### Routes
- /: Email management page
- /meetings: Meeting management page  
- /consol: Console/logging page
- /learning-management: Smart learning management

### API Endpoints
- /api/emails: Get emails with AI analysis
- /api/meetings: Get meetings with AI analysis
- /api/user-feedback: Record user feedback
- /api/analyze-emails-ai: AI analysis for emails
- /api/analyze-meetings-ai: AI analysis for meetings
- /api/create-backup: Full backup with prompts/docs

### Key Functions
- analyze_emails_smart(): Smart email analysis
- analyze_meetings_smart(): Smart meeting analysis
- refresh_data(): Data refresh with caching
- init_ai_analysis_table(): Database initialization

## Development Notes
- Use ui_block_start/end for console logging
- Implement proper error handling
- Maintain Hebrew language support
- Follow RESTful API conventions""",
                
                "03_Frontend_Development.txt": """# Frontend Development Guidelines

## Template Structure
- index.html: Email management interface
- meetings.html: Meeting management interface
- consol.html: Console/logging interface
- learning_management.html: Learning management interface

## Key Features
- Responsive design with CSS Grid/Flexbox
- Dark/light mode toggle
- Interactive priority buttons
- Real-time data updates
- Modal dialogs for detailed information
- Progress bars and visual indicators

## CSS Guidelines
- Use CSS custom properties for theming
- Implement smooth transitions
- Ensure accessibility
- Support RTL (Hebrew) text direction
- Maintain consistent spacing and typography

## JavaScript Features
- Async/await for API calls
- Real-time console updates
- Interactive modals and tooltips
- Form validation and feedback
- Local storage for preferences""",
                
                "04_Outlook_Integration.txt": """# Microsoft Outlook Integration

## COM Integration
The system uses Python's win32com.client to interact with Outlook:

### Key Classes
- EmailManager: Main email handling
- Outlook connection management
- Email/meeting data extraction

### Data Extraction
- Email properties (subject, sender, body, date)
- Meeting details (organizer, attendees, time)
- Attachment handling
- Importance flags

## Integration Points
- Real-time email monitoring
- Meeting calendar integration
- Contact information extraction
- Folder organization

## Development Notes
- Handle Outlook COM errors gracefully
- Implement proper connection management
- Support different Outlook versions
- Maintain performance with large mailboxes""",
                
                "05_AI_Integration.txt": """# AI Integration with Google Gemini

## AI Analysis Engine
Located in ai_analyzer.py, provides intelligent analysis:

### Features
- Email importance scoring
- Meeting priority assessment
- Keyword extraction
- Sentiment analysis
- Action item identification

### Integration Points
- Google Gemini API calls
- User profile integration
- Learning from feedback
- Pattern recognition

## API Usage
- Structured prompts for consistent results
- Error handling and fallbacks
- Rate limiting considerations
- Response parsing and validation

## Development Guidelines
- Use meaningful prompts
- Implement proper error handling
- Cache results when appropriate
- Monitor API usage and costs""",
                
                "06_Deployment.txt": """# Deployment and Maintenance

## Production Deployment
- Flask app with proper WSGI server
- Database maintenance and backups
- Log file management
- Error monitoring

## Backup Strategy
- Automated daily backups
- Version control with Git
- Documentation updates
- Prompt file maintenance

## Maintenance Tasks
- Database optimization
- Log file cleanup
- Performance monitoring
- Security updates

## Development Environment
- Python virtual environment
- Required packages in requirements.txt
- Development vs production configs
- Testing procedures"""
            }
            
            for filename, content in prompts_data.items():
                file_path = os.path.join(prompts_folder, filename)
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                ui_block_add(block_id, f"   ✅ נוצר: {filename}", "INFO")
            
            ui_block_add(block_id, "✅ פרומפטים נוצרו בהצלחה", "SUCCESS")
            
        except Exception as prompts_error:
            ui_block_add(block_id, f"⚠️ שגיאה ביצירת פרומפטים: {str(prompts_error)}", "WARNING")
        
        # שלב 2: יצירת תיעוד
        ui_block_add(block_id, "📚 שלב 2: יוצר תיעוד מעודכן...", "INFO")
        ui_block_add(block_id, "🚀 מתחיל יצירת/רענון קבצי תיעוד...", "INFO")
        try:
            
            # יצירת תיקיית תיעוד בפרויקט
            docs_folder = os.path.join(project_path, "docs")
            os.makedirs(docs_folder, exist_ok=True)
            
            # יצירת קבצי תיעוד
            docs_data = {
                "README.md": """# Outlook Email Manager

מערכת ניהול מיילים חכמה עם AI

## תכונות עיקריות
- ניהול מיילים ופגישות מ-Outlook
- ניתוח AI לחשיבות וחישוב ציונים
- מערכת למידה חכמה שמתאימה להעדפות המשתמש
- ניהול פרופיל משתמש עם למידה התנהגותית
- לוגים בזמן אמת עם בלוקים מתקפלים
- תמיכה בערכה כהה ובהירה
- קטגוריזציה לפי עדיפות (קריטי, חשוב, בינוני, נמוך)

## התקנה
1. התקן את הדרישות: `pip install -r requirements.txt`
2. הפעל את השרת: `python app_with_ai.py`
3. פתח בדפדפן: `http://localhost:5000`

## שימוש
- ניהול מיילים: דף ראשי
- ניהול פגישות: דף פגישות
- קונסול: מעקב לוגים
- ניהול למידה: הגדרות וסטטיסטיקות""",
                
                "API_DOCUMENTATION.md": """# API Documentation

## Email Management
- `GET /api/emails`: קבלת מיילים עם ניתוח AI
- `POST /api/user-feedback`: רישום משוב משתמש
- `POST /api/analyze-emails-ai`: ניתוח AI למיילים

## Meeting Management  
- `GET /api/meetings`: קבלת פגישות עם ניתוח AI
- `POST /api/analyze-meetings-ai`: ניתוח AI לפגישות

## Learning Management
- `GET /api/user-profile`: קבלת פרופיל משתמש
- `POST /api/update-preferences`: עדכון העדפות

## Backup & Maintenance
- `POST /api/create-backup`: יצירת גיבוי מלא""",
                
                "USER_GUIDE.md": """# מדריך משתמש

## התחלת עבודה
1. פתח את המערכת בדפדפן
2. בדוק חיבור ל-Outlook
3. רענן מיילים ופגישות

## ניהול מיילים
- צפייה במיילים עם ציוני חשיבות
- מתן משוב על חשיבות
- סימון קטגוריות
- ניתוח AI אוטומטי

## ניהול פגישות
- צפייה בפגישות עם ציוני חשיבות
- ניתוח AI לפגישות
- עדכון עדיפויות

## ניהול למידה
- צפייה בסטטיסטיקות למידה
- הגדרת העדפות
- ניתוח דפוסי למידה""",
                
                "DEVELOPER_GUIDE.md": """# מדריך מפתח

## מבנה הפרויקט
- `app_with_ai.py`: אפליקציית Flask הראשית
- `user_profile_manager.py`: ניהול פרופיל ולמידה
- `ai_analyzer.py`: מנוע ניתוח AI
- `templates/`: תבניות HTML
- `email_manager.db`: בסיס נתונים ראשי

## פיתוח
- השתמש ב-Python 3.8+
- התקן דרישות: `pip install -r requirements.txt`
- הפעל במצב debug: `python app_with_ai.py`

## תרומה לפרויקט
1. Fork את הפרויקט
2. צור branch חדש
3. בצע שינויים
4. שלח Pull Request""",
                
                "CHANGELOG.md": """# Changelog

## [Latest] - 2025-01-XX
### Added
- גיבוי מלא עם פרומפטים ותיעוד
- מודלים מפורטים לסטטיסטיקות למידה
- כפתורי עדיפות למיילים ופגישות
- מערכת למידה מתקדמת עם דפוסי זמן

### Changed
- שיפור חוויית משתמש במודלים
- אופטימיזציה של ניתוח AI
- שיפור ביצועים של בסיס הנתונים

### Fixed
- תיקון קריאות בערכה כהה
- תיקון חזרות בפעולות נדרשות
- שיפור יציבות החיבור ל-Outlook"""
            }
            
            for filename, content in docs_data.items():
                file_path = os.path.join(docs_folder, filename)
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                ui_block_add(block_id, f"   ✅ נוצר: {filename}", "INFO")
            
            ui_block_add(block_id, "✅ תיעוד נוצר בהצלחה", "SUCCESS")
            
        except Exception as docs_error:
            ui_block_add(block_id, f"⚠️ שגיאה ביצירת תיעוד: {str(docs_error)}", "WARNING")
        
        # שלב 3: יצירת גיבוי ZIP
        ui_block_add(block_id, "📦 שלב 3: יוצר גיבוי ZIP...", "INFO")
        
        # קבלת הסבר הגרסה מהבקשה
        data = request.get_json() or {}
        version_description = data.get('version_description', '').strip()
        
        # יצירת שם הקובץ עם תאריך ושעה
        now = datetime.now()
        timestamp = now.strftime("%d-%m-%Y_%H-%M")
        
        # הוספת הסבר הגרסה לשם הקובץ (אם קיים)
        if version_description:
            # המרת רווחים לקו תחתון והסרת תווים לא חוקיים
            safe_description = version_description.replace(' ', '_').replace('/', '_').replace('\\', '_').replace(':', '_')
            zip_filename = f"outlook_email_manager_{timestamp}_{safe_description}.zip"
            ui_block_add(block_id, f"📝 הסבר גרסה: {version_description}", "INFO")
        else:
            zip_filename = f"outlook_email_manager_{timestamp}.zip"
        
        # נתיב היעד
        downloads_path = r"c:\Users\ronni\Downloads"
        zip_path = os.path.join(downloads_path, zip_filename)
        
        # וידוא שהתיקייה קיימת
        os.makedirs(downloads_path, exist_ok=True)
        
        # נתיב הפרויקט הנוכחי
        project_path = os.getcwd()
        
        ui_block_add(block_id, f"📁 יוצר גיבוי מ: {project_path}", "INFO")
        ui_block_add(block_id, f"💾 שמירה ל: {zip_path}", "INFO")
        
        # יצירת ה-ZIP
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(project_path):
                # דילוג על תיקיות לא רצויות (כולל AIEmailManagerAddin)
                dirs[:] = [d for d in dirs if d not in ['__pycache__', '.git', 'node_modules', '.vscode', 'bin', 'obj', '.vs']]
                
                for file in files:
                    # דילוג על קבצים לא רצויים
                    if file.endswith(('.pyc', '.log', '.tmp', '.zip', '.pdb', '.suo', '.user', '.dll', '.exe')):
                        continue
                    
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, project_path)
                    zipf.write(file_path, arcname)
                    
        ui_block_add(block_id, "✅ גיבוי כולל את התוסף של C# (AIEmailManagerAddin)", "SUCCESS")
        
        # בדיקת גודל הקובץ
        file_size = os.path.getsize(zip_path)
        file_size_mb = file_size / (1024 * 1024)
        
        ui_block_add(block_id, f"📊 גודל הקובץ: {file_size_mb:.2f} MB", "INFO")
        ui_block_add(block_id, f"📁 מיקום: {zip_path}", "INFO")
        
        # שמירה ב-GitHub
        ui_block_add(block_id, "🔄 שומר שינויים ב-GitHub...", "INFO")
        try:
            
            # הוספת כל הקבצים
            result = subprocess.run(['git', 'add', '.'], capture_output=True, text=True, cwd=project_path)
            if result.returncode != 0:
                ui_block_add(block_id, f"⚠️ שגיאה בהוספת קבצים ל-Git: {result.stderr}", "WARNING")
            else:
                ui_block_add(block_id, "✅ קבצים נוספו ל-Git", "INFO")
            
            # יצירת קומיט
            commit_message = f"Backup: {zip_filename}"
            if version_description:
                commit_message = f"Backup: {version_description} ({zip_filename})"
            
            result = subprocess.run(['git', 'commit', '-m', commit_message], capture_output=True, text=True, cwd=project_path)
            if result.returncode != 0:
                ui_block_add(block_id, f"⚠️ שגיאה ביצירת קומיט: {result.stderr}", "WARNING")
            else:
                ui_block_add(block_id, "✅ קומיט נוצר בהצלחה", "INFO")
            
            # דחיפה ל-GitHub
            result = subprocess.run(['git', 'push'], capture_output=True, text=True, cwd=project_path)
            if result.returncode != 0:
                ui_block_add(block_id, f"⚠️ שגיאה בדחיפה ל-GitHub: {result.stderr}", "WARNING")
            else:
                ui_block_add(block_id, "✅ שינויים נדחפו ל-GitHub בהצלחה", "INFO")
                
        except Exception as git_error:
            ui_block_add(block_id, f"⚠️ שגיאה בשמירה ב-GitHub: {str(git_error)}", "WARNING")
        
        # סיכום כללי של כל התהליך
        ui_block_add(block_id, "🎉 סיכום תהליך הגיבוי המלא:", "SUCCESS")
        ui_block_add(block_id, "✅ פרומפטים ל-Cursor נוצרו בהצלחה", "SUCCESS")
        ui_block_add(block_id, "✅ תיעוד מעודכן נוצר בהצלחה", "SUCCESS")
        ui_block_add(block_id, f"✅ גיבוי ZIP נוצר: {zip_filename}", "SUCCESS")
        ui_block_add(block_id, f"✅ גודל הקובץ: {file_size_mb:.2f} MB", "SUCCESS")
        ui_block_add(block_id, "✅ שינויים נדחפו ל-GitHub בהצלחה", "SUCCESS")
        
        ui_block_end(block_id, "🎉 גיבוי מלא הושלם בהצלחה!", True)
        
        return jsonify({
            'success': True,
            'message': f'גיבוי נוצר בהצלחה!',
            'filename': zip_filename,
            'path': zip_path,
            'size_mb': round(file_size_mb, 2)
        })
        
    except Exception as e:
        error_msg = f'שגיאה ביצירת גיבוי: {str(e)}'
        try:
            ui_block_end(block_id, error_msg, False)
        except Exception:
            pass
        return jsonify({
            'success': False,
            'message': error_msg
        }), 500

@app.route('/api/create-cursor-prompts', methods=['POST'])
def create_cursor_prompts():
    """API ליצירת קבצי פרומפטים ל-Cursor"""
    try:
        block_id = ui_block_start("🧩 יצירת פרומפטים ל-Cursor")
        ui_block_add(block_id, "🚀 מתחיל יצירת קבצי פרומפטים ל-Cursor...", "INFO")
        
        # יצירת תיקיית פרומפטים בפרויקט
        project_path = os.getcwd()
        prompts_folder = os.path.join(project_path, "Cursor_Prompts")
        os.makedirs(prompts_folder, exist_ok=True)
        
        ui_block_add(block_id, f"📁 יוצר תיקיית פרומפטים: {prompts_folder}", "INFO")
        
        files_created = []
        
        # קובץ 0: הסברים על איך להשתמש בפרומפטים
        instructions_content = """# איך להשתמש בפרומפטים ל-Cursor - הוראות מפורטות

## 🎯 מטרה
הקבצים האלה מכילים פרומפטים מפורטים ליצירת מערכת ניהול מיילים חכמה עם AI באמצעות Cursor.

## 📋 רשימת הקבצים
1. **01_Main_Project_Prompt.txt** - פרומפט ראשי עם תיאור כללי
2. **02_Flask_Application.txt** - פרומפט לפיתוח Flask App
3. **03_Frontend_Development.txt** - פרומפט לפיתוח Frontend
4. **04_Outlook_Integration.txt** - פרומפט לאינטגרציה עם Outlook
5. **05_AI_Integration.txt** - פרומפט לאינטגרציה עם Gemini AI
6. **06_Deployment.txt** - פרומפט ל-Deployment והפעלה

## 🚀 איך להתחיל עם Cursor

### שלב 1: הכנת הסביבה
1. פתח Cursor
2. צור פרויקט חדש: `File > New Folder`
3. פתח את התיקייה החדשה ב-Cursor
4. צור קובץ `requirements.txt` עם התוכן מ-06_Deployment.txt

### שלב 2: יצירת הפרויקט הבסיסי
1. פתח את **01_Main_Project_Prompt.txt**
2. העתק את כל התוכן
3. ב-Cursor, לחץ `Ctrl+Shift+P` וחפש "Cursor: Chat"
4. הדבק את הפרומפט בצ'אט
5. Cursor יתחיל ליצור את הפרויקט הבסיסי

### שלב 3: פיתוח Flask App
1. פתח את **02_Flask_Application.txt**
2. העתק את התוכן
3. בצ'אט Cursor, בקש: "צור את קובץ app_with_ai.py לפי הפרומפט הזה"
4. הדבק את הפרומפט
5. Cursor ייצור את קובץ Flask המלא

### שלב 4: פיתוח Frontend
1. פתח את **03_Frontend_Development.txt**
2. העתק את התוכן
3. בצ'אט Cursor, בקש: "צור את קבצי HTML/CSS/JavaScript לפי הפרומפט"
4. הדבק את הפרומפט
5. Cursor ייצור את כל קבצי ה-Frontend

### שלב 5: אינטגרציה עם Outlook
1. פתח את **04_Outlook_Integration.txt**
2. העתק את התוכן
3. בצ'אט Cursor, בקש: "הוסף אינטגרציה עם Outlook לפי הפרומפט"
4. הדבק את הפרומפט
5. Cursor יוסיף את הקוד לחיבור Outlook

### שלב 6: אינטגרציה עם AI
1. פתח את **05_AI_Integration.txt**
2. העתק את התוכן
3. בצ'אט Cursor, בקש: "הוסף אינטגרציה עם Gemini AI לפי הפרומפט"
4. הדבק את הפרומפט
5. Cursor יוסיף את הקוד לניתוח AI

### שלב 7: Deployment
1. פתח את **06_Deployment.txt**
2. העתק את התוכן
3. בצ'אט Cursor, בקש: "צור קבצי deployment לפי הפרומפט"
4. הדבק את הפרומפט
5. Cursor ייצור את קבצי ההפעלה

## 💡 טיפים חשובים

### עבודה עם Cursor
- **השתמש בפרומפטים בסדר** - התחל מ-01 וסיים ב-06
- **הוסף הקשר** - תמיד תגיד ל-Cursor "לפי הפרומפט הזה"
- **בדוק את הקוד** - Cursor לא תמיד מושלם, בדוק את הקוד שנוצר
- **שאל שאלות** - אם משהו לא עובד, שאל את Cursor להסבר

### דרישות מערכת
- **Windows** עם Microsoft Outlook מותקן
- **Python 3.8+** מותקן
- **Cursor** מותקן ועודכן
- **API Key** של Google Gemini

### פתרון בעיות נפוצות
1. **Outlook לא נפתח** - ודא ש-Outlook מותקן ופתוח
2. **API Key לא עובד** - בדוק את המפתח ב-Google AI Studio
3. **Port תפוס** - שנה את הפורט ב-app.py מ-5000 ל-5001
4. **מודולים חסרים** - הרץ `pip install -r requirements.txt`

## 🎉 אחרי השלמת הפרויקט
1. הרץ `python app_with_ai.py`
2. פתח דפדפן ב-`http://localhost:5000`
3. בדוק שכל התכונות עובדות
4. התאם אישית לפי הצרכים שלך

## 📞 תמיכה
אם נתקלת בבעיות:
1. בדוק את הלוגים בקונסול
2. ודא שכל הדרישות מותקנות
3. נסה לפתור בעיה אחת בכל פעם
4. השתמש ב-Cursor Chat לשאלות נוספות

---
**בהצלחה בפיתוח! 🚀**
"""
        
        instructions_file = os.path.join(prompts_folder, "הסברים.txt")
        with open(instructions_file, 'w', encoding='utf-8') as f:
            f.write(instructions_content)
        files_created.append("הסברים.txt")
        
        # קובץ 1: פרומפט ראשי ליצירת הפרויקט
        main_prompt = """# Outlook Email Manager - Cursor Prompt

## תיאור הפרויקט
צור מערכת ניהול מיילים חכמה עם AI שמתחברת ל-Microsoft Outlook ומספקת ניתוח חכם של מיילים.

## דרישות טכניות
- Python Flask Framework
- Microsoft Outlook COM Integration (win32com.client)
- Google Gemini AI API
- SQLite3 Database
- HTML/CSS/JavaScript Frontend
- Responsive Design עם ערכה כהה/בהירה

## מבנה הפרויקט
```
outlook_email_manager/
├── app_with_ai.py          # Flask Application
├── templates/
│   ├── index.html          # דף ראשי - ניהול מיילים
│   ├── consol.html         # דף קונסול - לוגים
│   └── meetings.html       # דף פגישות
├── requirements.txt        # Dependencies
└── quick_start.ps1         # Script הפעלה
```

## תכונות עיקריות
1. **חיבור ל-Outlook** - קריאת מיילים ופגישות
2. **ניתוח AI** - שימוש ב-Gemini לניתוח חשיבות מיילים
3. **מערכת למידה** - שמירת העדפות משתמש
4. **ניהול פגישות** - הצגה וניתוח פגישות Outlook
5. **קונסול לוגים** - מעקב אחר פעילות המערכת
6. **גיבויים** - יצירת ZIP של הפרויקט
7. **ערכה כהה/בהירה** - החלפה בין ערכות

## הוראות פיתוח
1. התחל עם Flask app בסיסי
2. הוסף חיבור ל-Outlook COM
3. צור ממשק משתמש עם HTML/CSS/JavaScript
4. הוסף אינטגרציה עם Gemini AI
5. צור מערכת למידה עם SQLite
6. הוסף תכונות מתקדמות (פגישות, גיבויים, ערכות)

## קבצים נוספים
- requirements.txt עם כל ה-dependencies
- quick_start.ps1 להפעלה מהירה
- README.md עם הוראות התקנה ושימוש
"""
        
        main_file = os.path.join(prompts_folder, "01_Main_Project_Prompt.txt")
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(main_prompt)
        files_created.append("01_Main_Project_Prompt.txt")
        
        # קובץ 2: פרומפט ל-Flask App
        flask_prompt = """# Flask Application - app_with_ai.py

## מבנה Flask App
```python
from flask import Flask, render_template, jsonify, request, send_file
import win32com.client
import sqlite3
import json
from datetime import datetime
import os
import zipfile
import shutil

app = Flask(__name__)

# Global variables for console logs
# API Routes
@app.route('/')
def index():
    return render_template('index.html')

# Removed duplicate route - using the one at line 1503

@app.route('/meetings')
def meetings():
    return render_template('meetings.html')

# Email Management APIs
@app.route('/api/emails')
def get_emails():
    # קריאת מיילים מ-Outlook
    pass

@app.route('/api/stats')
def get_stats():
    # סטטיסטיקות מיילים
    pass

# Meeting Management APIs
@app.route('/api/meetings')
def get_meetings():
    # קריאת פגישות מ-Outlook
    pass

# Console APIs
# Backup APIs
@app.route('/api/create-backup', methods=['POST'])
def create_backup():
    # יצירת גיבוי ZIP
    pass

if __name__ == '__main__':
    app.config['TEMPLATES_AUTO_RELOAD'] = True
    
    # בדיקה אם קיימים קבצי SSL
    ssl_context = None
    if os.path.exists('server.crt') and os.path.exists('server.key'):
        ssl_context = ('server.crt', 'server.key')
        print("🔒 השרת רץ על HTTPS עם אישור SSL מקומי")
        print("🌐 כתובת: https://localhost:5000")
    else:
        print("⚠️ השרת רץ על HTTP (ללא SSL)")
        print("🌐 כתובת: http://localhost:5000")
    
    app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=True, ssl_context=ssl_context)
```

## EmailManager Class
```python
class EmailManager:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        
    def connect_to_outlook(self):
        # חיבור ל-Outlook
        pass
        
    def get_emails(self):
        # קריאת מיילים
        pass
        
    def analyze_emails_smart(self, emails):
        # ניתוח חכם של מיילים
        pass
```

## AI Integration
- שימוש ב-Google Gemini API
- ניתוח תוכן מיילים
- חישוב ציון חשיבות
- מערכת למידה מהמשוב
"""
        
        flask_file = os.path.join(prompts_folder, "02_Flask_Application.txt")
        with open(flask_file, 'w', encoding='utf-8') as f:
            f.write(flask_prompt)
        files_created.append("02_Flask_Application.txt")
        
        # קובץ 3: פרומפט ל-Frontend
        frontend_prompt = """# Frontend Development - HTML/CSS/JavaScript

## דף ראשי (index.html)
- כרטיסי סטטיסטיקות מיילים
- רשימת מיילים עם ניתוח AI
- כפתורי פעולה (רענון, ניתוח AI)
- ערכה כהה/בהירה
- עיצוב responsive

## דף קונסול (consol.html)
- הצגת לוגים בזמן אמת
- כפתורי בקרה (נקה, רענן, איפוס)
- יצירת גיבויים
- יצירת פרומפטים ל-Cursor
- ערכה כהה/בהירה

## דף פגישות (meetings.html)
- הצגת פגישות Outlook
- מערכת עדיפויות
- סינון לפי תאריכים
- ערכה כהה/בהירה

## CSS Features
- Gradients ו-animations
- Dark/Light mode toggle
- Responsive design
- Modern UI components

## JavaScript Features
- AJAX calls ל-APIs
- Real-time updates
- Local storage לעדפות
- Error handling
- Progress indicators

## Design System
- Colors: #667eea, #764ba2 (gradients)
- Dark mode: #1a1a2e, #16213e
- Typography: Segoe UI
- Icons: Emoji icons
- Layout: Flexbox/Grid
"""
        
        frontend_file = os.path.join(prompts_folder, "03_Frontend_Development.txt")
        with open(frontend_file, 'w', encoding='utf-8') as f:
            f.write(frontend_prompt)
        files_created.append("03_Frontend_Development.txt")
        
        # קובץ 4: פרומפט ל-Outlook Integration
        outlook_prompt = """# Outlook COM Integration

## חיבור ל-Outlook
```python
import win32com.client

class EmailManager:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        
    def connect_to_outlook(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            return True
        except Exception as e:
            log_to_console(f"שגיאה בחיבור ל-Outlook: {e}", "ERROR")
            return False
```

## קריאת מיילים
```python
def get_emails(self, limit=100):
    try:
        inbox = self.namespace.GetDefaultFolder(6)  # Inbox
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)  # Sort by date
        
        emails = []
        for i, message in enumerate(messages):
            if i >= limit:
                break
                
            email_data = {
                'id': i + 1,
                'subject': message.Subject,
                'sender': message.SenderName,
                'received_time': message.ReceivedTime,
                'body': message.Body,
                'is_read': message.UnRead == False,
                'importance': message.Importance
            }
            emails.append(email_data)
            
        return emails
    except Exception as e:
        log_to_console(f"שגיאה בקריאת מיילים: {e}", "ERROR")
        return []
```

## קריאת פגישות
```python
def get_meetings(self):
    try:
        calendar = self.namespace.GetDefaultFolder(9)  # Calendar
        appointments = calendar.Items
        
        meetings = []
        for appointment in appointments:
            meeting_data = {
                'id': appointment.EntryID,
                'subject': appointment.Subject,
                'start_time': appointment.Start,
                'end_time': appointment.End,
                'location': appointment.Location,
                'attendees': appointment.RequiredAttendees,
                'body': appointment.Body
            }
            meetings.append(meeting_data)
            
        return meetings
    except Exception as e:
        log_to_console(f"שגיאה בקריאת פגישות: {e}", "ERROR")
        return []
```

## טיפול בשגיאות
- Threading issues עם COM objects
- Datetime serialization
- Outlook permissions
- Error handling ו-fallback data
"""
        
        outlook_file = os.path.join(prompts_folder, "04_Outlook_Integration.txt")
        with open(outlook_file, 'w', encoding='utf-8') as f:
            f.write(outlook_prompt)
        files_created.append("04_Outlook_Integration.txt")
        
        # קובץ 5: פרומפט ל-AI Integration
        ai_prompt = """# AI Integration עם Google Gemini

## הגדרת Gemini API
```python
import google.generativeai as genai

# הגדרת API Key
genai.configure(api_key="YOUR_API_KEY")
model = genai.GenerativeModel('gemini-pro')
```

## ניתוח מיילים
```python
def analyze_email_with_ai(email_content, email_subject, sender):
    prompt = f\"\"\"
    נתח את החשיבות של המייל הבא:
    
    נושא: {email_subject}
    שולח: {sender}
    תוכן: {email_content}
    
    החזר ציון חשיבות בין 0-1 (0 = לא חשוב, 1 = קריטי)
    והסבר קצר למה.
    \"\"\"
    
    try:
        response = model.generate_content(prompt)
        # עיבוד התגובה וחילוץ הציון
        return parse_ai_response(response.text)
    except Exception as e:
        log_to_console(f"שגיאה בניתוח AI: {e}", "ERROR")
        return 0.5  # ציון ברירת מחדל
```

## מערכת למידה
```python
def learn_from_feedback(email_id, user_feedback, ai_score):
    # שמירת המשוב ב-SQLite
    conn = sqlite3.connect('learning.db')
    cursor = conn.cursor()
    
    cursor.execute('''
        INSERT INTO feedback (email_id, user_feedback, ai_score, timestamp)
        VALUES (?, ?, ?, ?)
    ''', (email_id, user_feedback, ai_score, datetime.now()))
    
    conn.commit()
    conn.close()
```

## Quota Management
- מעקב אחר שימוש ב-API
- הגבלת מספר בקשות
- Fallback לניתוח מקומי
- Caching של תוצאות

## Error Handling
- API rate limits
- Network errors
- Invalid responses
- Fallback mechanisms
"""
        
        ai_file = os.path.join(prompts_folder, "05_AI_Integration.txt")
        with open(ai_file, 'w', encoding='utf-8') as f:
            f.write(ai_prompt)
        files_created.append("05_AI_Integration.txt")
        
        # קובץ 6: פרומפט ל-Deployment
        deployment_prompt = """# Deployment והפעלה

## requirements.txt
```
Flask==2.3.3
pywin32==306
google-generativeai==0.3.2
requests==2.31.0
```

## quick_start.ps1
```powershell
# הפעלת השרת
python app_with_ai.py

# או עם virtual environment
python -m venv venv
venv\\Scripts\\activate
pip install -r requirements.txt
python app_with_ai.py
```

## הגדרות סביבה
- Windows עם Microsoft Outlook
- Python 3.8+
- Internet connection ל-Gemini API
- Outlook permissions

## Troubleshooting
- Outlook COM errors
- API key issues
- Port conflicts
- Permission problems

## Security
- API key protection
- Input validation
- Error handling
- Logging

## Performance
- Caching strategies
- Database optimization
- Async operations
- Memory management
"""
        
        deployment_file = os.path.join(prompts_folder, "06_Deployment.txt")
        with open(deployment_file, 'w', encoding='utf-8') as f:
            f.write(deployment_prompt)
        files_created.append("06_Deployment.txt")
        
        # קובץ README.md לתיקיית הפרומפטים
        readme_content = """# Cursor Prompts - Outlook Email Manager

## 📁 תוכן התיקייה
תיקייה זו מכילה פרומפטים מפורטים ליצירת מערכת ניהול מיילים חכמה עם AI באמצעות Cursor.

## 📋 קבצים
- **הסברים.txt** - הוראות מפורטות לשימוש
- **01_Main_Project_Prompt.txt** - פרומפט ראשי
- **02_Flask_Application.txt** - פרומפט Flask
- **03_Frontend_Development.txt** - פרומפט Frontend
- **04_Outlook_Integration.txt** - פרומפט Outlook
- **05_AI_Integration.txt** - פרומפט AI
- **06_Deployment.txt** - פרומפט Deployment

## 🚀 התחלה מהירה
1. פתח את **הסברים.txt**
2. עקוב אחר ההוראות המפורטות
3. התחל עם קובץ 01
4. המשך בסדר עד קובץ 06

## 💡 טיפ
השתמש בפרומפטים בסדר המספרי לקבלת התוצאות הטובות ביותר!

---
נוצר על ידי: Outlook Email Manager System
תאריך: """ + datetime.now().strftime("%d/%m/%Y %H:%M") + """
"""
        
        readme_file = os.path.join(prompts_folder, "README.md")
        with open(readme_file, 'w', encoding='utf-8') as f:
            f.write(readme_content)
        files_created.append("README.md")
        
        ui_block_add(block_id, f"📁 תיקייה: {prompts_folder}", "INFO")
        ui_block_add(block_id, f"📄 {len(files_created)} קבצים נוצרו", "INFO")
        ui_block_add(block_id, f"📖 קובץ הסברים: הסברים.txt", "INFO")
        ui_block_end(block_id, "קבצי פרומפטים נוצרו בהצלחה", True)
        
        return jsonify({
            'success': True,
            'message': 'קבצי פרומפטים נוצרו בהצלחה!',
            'folder_path': prompts_folder,
            'files_created': files_created
        })
        
    except Exception as e:
        error_msg = f'שגיאה ביצירת קבצי פרומפטים: {str(e)}'
        try:
            ui_block_end(block_id, error_msg, False)
        except Exception:
            pass
        return jsonify({
            'success': False,
            'message': error_msg
        }), 500

@app.route('/api/status', methods=['GET'])
def api_status():
    """API לבדיקת סטטוס השרת"""
    return jsonify({
        'status': 'running',
        'message': 'השרת פועל בהצלחה',
        'timestamp': datetime.now().isoformat()
    })

@app.route('/api/setup-outlook-addin', methods=['POST'])
def setup_outlook_addin():
    """API להגדרת תוסף Outlook - התקנה מלאה"""
    try:
        block_id = ui_block_start("🔌 התקנת תוסף Outlook")
        ui_block_add(block_id, "🚀 מתחיל התקנת תוסף Outlook...", "INFO")
        
        # שלב 1: בדיקת חיבור ל-Outlook
        ui_block_add(block_id, "📝 שלב 1: בודק חיבור ל-Outlook...", "INFO")
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)
            ui_block_add(block_id, "✅ חיבור ל-Outlook הצליח!", "SUCCESS")
        except Exception as e:
            ui_block_add(block_id, f"❌ שגיאה בחיבור ל-Outlook: {e}", "ERROR")
            ui_block_end(block_id, "התקנה נכשלה - לא ניתן להתחבר ל-Outlook", False)
            return jsonify({'success': False, 'error': str(e)}), 500
        
        # שלב 2: יצירת עמודות מותאמות אישית ב-Outlook
        ui_block_add(block_id, "📊 שלב 2: יוצר עמודות מותאמות אישית...", "INFO")
        try:
            # בדיקה אם העמודות כבר קיימות
            test_items = inbox.Items
            if test_items.Count > 0:
                test_item = test_items[1]
                
                # יצירת AISCORE (מספר)
                try:
                    aiscore_prop = test_item.UserProperties.Find("AISCORE")
                    if not aiscore_prop:
                        aiscore_prop = test_item.UserProperties.Add("AISCORE", 3, True)  # 3 = olNumber
                        test_item.Save()
                        ui_block_add(block_id, "✅ עמודת AISCORE נוצרה (מספר)", "SUCCESS")
                    else:
                        ui_block_add(block_id, "ℹ️ עמודת AISCORE כבר קיימת", "INFO")
                except Exception as e:
                    ui_block_add(block_id, f"⚠️ שגיאה ביצירת AISCORE: {e}", "WARNING")
                
                # יצירת AICategory (טקסט) - ללא קו תחתון!
                try:
                    category_prop = test_item.UserProperties.Find("AICategory")
                    if not category_prop:
                        category_prop = test_item.UserProperties.Add("AICategory", 1, True)  # 1 = olText
                        test_item.Save()
                        ui_block_add(block_id, "✅ עמודת AICategory נוצרה (טקסט)", "SUCCESS")
                    else:
                        ui_block_add(block_id, "ℹ️ עמודת AICategory כבר קיימת", "INFO")
                except Exception as e:
                    ui_block_add(block_id, f"⚠️ שגיאה ביצירת AICategory: {e}", "WARNING")
                
                # יצירת AISummary (טקסט) - ללא קו תחתון!
                try:
                    summary_prop = test_item.UserProperties.Find("AISummary")
                    if not summary_prop:
                        summary_prop = test_item.UserProperties.Add("AISummary", 1, True)  # 1 = olText
                        test_item.Save()
                        ui_block_add(block_id, "✅ עמודת AISummary נוצרה (טקסט)", "SUCCESS")
                    else:
                        ui_block_add(block_id, "ℹ️ עמודת AISummary כבר קיימת", "INFO")
                except Exception as e:
                    ui_block_add(block_id, f"⚠️ שגיאה ביצירת AISummary: {e}", "WARNING")
                    
            else:
                ui_block_add(block_id, "⚠️ אין מיילים ב-Inbox ליצירת עמודות", "WARNING")
                ui_block_add(block_id, "ℹ️ העמודות ייווצרו אוטומטית בניתוח המייל הראשון", "INFO")
            
        except Exception as e:
            ui_block_add(block_id, f"⚠️ שגיאה ביצירת עמודות: {e}", "WARNING")
        
        # שלב 3: רישום התוסף COM
        ui_block_add(block_id, "🔧 שלב 3: רושם תוסף COM...", "INFO")
        try:
            project_path = os.getcwd()
            addin_path = os.path.join(project_path, "outlook_com_addin_final.py")
            
            if os.path.exists(addin_path):
                ui_block_add(block_id, f"📁 מצא קובץ תוסף: {addin_path}", "INFO")
                
                # רישום התוסף
                result = subprocess.run(
                    ['python', addin_path, '--register'],
                    capture_output=True,
                    text=True,
                    cwd=project_path
                )
                
                if result.returncode == 0:
                    ui_block_add(block_id, "✅ תוסף COM נרשם בהצלחה!", "SUCCESS")
                else:
                    ui_block_add(block_id, f"⚠️ שגיאה ברישום COM: {result.stderr}", "WARNING")
            else:
                ui_block_add(block_id, f"⚠️ קובץ התוסף לא נמצא: {addin_path}", "WARNING")
                
        except Exception as e:
            ui_block_add(block_id, f"⚠️ שגיאה ברישום COM: {e}", "WARNING")
        
        # שלב 4: הוראות סיום
        ui_block_add(block_id, "📋 שלב 4: הוראות סיום...", "INFO")
        ui_block_add(block_id, "ℹ️ כדי להציג את עמודת AISCORE:", "INFO")
        ui_block_add(block_id, "   1. סגור Outlook אם פתוח", "INFO")
        ui_block_add(block_id, "   2. פתח Outlook מחדש", "INFO")
        ui_block_add(block_id, "   3. עבור ל-View → View Settings → Columns", "INFO")
        ui_block_add(block_id, "   4. בחר 'User Defined Fields in Folder'", "INFO")
        ui_block_add(block_id, "   5. הוסף את השדה AISCORE לתצוגה", "INFO")
        
        ui_block_end(block_id, "✅ התקנת תוסף Outlook הושלמה!", True)
        
        return jsonify({
            'success': True,
            'message': 'תוסף Outlook הותקן בהצלחה!'
        })
        
    except Exception as e:
        error_msg = f'שגיאה בהתקנת תוסף Outlook: {str(e)}'
        try:
            ui_block_end(block_id, error_msg, False)
        except Exception:
            pass
        return jsonify({
            'success': False,
            'message': error_msg
        }), 500

@app.route('/api/transfer-scores-to-outlook', methods=['POST'])
def transfer_scores_to_outlook():
    """API להעברת ציונים ל-Outlook"""
    try:
        # מניעת הרצה מקבילה של העברת ציונים
        global cached_data
        if cached_data.get('is_transferring_scores'):
            return jsonify({
                'success': False,
                'message': 'פעולת העברת ציונים כבר רצה. נא להמתין לסיומה.'
            }), 429
        cached_data['is_transferring_scores'] = True
        # בלוק לוג מפורש עבור העברת ציונים
        block_id = ui_block_start("📝 העברת ציונים ל-Outlook")
        ui_block_add(block_id, "🚀 מתחיל העברת ציונים ל-Outlook...", "INFO")
        
        # בדיקה שיש נתונים זמינים
        if not cached_data['emails']:
            ui_block_add(block_id, "❌ אין מיילים זמינים להעברה", "ERROR")
            return jsonify({
                'success': False,
                'message': 'אין מיילים זמינים להעברה. נא לטעון את המיילים קודם.'
            }), 400
        
        emails_processed = 0
        emails_success = 0
        emails_failed = 0
        
        ui_block_add(block_id, f"📧 נמצאו {len(cached_data['emails'])} מיילים עם ציונים מוכנים", "INFO")
        
        # עיבוד המיילים (כל המיילים)
        max_emails = len(cached_data['emails'])
        
        ui_block_add(block_id, f"⚡ מעבד {max_emails} מיילים (כל המיילים)", "INFO")
        
        # בדיקת חיבור ל-Outlook
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            ui_block_add(block_id, "✅ חיבור ל-Outlook הצליח!", "SUCCESS")
        except Exception as e:
            ui_block_add(block_id, f"❌ שגיאה בחיבור ל-Outlook: {e}", "ERROR")
            return jsonify({'success': False, 'error': str(e)})
        if not outlook:
            ui_block_add(block_id, "❌ לא ניתן להתחבר ל-Outlook", "ERROR")
            return jsonify({
                'success': False,
                'message': 'לא ניתן להתחבר ל-Outlook'
            }), 500
        
        ui_block_add(block_id, "✅ חיבור ל-Outlook הצליח!", "SUCCESS")
        
        # קבלת כל המיילים מ-Outlook
        try:
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)  # מיון לפי זמן קבלה
            
            ui_block_add(block_id, f"📧 נמצאו {messages.Count} מיילים ב-Outlook", "INFO")
            
            for i in range(max_emails):
                try:
                    # בדיקה שהמייל קיים
                    if i + 1 > messages.Count:
                        ui_block_add(block_id, f"⚠️ מייל {i+1} לא קיים (רק {messages.Count} מיילים)", "WARNING")
                        break
                    
                    message = messages[i + 1]  # Outlook מתחיל מ-1, לא מ-0
                    emails_processed += 1
                    
                    # שימוש בציונים שכבר מחושבים מהזיכרון
                    email_from_cache = cached_data['emails'][i]
                    
                    # יצירת analysis object מהנתונים הקיימים
                    analysis = {
                        'importance_score': email_from_cache.get('importance_score', 0.5),
                        'category': email_from_cache.get('category', 'work'),
                        'summary': f"מייל מ-{email_from_cache.get('sender', 'לא ידוע')}: {email_from_cache.get('subject', 'ללא נושא')}",
                        'action_items': []
                    }
                    
                    # הוספת הניתוח למייל ב-Outlook
                    try:
                        importance_percent = int(analysis['importance_score'] * 100)
                        
                        # הוספת AISCORE כמספר (לתצוגה בעמודה)
                        try:
                            score_prop = message.UserProperties.Find("AISCORE")
                            if not score_prop:
                                score_prop = message.UserProperties.Add("AISCORE", 3, True)  # 3 = olNumber
                            if score_prop:
                                score_prop.Value = importance_percent
                        except Exception as e:
                            ui_block_add(block_id, f"❌ שגיאה ב-AISCORE: {e}", "ERROR")
                        
                        # הוספת AICategory כטקסט (ללא קו תחתון!)
                        try:
                            category_prop = message.UserProperties.Find("AICategory")
                            if not category_prop:
                                category_prop = message.UserProperties.Add("AICategory", 1, True)  # 1 = olText
                            if category_prop:
                                category_prop.Value = analysis['category']
                        except Exception as e:
                            ui_block_add(block_id, f"❌ שגיאה ב-AICategory: {e}", "ERROR")
                        
                        # הוספת AISummary כטקסט (ללא קו תחתון!)
                        try:
                            summary_prop = message.UserProperties.Find("AISummary")
                            if not summary_prop:
                                summary_prop = message.UserProperties.Add("AISummary", 1, True)  # 1 = olText
                            if summary_prop:
                                summary_text = analysis.get('summary', '')[:255]  # מוגבל ל-255 תווים
                                summary_prop.Value = summary_text
                        except Exception as e:
                            ui_block_add(block_id, f"❌ שגיאה ב-AISummary: {e}", "ERROR")
                        
                        # שמירה
                        message.Save()
                        emails_success += 1
                        score_percent = int(analysis['importance_score'] * 100)
                        ui_block_add(block_id, f"✅ מייל {i+1}: {email_from_cache['subject']} - ציון: {score_percent}%", "SUCCESS")
                    except Exception as e:
                        emails_failed += 1
                        ui_block_add(block_id, f"❌ שגיאה במייל {i+1}: {e}", "ERROR")
                    
                except Exception as e:
                    emails_failed += 1
                    ui_block_add(block_id, f"❌ שגיאה במייל {i+1}: {e}", "ERROR")
                    
        except Exception as e:
            error_msg = f'שגיאה בעיבוד מיילים: {str(e)}'
            ui_block_add(block_id, error_msg, "ERROR")
            return jsonify({
                'success': False,
                'message': error_msg
            }), 500
        
        ui_block_end(block_id, f"✅ העברת ציונים הושלמה! עובדו: {emails_processed}, הצליחו: {emails_success}, נכשלו: {emails_failed}", True)
        
        response = jsonify({
            'success': True,
            'message': 'ציונים הועברו ל-Outlook בהצלחה',
            'emails_processed': emails_processed,
            'emails_success': emails_success,
            'emails_failed': emails_failed
        })
        cached_data['is_transferring_scores'] = False
        return response
        
    except Exception as e:
        error_msg = f'שגיאה בהעברת ציונים ל-Outlook: {str(e)}'
        try:
            ui_block_end(block_id, error_msg, False)
        except Exception:
            pass
        try:
            cached_data['is_transferring_scores'] = False
        except Exception:
            pass
        return jsonify({
            'success': False,
            'message': error_msg
        }), 500

@app.route('/api/analyze-email', methods=['POST'])
def analyze_single_email():
    """API לניתוח מייל בודד ועדכון Outlook"""
    try:
        email_data = request.json
        entry_id = email_data.get('entryID')  # מזהה המייל ב-Outlook
        
        # יצירת EmailManager
        email_manager = EmailManager()
        
        # ניתוח המייל
        analysis = email_manager.analyze_single_email(email_data)
        
        # הודעה נקייה עם הציון
        score_percent = int(analysis.get('importance_score', 0.5) * 100)
        subject = email_data.get('subject', 'ללא נושא')[:50]
        
        # עדכון Outlook אם יש entry_id
        outlook_update_success = False
        outlook_error_msg = None
        if entry_id:
            try:
                pythoncom.CoInitialize()
                outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                mail_item = outlook.GetItemFromID(entry_id)
                
                # עדכון PRIORITYNUM
                priority_prop = mail_item.UserProperties.Find("PRIORITYNUM")
                if not priority_prop:
                    priority_prop = mail_item.UserProperties.Add("PRIORITYNUM", 3)  # 3 = olNumber
                priority_prop.Value = score_percent
                
                # עדכון AISCORE
                aiscore_prop = mail_item.UserProperties.Find("AISCORE")
                if not aiscore_prop:
                    aiscore_prop = mail_item.UserProperties.Add("AISCORE", 1)  # 1 = olText
                aiscore_prop.Value = f"{score_percent}%"
                
                mail_item.Save()
                pythoncom.CoUninitialize()
                outlook_update_success = True
                print(f"✅ PRIORITYNUM עודכן בהצלחה ל-{score_percent} למייל: {subject}")
            except Exception as outlook_error:
                # אם יש שגיאה בעדכון Outlook - רושמים בלוג אבל ממשיכים
                outlook_error_msg = str(outlook_error)
                print(f"❌ שגיאה בעדכון PRIORITYNUM: {outlook_error_msg}")
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass
        
        return jsonify({
            **analysis,
            'success': True,
            'message': f'✅ ניתוח הושלם בהצלחה: {score_percent}%',
            'score_display': f'{score_percent}%',
            'priority_updated': outlook_update_success,
            'outlook_error': outlook_error_msg
        })
        
    except Exception as e:
        error_msg = f'❌ שגיאה בניתוח מייל: {str(e)}'
        return jsonify({
            'success': False,
            'error': error_msg
        }), 500

@app.route('/api/create-documentation', methods=['POST'])
def create_documentation():
    """API ליצירת/רענון קבצי תיעוד MD עם תרשימי Mermaid"""
    try:
        block_id = ui_block_start("📚 יצירת/רענון תיעוד")
        ui_block_add(block_id, "🚀 מתחיל יצירת/רענון קבצי תיעוד...", "INFO")
        
        # יצירת תיקיית תיעוד בפרויקט
        project_path = os.getcwd()
        docs_folder = os.path.join(project_path, "docs")
        os.makedirs(docs_folder, exist_ok=True)
        
        ui_block_add(block_id, f"📁 יוצר תיקיית תיעוד: {docs_folder}", "INFO")
        
        files_created = []
        
        # קובץ README.md
        readme_content = """# 📧 Outlook Email Manager with AI

מערכת ניהול מיילים חכמה המשלבת Microsoft Outlook עם בינה מלאכותית לניתוח אוטומטי של חשיבות המיילים וניהול פגישות.

## 🌟 תכונות עיקריות

### 📧 ניהול מיילים חכם
- **ניתוח AI אוטומטי** - ניתוח חשיבות המיילים עם Gemini AI
- **סינון חכם** - מיילים קריטיים, חשובים, בינוניים ונמוכים
- **משוב משתמש** - מערכת למידה מהמשוב שלך
- **ניתוח קטגוריות** - זיהוי אוטומטי של סוגי מיילים

### 📅 ניהול פגישות
- **סינכרון Outlook** - טעינה אוטומטית של פגישות
- **כפתורי עדיפות** - סימון עדיפות פגישות עם LED חזותי
- **סטטיסטיקות** - ניתוח דפוסי פגישות
- **ניהול למידה** - מערכת למידה מתקדמת

### 🖥️ קונסול ניהול
- **מעקב בזמן אמת** - לוגים חיים של פעילות המערכת
- **ניהול שרת** - הפעלה מחדש וגיבויים
- **פרומפטים ל-Cursor** - יצירת קבצי עזר לפיתוח
- **יצירת תיעוד** - יצירת/רענון קבצי MD עם תרשימי Mermaid

## 🚀 התחלה מהירה

### דרישות מערכת
- Windows 10/11
- Python 3.8+
- Microsoft Outlook
- Google Gemini API Key

### התקנה מהירה
```powershell
# הפעלת הפרויקט
.\\quick_start.ps1
```

### הפעלה ידנית
```powershell
# התקנת תלויות
pip install -r requirements.txt

# הפעלת השרת
python app_with_ai.py
```

## 📁 מבנה הפרויקט

```mermaid
graph TD
    A[📧 Outlook Email Manager] --> B[🐍 Backend Flask]
    A --> C[🎨 Frontend HTML/CSS/JS]
    A --> D[🤖 AI Engine]
    A --> E[💾 Database]
    
    B --> B1[app_with_ai.py]
    B --> B2[ai_analyzer.py]
    B --> B3[user_profile_manager.py]
    B --> B4[config.py]
    
    C --> C1[📧 index.html]
    C --> C2[📅 meetings.html]
    C --> C3[🖥️ consol.html]
    
    D --> D1[Google Gemini API]
    D --> D2[AI Analysis]
    D --> D3[Learning System]
    
    E --> E1[email_manager.db]
    E --> E2[email_preferences.db]
    
    F[📚 Documentation] --> F1[README.md]
    F --> F2[INSTALLATION.md]
    F --> F3[USER_GUIDE.md]
    F --> F4[API_DOCUMENTATION.md]
    F --> F5[DEVELOPER_GUIDE.md]
    F --> F6[CHANGELOG.md]
```

### 📂 מבנה קבצים
```
outlook_email_manager/
├── 📧 app_with_ai.py          # אפליקציה ראשית
├── 🤖 ai_analyzer.py          # מנוע AI
├── 👤 user_profile_manager.py # ניהול פרופיל משתמש
├── 📄 config.py               # הגדרות
├── 📁 templates/              # תבניות HTML
│   ├── index.html            # דף ניהול מיילים
│   ├── meetings.html         # דף ניהול פגישות
│   └── consol.html           # דף קונסול
├── 📁 docs/                  # תיעוד מפורט
├── 📁 Cursor_Prompts/        # פרומפטים לפיתוח
└── 📁 Old/                   # קבצים ישנים
```

## 📖 מדריכים מפורטים

- [📋 מדריך התקנה מפורט](INSTALLATION.md)
- [👤 מדריך משתמש](USER_GUIDE.md)
- [🔧 מדריך מפתח](DEVELOPER_GUIDE.md)
- [🌐 תיעוד API](API_DOCUMENTATION.md)
- [📝 יומן שינויים](CHANGELOG.md)

## 🔧 הגדרה

### 1. הגדרת Outlook
- התקן Microsoft Outlook
- התחבר לחשבון שלך
- הפעל את הפרויקט

### 2. הגדרת AI
- קבל API Key מ-Google Gemini
- הוסף את המפתח לקובץ `config.py`
- הפעל את המערכת

### 3. הגדרת בסיס נתונים
- המערכת יוצרת אוטומטית את בסיס הנתונים
- נתונים נשמרים ב-`email_manager.db`

## 🤝 תרומה לפרויקט

1. Fork את הפרויקט
2. צור branch חדש (`git checkout -b feature/amazing-feature`)
3. Commit את השינויים (`git commit -m 'Add amazing feature'`)
4. Push ל-branch (`git push origin feature/amazing-feature`)
5. פתח Pull Request

## 📝 רישיון

פרויקט זה מופץ תחת רישיון MIT. ראה קובץ `LICENSE` לפרטים נוספים.

## 📞 תמיכה

- 🐛 דיווח באגים: פתח Issue חדש
- 💡 הצעות תכונות: פתח Issue עם תווית "enhancement"
- 📧 שאלות: צור קשר דרך Issues

## 🏆 הישגים

- ✅ אינטגרציה מלאה עם Microsoft Outlook
- ✅ ניתוח AI מתקדם עם Gemini
- ✅ ממשק משתמש אינטואיטיבי
- ✅ מערכת למידה אדפטיבית
- ✅ ניהול פגישות חכם
- ✅ קונסול ניהול מתקדם
- ✅ תיעוד מפורט עם תרשימי Mermaid

---

**פותח עם ❤️ בישראל** 🇮🇱
"""
        
        readme_file = os.path.join(docs_folder, "README.md")
        with open(readme_file, 'w', encoding='utf-8') as f:
            f.write(readme_content)
        files_created.append("README.md")
        
        # קובץ INSTALLATION.md
        installation_content = """# 📋 מדריך התקנה מפורט

מדריך שלב-אחר-שלב להתקנת Outlook Email Manager with AI.

## 🔧 דרישות מערכת

### חומרה
- **מעבד**: Intel Core i3 או AMD Ryzen 3 ומעלה
- **זיכרון**: 4GB RAM (מומלץ 8GB)
- **אחסון**: 500MB מקום פנוי
- **מערכת הפעלה**: Windows 10/11

### תוכנה
- **Python 3.8+** - [הורדה](https://www.python.org/downloads/)
- **Microsoft Outlook** - גרסה 2016 ומעלה
- **Git** (אופציונלי) - [הורדה](https://git-scm.com/)

## 🚀 התקנה מהירה

### שלב 1: הורדת הפרויקט
```bash
# דרך Git
git clone https://github.com/your-repo/outlook-email-manager.git
cd outlook-email-manager

# או הורדה ישירה
# הורד את הקובץ ZIP ופתח אותו
```

### שלב 2: התקנת Python
1. הורד Python מ-[python.org](https://www.python.org/downloads/)
2. התקן עם אפשרות "Add to PATH"
3. בדוק התקנה:
```bash
python --version
pip --version
```

### שלב 3: התקנת תלויות
```bash
pip install -r requirements.txt
```

### שלב 4: הגדרת Gemini AI
1. עבור ל-[Google AI Studio](https://makersuite.google.com/app/apikey)
2. צור API Key חדש
3. העתק את המפתח
4. פתח את `config.py` והוסף:
```python
GEMINI_API_KEY = "your-api-key-here"
```

### שלב 5: הפעלה
```bash
python app_with_ai.py
```

## 🔧 התקנה ידנית מפורטת

### שלב 1: הכנת הסביבה

#### בדיקת Python
```bash
python --version
# צריך להציג Python 3.8.0 או גרסה חדשה יותר
```

#### יצירת סביבה וירטואלית (מומלץ)
```bash
python -m venv outlook_manager_env
outlook_manager_env\\Scripts\\activate
```

### שלב 2: התקנת חבילות

#### חבילות בסיסיות
```bash
pip install flask==2.3.3
pip install flask-cors==4.0.0
pip install pywin32>=307
pip install google-generativeai==0.3.2
```

#### או התקנה מקובץ requirements
```bash
pip install -r requirements.txt
```

### שלב 3: הגדרת Outlook

#### בדיקת Outlook
1. פתח Microsoft Outlook
2. התחבר לחשבון שלך
3. ודא שיש לך גישה למיילים ופגישות

#### הרשאות COM
- Outlook צריך להיות פתוח בעת הפעלת הפרויקט
- ודא שאין חסימות אנטי-וירוס ל-COM objects

### שלב 4: הגדרת AI

#### קבלת API Key
1. עבור ל-[Google AI Studio](https://makersuite.google.com/app/apikey)
2. התחבר עם חשבון Google
3. לחץ "Create API Key"
4. העתק את המפתח

#### הגדרת המפתח
```python
# בקובץ config.py
GEMINI_API_KEY = "AIzaSyBOUWyZ-Dq2yPopzSZ6oopN7V6oeoB2iNY"  # המפתח שלך
```

### שלב 5: בדיקת התקנה

#### בדיקת חיבורים
```bash
python -c "import win32com.client; print('Outlook COM: OK')"
python -c "import google.generativeai; print('Gemini AI: OK')"
```

#### הפעלת השרת
```bash
python app_with_ai.py
```

#### בדיקת דפדפן
פתח דפדפן ב-`http://localhost:5000`

## 🐛 פתרון בעיות נפוצות

### בעיה: Python לא נמצא
```bash
# פתרון: הוסף Python ל-PATH
# או השתמש בנתיב המלא
C:\\Python39\\python.exe app_with_ai.py
```

### בעיה: Outlook לא נפתח
- ודא ש-Outlook מותקן ופתוח
- בדוק שאין חסימות אנטי-וירוס
- נסה להפעיל את Outlook כמנהל

### בעיה: API Key לא עובד
- בדוק שהמפתח תקין ב-Google AI Studio
- ודא שיש לך quota זמין
- בדוק את החיבור לאינטרנט

### בעיה: Port תפוס
```bash
# שנה את הפורט בקובץ app_with_ai.py
app.run(host='0.0.0.0', port=5001)  # במקום 5000
```

### בעיה: מודולים חסרים
```bash
pip install --upgrade pip
pip install -r requirements.txt --force-reinstall
```

## 🔄 עדכון הפרויקט

### עדכון דרך Git
```bash
git pull origin main
pip install -r requirements.txt --upgrade
```

### עדכון ידני
1. הורד את הגרסה החדשה
2. החלף את הקבצים הישנים
3. התקן תלויות חדשות:
```bash
pip install -r requirements.txt --upgrade
```

## 📞 תמיכה טכנית

אם נתקלת בבעיות:

1. **בדוק את הלוגים** - פתח את הקונסול ב-`http://localhost:5000/consol`
2. **בדוק דרישות** - ודא שכל הדרישות מותקנות
3. **נסה פתרון אחד** - פתור בעיה אחת בכל פעם
4. **דווח על באג** - פתח Issue עם פרטי השגיאה

## 🎯 שלבים הבאים

לאחר התקנה מוצלחת:

1. 📖 קרא את [מדריך המשתמש](USER_GUIDE.md)
2. 🔧 עיין ב-[מדריך המפתח](DEVELOPER_GUIDE.md)
3. 🌐 בדוק את [תיעוד ה-API](API_DOCUMENTATION.md)
4. 🚀 התחל להשתמש במערכת!

---

**בהצלחה בהתקנה! 🎉**
"""
        
        installation_file = os.path.join(docs_folder, "INSTALLATION.md")
        with open(installation_file, 'w', encoding='utf-8') as f:
            f.write(installation_content)
        files_created.append("INSTALLATION.md")
        
        # קובץ API_DOCUMENTATION.md
        api_content = """# 🌐 תיעוד API מפורט

תיעוד מלא של כל ה-API endpoints ב-Outlook Email Manager with AI.

## 📋 סקירה כללית

המערכת מספקת REST API מלא לניהול מיילים, פגישות ו-AI analysis.

### תרשים API Endpoints

```mermaid
graph TD
    A[🌐 API Base URL: localhost:5000] --> B[📧 Email APIs]
    A --> C[📅 Meeting APIs]
    A --> D[🤖 AI APIs]
    A --> E[📊 Learning APIs]
    A --> F[🔧 System APIs]
    A --> G[🖥️ Console APIs]
    A --> H[📦 Backup APIs]
    
    B --> B1[GET /api/emails]
    B --> B2[POST /api/refresh-data]
    B --> B3[GET /api/stats]
    B --> B4[POST /api/user-feedback]
    B --> B5[POST /api/analyze-emails-ai]
    
    C --> C1[GET /api/meetings]
    C --> C2[POST /api/meetings/:id/priority]
    C --> C3[GET /api/meetings/stats]
    C --> C4[POST /api/analyze-meetings-ai]
    
    D --> D1[GET /api/ai-status]
    D --> D2[POST /api/analyze-emails-ai]
    D --> D3[POST /api/analyze-meetings-ai]
    
    E --> E1[GET /api/learning-stats]
    E --> E2[GET /api/learning-management]
    
    F --> F1[GET /api/test-outlook]
    F --> F2[GET /api/server-id]
    F --> F3[POST /api/restart-server]
    
    G --> G1[GET /api/console-logs]
    G --> G2[POST /api/clear-console]
    G --> G3[POST /api/console-reset]
    
    H --> H1[POST /api/create-backup]
    H --> H2[POST /api/create-cursor-prompts]
    H --> H3[POST /api/create-documentation]
```

**Base URL**: `http://localhost:5000`

**Content-Type**: `application/json`

## 📧 API מיילים

### GET /api/emails
מחזיר את כל המיילים מהזיכרון.

**Response**:
```json
[
  {
    "id": "email_123",
    "subject": "נושא המייל",
    "sender": "שולח",
    "sender_email": "sender@example.com",
    "received_time": "2025-09-30T10:30:00Z",
    "body_preview": "תצוגה מקדימה של התוכן...",
    "is_read": false,
    "importance_score": 0.85,
    "category": "work",
    "summary": "סיכום המייל",
    "action_items": ["פעולה 1", "פעולה 2"]
  }
]
```

### POST /api/refresh-data
מרענן את הנתונים מהזיכרון.

**Request**:
```json
{
  "type": "emails"  // או "meetings" או null לכל הנתונים
}
```

**Response**:
```json
{
  "success": true,
  "message": "נתונים עודכנו בהצלחה",
  "last_updated": "2025-09-30T10:35:00Z"
}
```

### GET /api/stats
מחזיר סטטיסטיקות מיילים.

**Response**:
```json
{
  "total_emails": 150,
  "unread_emails": 25,
  "critical_emails": 5,
  "high_priority_emails": 15,
  "medium_priority_emails": 50,
  "low_priority_emails": 80,
  "categories": {
    "work": 80,
    "personal": 40,
    "marketing": 20,
    "system": 10
  }
}
```

### POST /api/user-feedback
שולח משוב משתמש על ניתוח AI.

**Request**:
```json
{
  "email_id": "email_123",
  "feedback": "high",  // "high", "medium", "low"
  "ai_score": 0.85
}
```

**Response**:
```json
{
  "success": true,
  "message": "משוב נשמר בהצלחה"
}
```

### POST /api/analyze-emails-ai
מנתח מיילים נבחרים עם AI.

**Request**:
```json
{
  "emails": [
    {
      "id": "email_123",
      "subject": "נושא המייל",
      "sender": "שולח"
    }
  ]
}
```

**Response**:
```json
{
  "success": true,
  "message": "ניתוח AI הושלם",
  "updated_count": 5,
  "updated_emails": [
    {
      "id": "email_123",
      "ai_importance_score": 0.92,
      "ai_analyzed": true,
      "ai_analysis_date": "2025-09-29T10:35:00Z"
    }
  ]
}
```

## 📅 API פגישות

### GET /api/meetings
מחזיר את כל הפגישות מהזיכרון.

**Response**:
```json
[
  {
    "id": "meeting_456",
    "subject": "נושא הפגישה",
    "organizer": "מארגן",
    "organizer_email": "organizer@example.com",
    "start_time": "2025-09-30T14:00:00Z",
    "end_time": "2025-09-30T15:00:00Z",
    "location": "חדר ישיבות A",
    "attendees": ["participant1@example.com", "participant2@example.com"],
    "body": "תיאור הפגישה...",
    "importance_score": 0.75,
    "ai_analyzed": false,
    "priority": "medium"
  }
]
```

### POST /api/meetings/<meeting_id>/priority
מעדכן עדיפות פגישה.

**Request**:
```json
{
  "priority": "high"
}
```

**Response**:
```json
{
  "success": true,
  "message": "עדיפות עודכנה בהצלחה"
}
```

**Priority Values**:
- `critical` - קריטי
- `high` - חשוב
- `medium` - בינוני
- `low` - נמוך

### GET /api/meetings/stats
מחזיר סטטיסטיקות פגישות.

**Response**:
```json
{
  "total_meetings": 25,
  "critical_meetings": 3,
  "high_meetings": 6,
  "medium_meetings": 10,
  "low_meetings": 6,
  "today_meetings": 5,
  "week_meetings": 12
}
```

## 🤖 API AI

### GET /api/ai-status
מחזיר מצב מערכת ה-AI.

**Response**:
```json
{
  "ai_available": true,
  "use_ai": true,
  "api_key_configured": true,
  "last_check": "2025-09-29T10:30:00Z",
  "quota_remaining": 95
}
```

### POST /api/analyze-meetings-ai
מנתח פגישות נבחרות עם AI.

**Request**:
```json
{
  "meetings": [
    {
      "id": "meeting_456",
      "subject": "נושא הפגישה",
      "organizer": "מארגן"
    }
  ]
}
```

**Response**:
```json
{
  "success": true,
  "message": "ניתוח AI הושלם",
  "updated_count": 3,
  "updated_meetings": [
    {
      "id": "meeting_456",
      "ai_importance_score": 0.88,
      "ai_analyzed": true,
      "ai_analysis_date": "2025-09-29T10:35:00Z"
    }
  ]
}
```

## 🔧 API מערכת

### GET /api/test-outlook
בודק חיבור ל-Outlook.

**Response**:
```json
{
  "outlook_connected": true,
  "emails_count": 150,
  "meetings_count": 25,
  "last_check": "2025-09-29T10:30:00Z"
}
```

### GET /api/server-id
מחזיר מזהה ייחודי לשרת.

**Response**:
```json
{
  "server_id": "20250930_103000",
  "uptime": "2 hours 15 minutes",
  "version": "1.0.0"
}
```

### POST /api/restart-server
מפעיל מחדש את השרת.

**Response**:
```json
{
  "success": true,
  "message": "שרת הופעל מחדש",
  "restart_time": "2025-09-29T10:35:00Z"
}
```

## 🖥️ API קונסול

### GET /api/console-logs
מחזיר את הלוגים מהקונסול.

**Response**:
```json
{
  "logs": [
    "[10:30:00] INFO: Server started",
    "[10:30:15] SUCCESS: Outlook connected",
    "[10:30:30] INFO: AI analysis completed"
  ],
  "count": 50
}
```

### POST /api/clear-console
מנקה את הלוגים מהקונסול.

**Response**:
```json
{
  "success": true,
  "message": "קונסול נוקה בהצלחה"
}
```

### POST /api/console-reset
מאפס את הקונסול ומטען מחדש.

**Response**:
```json
{
  "success": true,
  "message": "קונסול אופס בהצלחה"
}
```

## 📦 API גיבויים

### POST /api/create-backup
יוצר גיבוי של הפרויקט.

**Request**:
```json
{
  "version_description": "גרסה יציבה"
}
```

**Response**:
```json
{
  "success": true,
  "message": "גיבוי נוצר בהצלחה",
  "backup_path": "C:\\Users\\user\\Downloads\\outlook_manager_backup_20250930.zip",
  "file_size": "15.2 MB"
}
```

### POST /api/create-cursor-prompts
יוצר קבצי פרומפטים ל-Cursor.

**Response**:
```json
{
  "success": true,
  "message": "פרומפטים נוצרו בהצלחה",
  "folder_path": "C:\\Users\\user\\outlook_email_manager\\Cursor_Prompts",
  "files_created": ["01_Main_Project_Prompt.txt", "02_Flask_Application.txt"]
}
```

### POST /api/create-documentation
יוצר/מרענן קבצי תיעוד MD.

**Response**:
```json
{
  "success": true,
  "message": "תיעוד נוצר בהצלחה",
  "folder_path": "C:\\Users\\user\\outlook_email_manager\\docs",
  "files_created": ["README.md", "INSTALLATION.md", "API_DOCUMENTATION.md"]
}
```

## 🔒 אבטחה

### Rate Limiting
- מקסימום 100 בקשות לדקה לכל IP
- מקסימום 10 בקשות AI לדקה

### Authentication
- כרגע אין אימות (פיתוח מקומי)
- בעתיד: JWT tokens או API keys

### CORS
- מותר מ-`localhost:5000` בלבד
- בעתיד: הגדרה גמישה יותר

## 📊 סטטוס קודים

| קוד | משמעות |
|-----|---------|
| 200 | הצלחה |
| 400 | בקשה שגויה |
| 404 | לא נמצא |
| 500 | שגיאת שרת |

## 🐛 טיפול בשגיאות

### שגיאות נפוצות
```json
{
  "success": false,
  "error": "outlook_not_connected",
  "message": "Outlook לא מחובר",
  "details": "נסה לפתוח את Outlook ולהפעיל מחדש"
}
```

### שגיאות AI
```json
{
  "success": false,
  "error": "ai_quota_exceeded",
  "message": "חרגת ממכסת ה-API",
  "details": "נסה שוב מאוחר יותר"
}
```

## 📈 ביצועים

### זמני תגובה ממוצעים
- GET /api/emails: 200ms
- POST /api/analyze-emails-ai: 2-5s
- GET /api/meetings: 150ms
- POST /api/refresh-data: 1-3s

### הגבלות
- ללא הגבלת מיילים לטעינה (יטען את כל המיילים)
- ללא הגבלת פגישות לטעינה (יטען את כל הפגישות)
- מקסימום 10 מיילים לניתוח AI בו-זמנית

---

**תיעוד זה נוצר אוטומטית על ידי המערכת** 📚
"""
        
        api_file = os.path.join(docs_folder, "API_DOCUMENTATION.md")
        with open(api_file, 'w', encoding='utf-8') as f:
            f.write(api_content)
        files_created.append("API_DOCUMENTATION.md")
        
        ui_block_end(block_id, f"נוצרו {len(files_created)} קבצי תיעוד", True)
        return jsonify({
            'success': True,
            'message': 'קבצי תיעוד נוצרו/עודכנו בהצלחה',
            'folder_path': docs_folder,
            'files_created': files_created
        })
        
    except Exception as e:
        error_msg = f'שגיאה ביצירת קבצי תיעוד: {str(e)}'
        try:
            ui_block_end(block_id, error_msg, False)
        except Exception:
            pass
        return jsonify({
            'success': False,
            'message': error_msg
        }), 500

@app.route('/api/sync-outlook', methods=['POST'])
def sync_outlook():
    """API לסנכרון ידני עם Outlook"""
    try:
        from auto_sync_outlook import AutoSyncManager
        
        block_id = ui_block_start("🔄 סנכרון Outlook")
        ui_block_add(block_id, "מתחיל סנכרון עם Outlook...", "INFO")
        
        manager = AutoSyncManager()
        success = manager.sync_all()
        
        if success:
            ui_block_end(block_id, "סנכרון Outlook הושלם בהצלחה!", True)
            return jsonify({
                'success': True,
                'message': 'סנכרון הושלם בהצלחה'
            })
        else:
            ui_block_end(block_id, "שגיאה בסנכרון Outlook", False)
            return jsonify({
                'success': False,
                'message': 'שגיאה בסנכרון'
            }), 500
            
    except Exception as e:
        error_msg = f'שגיאה בסנכרון Outlook: {str(e)}'
        try:
            ui_block_end(block_id, error_msg, False)
        except:
            pass
        return jsonify({
            'success': False,
            'message': error_msg
        }), 500

if __name__ == '__main__':
    # כל ההשתקות כבר נעשו בראש הקובץ
    import psutil
    
    current_pid = os.getpid()
    current_script = os.path.abspath(__file__)
    killed_count = 0
    
    print("=" * 60)
    print("🔍 Checking for previous server instances...")
    
    try:
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                # בדיקה אם זה תהליך Python שמריץ את אותו הסקריפט
                if proc.info['pid'] != current_pid and proc.info['name'] and 'python' in proc.info['name'].lower():
                    cmdline = proc.info.get('cmdline', [])
                    if cmdline and any('app_with_ai.py' in str(arg) for arg in cmdline):
                        print(f"🔪 Killing old process: PID {proc.info['pid']}")
                        proc.kill()
                        killed_count += 1
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    except Exception as e:
        print(f"⚠️ Error searching for old processes: {e}")
    
    if killed_count > 0:
        print(f"✅ Killed {killed_count} old server instance(s)")
        import time
        time.sleep(1)  # המתנה שניה להבטחת סגירת התהליכים
    else:
        print("✅ No previous instances found")
    
    # ניקוי כל הלוגים הקודמים כשהשרת מתחיל מחדש
    clear_all_console_logs()

    # לא מבצעים connect_to_outlook כאן כדי למנוע בלוקים כפולים בעלייה
    # נבצע טעינת נתונים ראשונית רק אם לא נטענו מיילים ועדיין אין טעינה פעילה
    import threading
    try:
        if not (cached_data.get('emails')) and not cached_data.get('is_loading'):
            threading.Thread(target=load_initial_data, daemon=True).start()
    except Exception:
        threading.Thread(target=load_initial_data, daemon=True).start()

    # Port ניתן לקינפוג דרך משתני סביבה APP_PORT/PORT (ברירת מחדל 5000)
    try:
        chosen_port = int(os.environ.get('APP_PORT') or os.environ.get('PORT') or '5000')
    except Exception:
        chosen_port = 5000
    
    print(f"🚀 Starting Flask server on http://127.0.0.1:{chosen_port}")
    print("=" * 60)
    print("✨ Server is running! Press Ctrl+C to stop")
    print("=" * 60)
    print()  # שורה ריקה
    
    # השתקת CLI של werkzeug
    cli = sys.modules.get('flask.cli')
    if cli is not None:
        cli.show_server_banner = lambda *args, **kwargs: None
    
    # סינון stdout להסתרת הודעות Flask
    class QuietStdout:
        def __init__(self, stdout):
            self.stdout = stdout
            
        def write(self, text):
            # מסנן הודעות מיותרות
            if any(x in text for x in ['Tip:', 'Serving Flask', 'Debug mode:', 'WARNING: This is']):
                return
            self.stdout.write(text)
            
        def flush(self):
            self.stdout.flush()
    
    original_stdout = sys.stdout
    sys.stdout = QuietStdout(original_stdout)
    
    try:
        # הרצת השרת
        app.run(debug=False, host='127.0.0.1', port=chosen_port, use_reloader=False, threaded=True)
    except KeyboardInterrupt:
        sys.stdout = original_stdout
        print("\n" + "=" * 60)
        print("🛑 Server stopped")
        print("=" * 60)
    finally:
        # החזרת stdout ו-stderr למצב רגיל
        sys.stdout = original_stdout
        sys.stderr = _original_stderr
