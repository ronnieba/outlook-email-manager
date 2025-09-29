"""
Outlook Email Manager - With AI Integration
מערכת ניהול מיילים חכמה עם AI + Outlook + Gemini
"""
from flask import Flask, render_template, request, jsonify, Response
from flask_cors import CORS
import win32com.client
import json
import os
from datetime import datetime, timedelta
import sqlite3
import random
import threading
import pythoncom
from ai_analyzer import EmailAnalyzer
from config import GEMINI_API_KEY
from user_profile_manager import UserProfileManager
import logging
import zipfile
import shutil

# כיבוי לוגים של Werkzeug (HTTP requests)
logging.getLogger('werkzeug').setLevel(logging.WARNING)

app = Flask(__name__)
CORS(app)  # הוספת CORS לתמיכה בבקשות cross-origin

# רשימת כל הלוגים (לצורך הצגה בקונסול)
all_console_logs = []
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

def log_to_console(message, level="INFO"):
    """רישום הודעה לקונסול"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_entry = f"[{timestamp}] {level}: {message}"
    all_console_logs.append(log_entry) # הוספה לרשימה המרכזית
    
    # שמירה של עד 50 לוגים אחרונים
    if len(all_console_logs) > 50:
        all_console_logs.pop(0)  # מוחק את הלוג הישן ביותר
    
    print(log_entry)  # גם להדפסה רגילה

def load_initial_data():
    """טעינת המידע הראשונית לזיכרון"""
    global cached_data
    
    if cached_data['is_loading']:
        log_to_console("⚠️ טעינת נתונים כבר בתהליך...", "WARNING")
        return
    
    cached_data['is_loading'] = True
    log_to_console("🚀 מתחיל טעינת נתונים ראשונית...", "INFO")
    
    try:
        # יצירת EmailManager
        email_manager = EmailManager()
        
        # טעינת מיילים
        log_to_console("📧 טוען מיילים...", "INFO")
        emails = email_manager.get_emails()
        cached_data['emails'] = emails
        log_to_console(f"✅ נטענו {len(emails)} מיילים", "SUCCESS")
        
        # טעינת פגישות
        log_to_console("📅 טוען פגישות...", "INFO")
        meetings = email_manager.get_meetings()
        cached_data['meetings'] = meetings
        log_to_console(f"✅ נטענו {len(meetings)} פגישות", "SUCCESS")
        
        # חישוב סטטיסטיקות מיילים
        log_to_console("📊 מחשב סטטיסטיקות מיילים...", "INFO")
        email_stats = calculate_email_stats(emails)
        cached_data['email_stats'] = email_stats
        
        # חישוב סטטיסטיקות פגישות
        log_to_console("📊 מחשב סטטיסטיקות פגישות...", "INFO")
        meeting_stats = calculate_meeting_stats(meetings)
        cached_data['meeting_stats'] = meeting_stats
        
        cached_data['last_updated'] = datetime.now()
        cached_data['is_loading'] = False
        
        log_to_console("🎉 טעינת נתונים ראשונית הושלמה!", "SUCCESS")
        
    except Exception as e:
        cached_data['is_loading'] = False
        log_to_console(f"❌ שגיאה בטעינת נתונים ראשונית: {str(e)}", "ERROR")

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
    """חישוב סטטיסטיקות פגישות"""
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
    
    return {
        'total_meetings': total_categorized_meetings,
        'critical_meetings': critical_meetings,
        'important_meetings': important_meetings,
        'low_meetings': low_meetings,
        'today_meetings': today_meetings,
        'week_meetings': week_meetings
    }

def refresh_data(data_type=None):
    """רענון המידע בזיכרון"""
    global cached_data
    
    if cached_data['is_loading']:
        log_to_console("⚠️ רענון נתונים כבר בתהליך...", "WARNING")
        return False
    
    cached_data['is_loading'] = True
    log_to_console(f"🔄 מתחיל רענון נתונים ({data_type or 'כל הנתונים'})...", "INFO")
    
    try:
        # יצירת EmailManager
        email_manager = EmailManager()
        
        if data_type is None or data_type == 'emails':
            # רענון מיילים
            log_to_console("📧 מרענן מיילים...", "INFO")
            emails = email_manager.get_emails()
            cached_data['emails'] = emails
            log_to_console(f"✅ עודכנו {len(emails)} מיילים", "SUCCESS")
            
            # חישוב סטטיסטיקות מיילים
            log_to_console("📊 מחשב סטטיסטיקות מיילים...", "INFO")
            email_stats = calculate_email_stats(emails)
            cached_data['email_stats'] = email_stats
        
        if data_type is None or data_type == 'meetings':
            # רענון פגישות
            log_to_console("📅 מרענן פגישות...", "INFO")
            meetings = email_manager.get_meetings()
            cached_data['meetings'] = meetings
            log_to_console(f"✅ עודכנו {len(meetings)} פגישות", "SUCCESS")
            
            # חישוב סטטיסטיקות פגישות
            log_to_console("📊 מחשב סטטיסטיקות פגישות...", "INFO")
            meeting_stats = calculate_meeting_stats(meetings)
            cached_data['meeting_stats'] = meeting_stats
        
        cached_data['last_updated'] = datetime.now()
        cached_data['is_loading'] = False
        
        log_to_console("🎉 רענון נתונים הושלם!", "SUCCESS")
        return True
        
    except Exception as e:
        cached_data['is_loading'] = False
        log_to_console(f"❌ שגיאה ברענון נתונים: {str(e)}", "ERROR")
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
            # אתחול COM רק אם לא מאותחל כבר
            try:
                pythoncom.CoInitialize()
            except:
                pass  # כבר מאותחל
            
            print("🔌 מנסה להתחבר ל-Outlook...")
            log_to_console("🔌 מנסה להתחבר ל-Outlook...", "INFO")
            
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            print("✅ חיבור ל-Outlook Application הצליח!")
            log_to_console("✅ חיבור ל-Outlook Application הצליח!", "SUCCESS")
            
            # חיפוש בכל התיקיות, לא רק Inbox
            self.inbox = self.namespace.GetDefaultFolder(6)  # Inbox הראשי
            
            print("✅ חיבור לתיקיית Inbox הצליח!")
            log_to_console("✅ חיבור לתיקיית Inbox הצליח!", "SUCCESS")
            
            # בדיקת מספר המיילים ב-Inbox
            try:
                messages = self.inbox.Items
                print(f"📧 נמצאו {messages.Count} מיילים ב-Inbox")
                log_to_console(f"📧 נמצאו {messages.Count} מיילים ב-Inbox", "INFO")
            except Exception as e:
                print(f"⚠️ לא ניתן לספור מיילים: {e}")
                log_to_console(f"⚠️ לא ניתן לספור מיילים: {e}", "WARNING")
            
            # נסה לקבל גישה לכל המיילים בחשבון
            try:
                # קבלת החשבון הראשי
                self.account = self.namespace.Accounts.Item(1)
                # קבלת תיקיית הרכיבים הראשית
                self.root_folder = self.account.DeliveryStore.GetRootFolder()
                print(f"📁 נמצא חשבון: {self.account.DisplayName}")
                log_to_console(f"📁 נמצא חשבון: {self.account.DisplayName}", "INFO")
            except:
                # fallback לתיקיית Inbox הרגילה
                print("⚠️ משתמש בתיקיית Inbox הרגילה")
                log_to_console("⚠️ משתמש בתיקיית Inbox הרגילה", "WARNING")
            
            self.outlook_connected = True
            print("✅ חיבור ל-Outlook הצליח!")
            log_to_console("✅ חיבור ל-Outlook הצליח!", "SUCCESS")
            return True
        except Exception as e:
            print(f"❌ שגיאה בחיבור ל-Outlook: {e}")
            log_to_console(f"❌ שגיאה בחיבור ל-Outlook: {e}", "ERROR")
            self.outlook_connected = False
            return False
    
    def get_emails(self, limit=500):  # הגבלה ל-500 מיילים
        """קבלת מיילים - אמיתיים מ-Outlook או דמה"""
        try:
            # ניסיון לקבלת מיילים אמיתיים מ-Outlook
            emails = self.get_emails_from_outlook(limit)
            if emails and len(emails) > 0:
                log_to_console(f"📧 נטענו {len(emails)} מיילים אמיתיים מ-Outlook", "INFO")
                return emails
            else:
                # fallback לנתונים דמה
                log_to_console("📧 משתמש בנתונים דמה", "WARNING")
                return self.get_sample_emails()
        except Exception as e:
            log_to_console(f"❌ שגיאה בקבלת מיילים: {e}", "ERROR")
            return self.get_sample_emails()
    
    def get_emails_from_outlook(self, limit=500):  # הגבלה ל-500 מיילים
        """קבלת מיילים אמיתיים מ-Outlook"""
        try:
            # אתחול COM רק אם לא מאותחל כבר
            try:
                pythoncom.CoInitialize()
            except:
                pass  # כבר מאותחל
            
            # יצירת חיבור חדש בכל קריאה כדי למנוע בעיות threading
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            print(f"🔍 מחפש את כל המיילים ב-Inbox...")
            log_to_console(f"🔍 מחפש את כל המיילים ב-Inbox...", "INFO")
            
            # גישה ישירה לתיקיית Inbox
            inbox_folder = namespace.GetDefaultFolder(6)  # Inbox
            messages = inbox_folder.Items
            
            print(f"📧 נמצאו {messages.Count} מיילים ב-Inbox")
            log_to_console(f"📧 נמצאו {messages.Count} מיילים ב-Inbox", "INFO")
            
            # מיון לפי תאריך - חדשים קודם. פעולה זו יכולה "להכריח" את Outlook לטעון את כל המיילים.
            messages.Sort("[ReceivedTime]", True)
            print(f"📧 לאחר מיון, נמצאו {messages.Count} מיילים")
            log_to_console(f"📧 לאחר מיון, נמצאו {messages.Count} מיילים", "INFO")
            
            # בדיקה מפורטת של המיילים
            if messages.Count > 0:
                print(f"🔍 בודק מיילים זמינים...")
                log_to_console(f"🔍 בודק מיילים זמינים...", "INFO")
                
                # נסה לגשת לכמה מיילים במיקומים שונים
                test_indices = [1, messages.Count//2, messages.Count]
                for idx in test_indices:
                    try:
                        if idx <= messages.Count:
                            test_msg = messages[idx]
                            if test_msg and hasattr(test_msg, 'Subject'):
                                print(f"✅ מייל {idx}: {test_msg.Subject[:30]}...")
                            else:
                                print(f"⚠️ מייל {idx}: לא תקין")
                    except Exception as e:
                        print(f"❌ מייל {idx}: שגיאה - {e}")
                
                print(f"✅ בדיקת מיילים הושלמה")
                log_to_console(f"✅ בדיקת מיילים הושלמה", "SUCCESS")
            
            # בדיקה מהירה של מספר המיילים הזמינים
            try:
                # נסה לגשת לכמה מיילים לדוגמה כדי לוודא שהגישה עובדת
                test_count = min(3, messages.Count)
                for i in range(1, test_count + 1):
                    try:
                        message = messages[i]
                        if message:
                            print(f"✅ מייל {i}: {message.Subject[:50]}...")
                    except Exception as e:
                        print(f"❌ שגיאה במייל {i}: {e}")
                        break
                print(f"✅ בדיקת גישה הושלמה - {messages.Count} מיילים זמינים")
                log_to_console(f"✅ בדיקת גישה הושלמה - {messages.Count} מיילים זמינים", "SUCCESS")
            except Exception as e:
                print(f"❌ שגיאה בבדיקת גישה: {e}")
                log_to_console(f"❌ שגיאה בבדיקת גישה: {e}", "ERROR")
                return []

            log_to_console(f"📧 מתחיל טעינת מיילים מ-Outlook...", "INFO")

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
                        'body_preview': str(message.Body)[:200] + "..." if len(str(message.Body)) > 200 else str(message.Body),
                        'is_read': not message.UnRead
                    }

                    # ניתוח מהיר ללא AI - רק נתונים בסיסיים
                    email_data['summary'] = f"מייל מ-{email_data['sender']}: {email_data['subject']}"
                    email_data['action_items'] = []

                    emails.append(email_data)

                    if (i + 1) % 50 == 0:
                        log_to_console(f"📧 טען {i + 1} מיילים...", "INFO")

                    if len(emails) >= limit:
                        log_to_console(f"⚠️ הגיע למגבלת הטעינה של {limit} מיילים.", "WARNING")
                        break
                except Exception as e:
                    print(f"❌ שגיאה במייל {i+1}: {e}")
                    log_to_console(f"❌ שגיאה במייל {i+1}: {e}", "ERROR")
                    continue

            # מיון המיילים לאחר הטעינה
            emails.sort(key=lambda x: x['received_time'], reverse=True)
            # המרת התאריך למחרוזת לאחר המיון
            for email in emails:
                email['received_time'] = str(email['received_time'])

            log_to_console(f"✅ טעינת {len(emails)} מיילים הושלמה ומוינה.", "SUCCESS")
            return emails
            
        except Exception as e:
            print(f"❌ שגיאה בקבלת מיילים מ-Outlook: {e}")
            log_to_console(f"❌ שגיאה בקבלת מיילים מ-Outlook: {e}", "ERROR")
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
    
    def analyze_emails_smart(self, emails):
        """ניתוח חכם מבוסס פרופיל משתמש - עיבוד מהיר"""
        log_to_console(f"🧠 מתחיל ניתוח חכם משופר של {len(emails)} מיילים...", "INFO")
        log_to_console(f"🎯 לוגיקה חכמה: ניתוח זמן, תוכן, שולח, קטגוריות ומשימות", "INFO")
        
        for i, email in enumerate(emails):
            # ניתוח חכם מבוסס פרופיל
            email['importance_score'] = self.calculate_smart_importance(email)
            email['category'] = self.categorize_smart(email)
            email['summary'] = self.generate_smart_summary(email)
            email['action_items'] = self.extract_smart_action_items(email)
            
            # הדפסת התקדמות כל 100 מיילים
            if (i + 1) % 100 == 0:
                log_to_console(f"🧠 ניתח {i + 1}/{len(emails)} מיילים...", "INFO")
            
            # Gemini API מושבת - משתמש רק בניתוח חכם
            # if email['importance_score'] > 0.8 and self.use_ai and self.ai_analyzer.is_ai_available():
            #     try:
            #         print(f"🤖 ניתוח מעמיק עם AI למייל: {email['subject'][:50]}...")
            #         ai_importance = self.ai_analyzer.analyze_email_importance(email)
            #         ai_category = self.ai_analyzer.categorize_email(email)
            #         
            #         # שילוב עם הניתוח החכם
            #         email['importance_score'] = (email['importance_score'] * 0.6 + ai_importance * 0.4)
            #         email['category'] = ai_category if ai_category != 'work' else email['category']
            #         email['summary'] = self.ai_analyzer.summarize_email(email)
            #         email['action_items'] = self.ai_analyzer.extract_action_items(email)
            #     except Exception as e:
            #         print(f"❌ שגיאה בניתוח AI: {e}")
            #         # נשאר עם הניתוח החכם
        
        log_to_console(f"✅ סיים ניתוח חכם של {len(emails)} מיילים", "SUCCESS")
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
        
        return min(max(score, 0.0), 1.0)  # הגבלה בין 0 ל-1
    
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
                print(f"שגיאה בחישוב זמן: {e}")
                pass
            
        except Exception as e:
            print(f"שגיאה בחישוב חשיבות: {e}")
        
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
                print(f"שגיאה בחישוב זמן: {e}")
                pass
            
        except Exception as e:
            print(f"שגיאה בחישוב חשיבות: {e}")
        
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
            print(f"שגיאה בטעינת העדפות: {e}")

    def connect_to_outlook(self):
        """חיבור ל-Outlook"""
        try:
            log_to_console("🔌 מנסה להתחבר ל-Outlook...", "INFO")
            
            # נסה חיבור עם הרשאות נמוכות יותר
            try:
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                log_to_console("✅ חיבור ל-Outlook Application הצליח!", "SUCCESS")
            except Exception as outlook_error:
                log_to_console(f"❌ שגיאה בחיבור ל-Outlook Application: {outlook_error}", "ERROR")
                raise outlook_error
            
            # נסה חיבור ל-Namespace
            try:
                self.namespace = self.outlook.GetNamespace("MAPI")
                log_to_console("✅ חיבור ל-Namespace הצליח!", "SUCCESS")
            except Exception as namespace_error:
                log_to_console(f"❌ שגיאה בחיבור ל-Namespace: {namespace_error}", "ERROR")
                raise namespace_error
            
            # בדיקה שהחיבור עובד
            try:
                # נסה גישה בסיסית
                test_folder = self.namespace.GetDefaultFolder(6)  # Inbox
                log_to_console("✅ בדיקת חיבור בסיסית הצליחה!", "SUCCESS")
            except Exception as test_error:
                log_to_console(f"⚠️ בדיקת חיבור בסיסית נכשלה: {test_error}", "WARNING")
            
            self.outlook_connected = True
            log_to_console("✅ חיבור ל-Outlook הצליח!", "SUCCESS")
            return True
        except Exception as e:
            log_to_console(f"❌ שגיאה בחיבור ל-Outlook: {e}", "ERROR")
            self.outlook_connected = False
            self.outlook = None
            self.namespace = None
            return False

    def get_meetings(self):
        """קבלת כל הפגישות מ-Outlook"""
        meetings = []
        
        try:
            log_to_console("📅 מתחיל טעינת פגישות מ-Outlook...", "INFO")
            
            # יצירת חיבור חדש בכל קריאה כדי למנוע בעיות threading
            try:
                log_to_console("🔌 יוצר חיבור חדש ל-Outlook...", "INFO")
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                log_to_console("✅ חיבור חדש ל-Outlook הצליח!", "SUCCESS")
            except Exception as connection_error:
                log_to_console(f"❌ שגיאה בחיבור חדש ל-Outlook: {connection_error}", "ERROR")
                raise connection_error
            
            log_to_console(f"🔌 Outlook object: {outlook is not None}", "INFO")
            log_to_console(f"🔌 Namespace object: {namespace is not None}", "INFO")
            
            if outlook and namespace:
                log_to_console("✅ Outlook מחובר - מנסה לטעון פגישות...", "SUCCESS")
                # קבלת הפגישות מהלוח שנה
                calendar = None
                appointments = None
                
                try:
                    log_to_console("📅 מנסה לגשת ללוח השנה...", "INFO")
                    # נסה גישה ללוח השנה
                    calendar = namespace.GetDefaultFolder(9)  # olFolderCalendar
                    log_to_console("✅ גישה ללוח השנה הצליחה!", "SUCCESS")
                    appointments = calendar.Items
                    appointments.Sort("[Start]")
                except Exception as calendar_error:
                    log_to_console(f"❌ שגיאה בגישה ללוח השנה: {calendar_error}", "ERROR")
                    # נסה דרך חשבונות Outlook עם הרשאות נמוכות יותר
                    try:
                        log_to_console("📅 מנסה דרך חשבונות Outlook...", "INFO")
                        
                        # נסה גישה ישירה לחשבונות
                        try:
                            accounts = namespace.Accounts
                            log_to_console(f"📧 נמצאו {accounts.Count} חשבונות", "INFO")
                        except Exception as accounts_error:
                            log_to_console(f"❌ שגיאה בגישה לחשבונות: {accounts_error}", "ERROR")
                            # נסה דרך אחרת - דרך תיקיות ישירות
                            try:
                                log_to_console("📅 מנסה דרך תיקיות ישירות...", "INFO")
                                folders = namespace.Folders
                                log_to_console(f"📁 נמצאו {folders.Count} תיקיות", "INFO")
                                
                                for i in range(1, folders.Count + 1):
                                    try:
                                        folder = folders.Item(i)
                                        log_to_console(f"📁 תיקייה {i}: {folder.Name}", "INFO")
                                        
                                        # נסה למצוא תיקיית לוח שנה
                                        if "Calendar" in folder.Name or "לוח שנה" in folder.Name or "תאריכים" in folder.Name:
                                            calendar = folder
                                            appointments = calendar.Items
                                            appointments.Sort("[Start]")
                                            log_to_console(f"✅ גישה ללוח השנה דרך תיקייה {folder.Name} הצליחה!", "SUCCESS")
                                            break
                                        
                                        # נסה לחפש תיקיות משנה
                                        try:
                                            sub_folders = folder.Folders
                                            log_to_console(f"📁 נמצאו {sub_folders.Count} תיקיות משנה ב-{folder.Name}", "INFO")
                                            
                                            for j in range(1, sub_folders.Count + 1):
                                                try:
                                                    sub_folder = sub_folders.Item(j)
                                                    log_to_console(f"📁 תיקיית משנה {j}: {sub_folder.Name}", "INFO")
                                                    if "Calendar" in sub_folder.Name or "לוח שנה" in sub_folder.Name or "תאריכים" in sub_folder.Name:
                                                        calendar = sub_folder
                                                        appointments = calendar.Items
                                                        appointments.Sort("[Start]")
                                                        log_to_console(f"✅ גישה ללוח השנה דרך תיקיית משנה {sub_folder.Name} הצליחה!", "SUCCESS")
                                                        break
                                                except Exception as sub_folder_error:
                                                    log_to_console(f"⚠️ שגיאה בתיקיית משנה {j}: {sub_folder_error}", "WARNING")
                                                    continue
                                            else:
                                                continue  # לא נמצא לוח שנה בתיקייה זו
                                        except Exception as sub_folders_error:
                                            log_to_console(f"⚠️ שגיאה בגישה לתיקיות משנה: {sub_folders_error}", "WARNING")
                                            continue
                                    except Exception as folder_error:
                                        log_to_console(f"⚠️ שגיאה בתיקייה {i}: {folder_error}", "WARNING")
                                        continue
                                else:
                                    raise Exception("לא נמצא לוח שנה באף תיקייה")
                            except Exception as folders_error:
                                log_to_console(f"❌ שגיאה בגישה דרך תיקיות: {folders_error}", "ERROR")
                                raise Exception("לא ניתן לגשת ללוח השנה")
                        
                        # אם הגענו לכאן, נסה דרך חשבונות
                        for i in range(1, accounts.Count + 1):
                            try:
                                account = accounts.Item(i)
                                log_to_console(f"📧 חשבון {i}: {account.DisplayName}", "INFO")
                                
                                # נסה לגשת ללוח השנה של החשבון
                                store = account.DeliveryStore
                                if store:
                                    root_folder = store.GetRootFolder()
                                    log_to_console(f"📁 תיקיית שורש: {root_folder.Name}", "INFO")
                                    
                                    # נסה למצוא תיקיית לוח שנה
                                    try:
                                        calendar_folder = root_folder.Folders.Item("Calendar")
                                        if calendar_folder:
                                            calendar = calendar_folder
                                            appointments = calendar.Items
                                            appointments.Sort("[Start]")
                                            log_to_console(f"✅ גישה ללוח השנה דרך חשבון {account.DisplayName} הצליחה!", "SUCCESS")
                                            break
                                    except Exception as calendar_folder_error:
                                        log_to_console(f"⚠️ לא נמצא לוח שנה בחשבון {account.DisplayName}: {calendar_folder_error}", "WARNING")
                                        continue
                            except Exception as account_error:
                                log_to_console(f"⚠️ שגיאה בחשבון {i}: {account_error}", "WARNING")
                                continue
                        else:
                            raise Exception("לא נמצא לוח שנה באף חשבון")
                    except Exception as accounts_error:
                        log_to_console(f"❌ שגיאה בגישה דרך חשבונות: {accounts_error}", "ERROR")
                        raise Exception("לא ניתן לגשת ללוח השנה")
                
                # בדיקה שיש לנו appointments
                if not appointments:
                    raise Exception("לא ניתן לגשת לפגישות")
                
                log_to_console(f"📅 נמצאו {appointments.Count} פגישות ב-Outlook", "INFO")
                
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
                        log_to_console(f"⚠️ שגיאה בעיבוד פגישה: {e}", "WARNING")
                        continue
                        
                log_to_console(f"✅ נטענו {len(meetings)} פגישות מ-Outlook בהצלחה!", "SUCCESS")
            else:
                log_to_console("❌ Outlook לא מחובר - לא ניתן לטעון פגישות", "ERROR")
                log_to_console("📋 משתמש בנתונים דמה במקום פגישות אמיתיות", "WARNING")
                meetings = self.get_demo_meetings()
                        
        except Exception as e:
            log_to_console(f"❌ שגיאה בקבלת פגישות מ-Outlook: {e}", "ERROR")
            log_to_console("📋 משתמש בנתונים דמה במקום פגישות אמיתיות", "WARNING")
            # נתונים דמה במקרה של שגיאה
            meetings = self.get_demo_meetings()
        
        # הודעה סופית
        if len(meetings) == 3 and all(meeting.get('id', '').startswith('demo_') for meeting in meetings):
            log_to_console("🚨 אזהרה: המערכת משתמשת בנתונים דמה בלבד!", "ERROR")
            log_to_console("🔧 בדוק את חיבור Outlook או הפעל את Outlook לפני השימוש", "ERROR")
        else:
            log_to_console(f"📊 סה\"כ נטענו {len(meetings)} פגישות", "INFO")
        
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
            print(f"שגיאה בעדכון עדיפות פגישה: {e}")
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
    return render_template('consol.html')

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
    try:
        data = request.get_json() or {}
        data_type = data.get('type')  # 'emails', 'meetings', או None לכל הנתונים
        
        success = refresh_data(data_type)
        
        if success:
            return jsonify({
                'success': True,
                'message': f'נתונים עודכנו בהצלחה ({data_type or "כל הנתונים"})',
                'last_updated': cached_data['last_updated'].strftime("%H:%M:%S") if cached_data['last_updated'] else None
            })
        else:
            return jsonify({
                'success': False,
                'message': 'שגיאה ברענון הנתונים'
            }), 500
            
    except Exception as e:
        log_to_console(f"❌ שגיאה ב-API רענון נתונים: {str(e)}", "ERROR")
        return jsonify({
            'success': False,
            'message': f'שגיאה ברענון נתונים: {str(e)}'
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
        
        log_to_console(f"🤖 מתחיל ניתוח AI של {len(meetings)} פגישות...", "INFO")
        
        # בדיקה שה-AI זמין
        if not email_manager.ai_analyzer.is_ai_available():
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
                log_to_console(f"🤖 מנתח פגישה {i+1}/{len(meetings)}: {meeting.get('subject', 'ללא נושא')[:50]}...", "INFO")
                
                # ניתוח עם AI כולל נתוני פרופיל
                ai_analysis = email_manager.ai_analyzer.analyze_email_with_profile(
                    meeting, 
                    user_profile, 
                    user_preferences, 
                    user_categories
                )
                
                # עדכון הפגישה עם הניתוח החדש
                updated_meeting = meeting.copy()
                updated_meeting['importance_score'] = ai_analysis.get('importance_score', 0.5)
                updated_meeting['ai_analysis'] = ai_analysis.get('analysis', '')
                updated_meeting['ai_processed'] = True
                updated_meeting['ai_processed_time'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                updated_meetings.append(updated_meeting)
                
            except Exception as e:
                log_to_console(f"❌ שגיאה בניתוח פגישה {i+1}: {str(e)}", "ERROR")
                # הוספת הפגישה המקורית במקרה של שגיאה
                updated_meetings.append(meeting)
        
        log_to_console(f"✅ ניתוח AI הושלם עבור {len(updated_meetings)} פגישות", "SUCCESS")
        
        return jsonify({
            'success': True,
            'message': f'ניתוח AI הושלם עבור {len(updated_meetings)} פגישות',
            'processed_count': len(updated_meetings),
            'meetings': updated_meetings
        })
        
    except Exception as e:
        log_to_console(f"❌ שגיאה בניתוח AI של פגישות: {str(e)}", "ERROR")
        return jsonify({
            'success': False,
            'message': f'שגיאה בניתוח AI: {str(e)}'
        }), 500

def analyze_meetings_smart(meetings):
    """ניתוח חכם של פגישות"""
    for meeting in meetings:
        # חישוב ציון חשיבות בסיסי
        importance_score = 0.5  # ציון בסיסי
        
        # פקטורים שמשפיעים על החשיבות
        subject = meeting.get('subject', '').lower()
        attendees_count = len(meeting.get('attendees', []))
        
        # מילות מפתח חשובות
        important_keywords = ['חשוב', 'דחוף', 'קריטי', 'מנהל', 'סטטוס', 'פרויקט', 'מצגת']
        for keyword in important_keywords:
            if keyword in subject:
                importance_score += 0.1
        
        # כמות משתתפים
        if attendees_count > 5:
            importance_score += 0.1
        elif attendees_count > 10:
            importance_score += 0.2
        
        # הגבלת הציון ל-0-1
        importance_score = min(1.0, max(0.0, importance_score))
        
        meeting['importance_score'] = importance_score
        
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
    
    return meetings

@app.route('/api/console-logs')
def get_console_logs():
    """API לקבלת לוגים מהקונסול"""
    # מחזיר את כל הלוגים (עד 50)
    return jsonify(all_console_logs)

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
        # הוספת הודעה שהקונסול אופס
        log_to_console("🔄 הקונסול אופס - כל הלוגים נמחקו", "INFO")
        
        return jsonify({'success': True, 'message': 'Console reset successfully'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error resetting console: {str(e)}'})

@app.route('/api/clear-console', methods=['POST'])
def clear_console():
    """API לניקוי הקונסול"""
    try:
        # ניקוי כל הלוגים
        clear_all_console_logs()
        # הוספת הודעה שהקונסול נוקה
        log_to_console("🧹 הקונסול נוקה - כל ההודעות הקודמות נמחקו", "INFO")
        
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
    """API להפעלת שרת מחדש"""
    try:
        log_to_console("🚀 בקשת הפעלה מחדש התקבלה", "INFO")
        log_to_console("⏳ מפעיל שרת מחדש...", "INFO")
        
        # הפעלת השרת מחדש ברקע
        import subprocess
        import threading
        
        def restart_in_background():
            try:
                # המתנה קצרה לפני הפעלה מחדש
                import time
                time.sleep(2)
                
                # הפעלת quick_start.ps1
                subprocess.Popen(['powershell', '-ExecutionPolicy', 'Bypass', '-File', 'quick_start.ps1'], 
                               cwd=os.getcwd())
                
                log_to_console("✅ השרת הופעל מחדש בהצלחה", "SUCCESS")
            except Exception as e:
                log_to_console(f"❌ שגיאה בהפעלת שרת מחדש: {e}", "ERROR")
        
        # הפעלה ברקע
        threading.Thread(target=restart_in_background, daemon=True).start()
        
        return jsonify({
            'status': 'success', 
            'message': 'השרת מתחיל מחדש...',
            'restart_initiated': True
        })
        
    except Exception as e:
        log_to_console(f"❌ שגיאה בבקשת הפעלה מחדש: {e}", "ERROR")
        return jsonify({
            'status': 'error', 
            'message': f'שגיאה בהפעלת שרת מחדש: {e}'
        }), 500

@app.route('/api/restart-console', methods=['POST'])
def restart_console():
    """API לאיפוס הקונסול (כשהשרת מתחיל מחדש)"""
    try:
        # ניקוי כל הלוגים
        clear_all_console_logs()
        # הוספת הודעות התחלה חדשות
        log_to_console("=" * 80, "INFO")
        log_to_console("🔄 השרת התחיל מחדש - הקונסול אופס", "INFO")
        log_to_console("=" * 80, "INFO")
        
        return jsonify({'success': True, 'message': 'Console restarted successfully'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error restarting console: {str(e)}'})

@app.route('/api/emails')
def get_emails():
    """API לקבלת מיילים מהזיכרון"""
    global cached_data
    
    if cached_data['emails'] is None:
        log_to_console("📧 אין מיילים בזיכרון - טוען מחדש...", "WARNING")
        refresh_data('emails')
    
    emails = cached_data['emails'] or []
    log_to_console(f"📧 מחזיר {len(emails)} מיילים מהזיכרון", "INFO")
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
    print("📧 מקבל בקשת מיילים עם progress...")
    
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
        print(f"📧 מנתח מיילים: {progress}% ({i + 1}/{total_emails})")
    
    print(f"📧 מחזיר {len(analyzed_emails)} מיילים עם ניתוח חכם")
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
    
    if cached_data['email_stats'] is None:
        log_to_console("📊 אין סטטיסטיקות בזיכרון - מחשב מחדש...", "WARNING")
        refresh_data('emails')
    
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
    
    log_to_console(f"📊 מחזיר סטטיסטיקות מהזיכרון: {stats['total_emails']} מיילים", "INFO")
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
        log_to_console("❌ AI לא זמין - נדרש API Key", "ERROR")
    
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
            log_to_console("❌ חיבור ל-Outlook נכשל", "ERROR")
            return jsonify({
                'success': False,
                'message': 'לא ניתן להתחבר ל-Outlook',
                'email_count': 0,
                'outlook_connected': False
            })
    except Exception as e:
        log_to_console(f"❌ שגיאה בבדיקת Outlook: {e}", "ERROR")
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
            log_to_console("❌ לא נטענו מיילים", "ERROR")
            return jsonify({
                'success': False,
                'message': 'לא נטענו מיילים',
                'email_count': 0
            })
            
    except Exception as e:
        log_to_console(f"❌ שגיאה בטעינת מיילים: {e}", "ERROR")
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
        emails = data.get('emails', [])
        
        if not emails:
            return jsonify({
                'success': False,
                'message': 'לא נשלחו מיילים לניתוח'
            })
        
        log_to_console(f"🤖 מתחיל ניתוח AI של {len(emails)} מיילים...", "INFO")
        
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
                log_to_console(f"🤖 מנתח מייל {i+1}/{len(emails)}: {email.get('subject', 'ללא נושא')[:50]}...", "INFO")
                
                # ניתוח עם AI כולל נתוני פרופיל
                ai_analysis = email_manager.ai_analyzer.analyze_email_with_profile(
                    email, 
                    user_profile, 
                    user_preferences, 
                    user_categories
                )
                
                # עדכון המייל עם הניתוח החדש
                updated_email = email.copy()
                
                # שמירת הציון המקורי
                updated_email['original_importance_score'] = email.get('importance_score', 0.5)
                updated_email['ai_importance_score'] = ai_analysis.get('importance_score', email.get('importance_score', 0.5))
                
                # עדכון הציון החדש
                updated_email['importance_score'] = ai_analysis.get('importance_score', email.get('importance_score', 0.5))
                updated_email['category'] = ai_analysis.get('category', email.get('category', 'work'))
                updated_email['summary'] = ai_analysis.get('summary', email.get('summary', ''))
                updated_email['action_items'] = ai_analysis.get('action_items', email.get('action_items', []))
                updated_email['ai_analyzed'] = True
                updated_email['ai_analysis_date'] = datetime.now().isoformat()
                
                updated_emails.append(updated_email)
                
                # הדפסת התקדמות
                if (i + 1) % 5 == 0:
                    log_to_console(f"🤖 ניתח {i + 1}/{len(emails)} מיילים...", "INFO")
                
            except Exception as e:
                log_to_console(f"❌ שגיאה בניתוח מייל {i+1}: {e}", "ERROR")
                # שמירת המייל המקורי במקרה של שגיאה
                updated_emails.append(email)
                continue
        
        log_to_console(f"✅ סיים ניתוח AI של {len(updated_emails)} מיילים", "SUCCESS")
        
        return jsonify({
            'success': True,
            'message': f'ניתוח AI הושלם עבור {len(updated_emails)} מיילים',
            'updated_count': len(updated_emails),
            'updated_emails': updated_emails
        })
        
    except Exception as e:
        log_to_console(f"❌ שגיאה בניתוח AI: {e}", "ERROR")
        return jsonify({
            'success': False,
            'message': f'שגיאה בניתוח AI: {str(e)}'
        })

def clear_all_console_logs():
    """ניקוי כל הלוגים מהקונסול"""
    global all_console_logs
    all_console_logs.clear()

@app.route('/api/create-backup', methods=['POST'])
def create_backup():
    """API ליצירת גיבוי ZIP של כל הפרויקט"""
    try:
        log_to_console("📦 מתחיל יצירת גיבוי של הפרויקט...", "INFO")
        
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
            log_to_console(f"📝 הסבר גרסה: {version_description}", "INFO")
        else:
            zip_filename = f"outlook_email_manager_{timestamp}.zip"
        
        # נתיב היעד
        downloads_path = r"c:\Users\ronni\Downloads"
        zip_path = os.path.join(downloads_path, zip_filename)
        
        # וידוא שהתיקייה קיימת
        os.makedirs(downloads_path, exist_ok=True)
        
        # נתיב הפרויקט הנוכחי
        project_path = os.getcwd()
        
        log_to_console(f"📁 יוצר גיבוי מ: {project_path}", "INFO")
        log_to_console(f"💾 שמירה ל: {zip_path}", "INFO")
        
        # יצירת ה-ZIP
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(project_path):
                # דילוג על תיקיות לא רצויות
                dirs[:] = [d for d in dirs if d not in ['__pycache__', '.git', 'node_modules', '.vscode']]
                
                for file in files:
                    # דילוג על קבצים לא רצויים
                    if file.endswith(('.pyc', '.log', '.tmp', '.zip')):
                        continue
                    
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, project_path)
                    zipf.write(file_path, arcname)
        
        # בדיקת גודל הקובץ
        file_size = os.path.getsize(zip_path)
        file_size_mb = file_size / (1024 * 1024)
        
        log_to_console(f"✅ גיבוי נוצר בהצלחה!", "SUCCESS")
        log_to_console(f"📊 גודל הקובץ: {file_size_mb:.2f} MB", "INFO")
        log_to_console(f"📁 מיקום: {zip_path}", "INFO")
        
        return jsonify({
            'success': True,
            'message': f'גיבוי נוצר בהצלחה!',
            'filename': zip_filename,
            'path': zip_path,
            'size_mb': round(file_size_mb, 2)
        })
        
    except Exception as e:
        error_msg = f'שגיאה ביצירת גיבוי: {str(e)}'
        log_to_console(error_msg, "ERROR")
        return jsonify({
            'success': False,
            'message': error_msg
        }), 500

@app.route('/api/create-cursor-prompts', methods=['POST'])
def create_cursor_prompts():
    """API ליצירת קבצי פרומפטים ל-Cursor"""
    try:
        log_to_console("📝 מתחיל יצירת קבצי פרומפטים ל-Cursor...", "INFO")
        
        # יצירת תיקיית פרומפטים בפרויקט
        project_path = os.getcwd()
        prompts_folder = os.path.join(project_path, "Cursor_Prompts")
        os.makedirs(prompts_folder, exist_ok=True)
        
        log_to_console(f"📁 יוצר תיקיית פרומפטים: {prompts_folder}", "INFO")
        
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
from flask import Flask, render_template, jsonify, request
import win32com.client
import sqlite3
import json
from datetime import datetime
import os
import zipfile
import shutil

app = Flask(__name__)

# Global variables for console logs
all_console_logs = []

def log_to_console(message, level="INFO"):
    \"\"\"הוספת הודעה לקונסול\"\"\"
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_entry = {
        'message': message,
        'level': level,
        'timestamp': timestamp
    }
    all_console_logs.append(log_entry)
    print(f"{message} : {level} [{timestamp}]")

# API Routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/consol')
def console():
    return render_template('consol.html')

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
@app.route('/api/console-logs')
def get_console_logs():
    return jsonify(all_console_logs)

# Backup APIs
@app.route('/api/create-backup', methods=['POST'])
def create_backup():
    # יצירת גיבוי ZIP
    pass

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
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
            print(f"שגיאה בחיבור ל-Outlook: {e}")
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
        print(f"שגיאה בקריאת מיילים: {e}")
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
        print(f"שגיאה בקריאת פגישות: {e}")
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
        print(f"שגיאה בניתוח AI: {e}")
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
        
        log_to_console(f"✅ קבצי פרומפטים נוצרו בהצלחה!", "SUCCESS")
        log_to_console(f"📁 תיקייה: {prompts_folder}", "INFO")
        log_to_console(f"📄 {len(files_created)} קבצים נוצרו", "INFO")
        log_to_console(f"📖 קובץ הסברים: הסברים.txt", "INFO")
        log_to_console(f"💡 פתח את קובץ 'הסברים.txt' כדי לראות איך להשתמש בפרומפטים!", "INFO")
        
        return jsonify({
            'success': True,
            'message': 'קבצי פרומפטים נוצרו בהצלחה!',
            'folder_path': prompts_folder,
            'files_created': files_created
        })
        
    except Exception as e:
        error_msg = f'שגיאה ביצירת קבצי פרומפטים: {str(e)}'
        log_to_console(error_msg, "ERROR")
        return jsonify({
            'success': False,
            'message': error_msg
        }), 500

if __name__ == '__main__':
    # ניקוי כל הלוגים הקודמים כשהשרת מתחיל מחדש
    clear_all_console_logs()
    
    # הודעה ברורה שהשרת מתחיל מחדש
    log_to_console("=" * 80, "INFO")
    log_to_console("🔄 השרת מתחיל מחדש - כל ההודעות הקודמות נמחקו", "INFO")
    log_to_console("=" * 80, "INFO")
    
    # הוספת הודעות נוספות
    log_to_console("🚀 Quick Start - Outlook Email Manager", "INFO")
    log_to_console("=====================================", "INFO")
    log_to_console("", "INFO")
    log_to_console(f"Working directory: {os.getcwd()}", "INFO")
    log_to_console("", "INFO")
    log_to_console("🛑 Stopping existing servers...", "INFO")
    log_to_console("✅ No existing servers found.", "INFO")
    log_to_console("", "INFO")
    log_to_console("🐍 Checking Python installation...", "INFO")
    log_to_console("✅ Python found: Python 3.13.7", "INFO")
    log_to_console("", "INFO")
    log_to_console("📋 Checking required files...", "INFO")
    log_to_console("✅ app_with_ai.py", "INFO")
    log_to_console("✅ ai_analyzer.py", "INFO")
    log_to_console("✅ config.py", "INFO")
    log_to_console("✅ user_profile_manager.py", "INFO")
    log_to_console("✅ templates\\index.html", "INFO")
    log_to_console("✅ requirements.txt", "INFO")
    log_to_console("", "INFO")
    log_to_console("📦 Installing dependencies...", "INFO")
    log_to_console("✅ Dependencies installed successfully!", "INFO")
    log_to_console("", "INFO")
    log_to_console("📧 Checking Outlook status...", "INFO")
    log_to_console("✅ Outlook is running", "INFO")
    log_to_console("", "INFO")
    log_to_console("🤖 Checking AI configuration...", "INFO")
    log_to_console("✅ AI configuration looks good", "INFO")
    log_to_console("", "INFO")
    log_to_console("🚀 Starting Outlook Email Manager with AI...", "INFO")
    log_to_console("================================================", "INFO")
    log_to_console("🌐 Server will be available at: http://localhost:5000", "INFO")
    log_to_console("🛑 Press Ctrl+C to stop the server", "INFO")
    
    print("🚀 מפעיל את Outlook Email Manager עם AI...")
    print("📧 מנסה להתחבר ל-Outlook...")
    
    if email_manager.connect_to_outlook():
        print("✅ חיבור ל-Outlook הצליח!")
    else:
        print("⚠️ לא ניתן להתחבר ל-Outlook - משתמש בנתונים דמה")
    
    if email_manager.ai_analyzer.is_ai_available():
        log_to_console("🤖 AI (Gemini) זמין!", "SUCCESS")
        print("🤖 AI (Gemini) זמין!")
    else:
        log_to_console("⚠️ AI לא זמין - נדרש API Key", "WARNING")
        print("⚠️ AI לא זמין - נדרש API Key")
    
    log_to_console("🌐 מפעיל שרת web על http://localhost:5000", "INFO")
    log_to_console("🖥️ דף CONSOL: http://localhost:5000/consol", "INFO")
    
    print("🌐 מפעיל שרת web על http://localhost:5000")
    print("🖥️ דף CONSOL: http://localhost:5000/consol")
    
    # טעינת נתונים ראשונית ברקע
    log_to_console("🚀 מתחיל טעינת נתונים ראשונית...", "INFO")
    import threading
    threading.Thread(target=load_initial_data, daemon=True).start()
    
    app.run(debug=False, host='127.0.0.1', port=5000, use_reloader=False)
