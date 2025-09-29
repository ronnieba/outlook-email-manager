"""
Outlook Email Manager - With AI Integration
מערכת ניהול מיילים חכמה עם AI + Outlook + Gemini
"""
from flask import Flask, render_template, request, jsonify
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

app = Flask(__name__)

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
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)  # Inbox
            self.outlook_connected = True
            print("✅ חיבור ל-Outlook הצליח!")
            return True
        except Exception as e:
            print(f"❌ שגיאה בחיבור ל-Outlook: {e}")
            self.outlook_connected = False
            return False
    
    def get_emails_from_outlook(self, limit=20):
        """קבלת מיילים אמיתיים מ-Outlook"""
        try:
            if not self.outlook_connected:
                if not self.connect_to_outlook():
                    return []
            
            # קבלת המיילים ישירות
            messages = self.inbox.Items
            emails = []
            
            print(f"📧 נמצאו {messages.Count} מיילים ב-Outlook")
            
            # מיון לפי תאריך - חדשים קודם
            messages.Sort("[ReceivedTime]", True)
            
            for i in range(min(limit, messages.Count)):
                try:
                    message = messages[i + 1]  # Outlook מתחיל מ-1
                    
                    if message is None:
                        continue
                    
                    email_data = {
                        'id': i + 1,
                        'subject': str(message.Subject) if message.Subject else "ללא נושא",
                        'sender': str(message.SenderName) if message.SenderName else "שולח לא ידוע",
                        'sender_email': str(message.SenderEmailAddress) if message.SenderEmailAddress else "",
                        'received_time': str(message.ReceivedTime),
                        'body_preview': str(message.Body)[:200] + "..." if len(str(message.Body)) > 200 else str(message.Body),
                        'is_read': not message.UnRead
                    }
                    
                    # ניתוח AI + למידה
                    if self.use_ai and self.ai_analyzer.is_ai_available():
                        ai_importance = self.ai_analyzer.analyze_email_importance(email_data)
                        ai_category = self.ai_analyzer.categorize_email(email_data)
                        
                        # שילוב עם למידה מותאמת אישית
                        if self.use_learning:
                            learned_importance = self.profile_manager.get_personalized_importance_score(email_data)
                            learned_category = self.profile_manager.get_personalized_category(email_data)
                            
                            # ממוצע משוקלל בין AI ולמידה
                            email_data['importance_score'] = (ai_importance * 0.7 + learned_importance * 0.3)
                            email_data['category'] = learned_category if learned_category != 'work' else ai_category
                        else:
                            email_data['importance_score'] = ai_importance
                            email_data['category'] = ai_category
                        
                        email_data['summary'] = self.ai_analyzer.summarize_email(email_data)
                        email_data['action_items'] = self.ai_analyzer.extract_action_items(email_data)
                    else:
                        # ניתוח בסיסי + למידה
                        if self.use_learning:
                            email_data['importance_score'] = self.profile_manager.get_personalized_importance_score(email_data)
                            email_data['category'] = self.profile_manager.get_personalized_category(email_data)
                        else:
                            email_data['importance_score'] = self.calculate_basic_importance(email_data)
                            email_data['category'] = 'work'
                        
                        email_data['summary'] = f"מייל מ-{email_data['sender']}: {email_data['subject']}"
                        email_data['action_items'] = []
                    
                    emails.append(email_data)
                    print(f"✅ מייל {i+1}: {email_data['subject'][:30]}... (חשיבות: {email_data['importance_score']:.2f})")
                    
                except Exception as e:
                    print(f"❌ שגיאה במייל {i+1}: {e}")
                    continue
            
            print(f"📧 הוחזרו {len(emails)} מיילים אמיתיים")
            return emails
            
        except Exception as e:
            print(f"❌ שגיאה בקבלת מיילים מ-Outlook: {e}")
            self.outlook_connected = False
            return []
    
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
    
    def get_emails(self, limit=50):
        """קבלת מיילים מ-Outlook"""
        try:
            emails = self.get_emails_from_outlook(limit)
            if emails:
                print(f"📧 הוחזרו {len(emails)} מיילים מ-Outlook")
                return emails
            else:
                print("📧 Outlook לא עובד - מנסה להתחבר מחדש...")
                if self.connect_to_outlook():
                    emails = self.get_emails_from_outlook(limit)
                    if emails:
                        print(f"📧 חיבור מחדש הצליח - הוחזרו {len(emails)} מיילים")
                        return emails
                print("📧 Outlook לא זמין - משתמש בנתונים דמה")
                return self.get_sample_emails()[:limit]
        except Exception as e:
            print(f"❌ שגיאה כללית בקבלת מיילים: {e}")
            return self.get_sample_emails()[:limit]
    
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

# יצירת מופע של מנהל המיילים
email_manager = EmailManager()

@app.route('/')
def index():
    """דף הבית"""
    return render_template('index.html')

@app.route('/api/emails')
def get_emails():
    """API לקבלת מיילים"""
    print("📧 מקבל בקשת מיילים...")
    emails = email_manager.get_emails()
    print(f"📧 מחזיר {len(emails)} מיילים")
    return jsonify(emails)

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
    """API לקבלת מיילים חשובים"""
    print("⭐ מקבל בקשת מיילים חשובים...")
    emails = email_manager.get_emails()
    # מיון לפי ציון חשיבות
    important_emails = sorted(emails, key=lambda x: x['importance_score'], reverse=True)
    print(f"⭐ מחזיר {len(important_emails[:10])} מיילים חשובים")
    return jsonify(important_emails[:10])  # 10 המיילים החשובים ביותר

@app.route('/api/stats')
def get_stats():
    """API לקבלת סטטיסטיקות"""
    emails = email_manager.get_emails()
    total_emails = len(emails)
    important_emails = len([e for e in emails if e['importance_score'] >= 0.7])
    unread_emails = len([e for e in emails if not e['is_read']])
    
    return jsonify({
        'total_emails': total_emails,
        'important_emails': important_emails,
        'unread_emails': unread_emails
    })

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
    return jsonify({
        'ai_available': email_manager.ai_analyzer.is_ai_available(),
        'use_ai': email_manager.use_ai,
        'message': 'AI זמין' if email_manager.ai_analyzer.is_ai_available() else 'AI לא זמין - נדרש API Key'
    })

@app.route('/api/toggle-ai')
def toggle_ai():
    """API למעבר בין AI לניתוח בסיסי"""
    email_manager.use_ai = not email_manager.use_ai
    return jsonify({
        'use_ai': email_manager.use_ai,
        'message': 'AI מופעל' if email_manager.use_ai else 'AI מושבת'
    })

@app.route('/api/test-outlook')
def test_outlook():
    """API לבדיקת חיבור ל-Outlook"""
    try:
        if email_manager.connect_to_outlook():
            return jsonify({
                'success': True,
                'message': 'חיבור ל-Outlook הצליח!'
            })
        else:
            return jsonify({
                'success': False,
                'message': 'לא ניתן להתחבר ל-Outlook'
            })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'שגיאה: {str(e)}'
        })

# Removed duplicate function - see line 692 for the actual implementation
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

@app.route('/learning-management')
def learning_management():
    """דף ניהול למידה חכמה"""
    return render_template('learning_management.html')

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

if __name__ == '__main__':
    print("🚀 מפעיל את Outlook Email Manager עם AI...")
    print("📧 מנסה להתחבר ל-Outlook...")
    
    if email_manager.connect_to_outlook():
        print("✅ חיבור ל-Outlook הצליח!")
    else:
        print("⚠️ לא ניתן להתחבר ל-Outlook - משתמש בנתונים דמה")
    
    if email_manager.ai_analyzer.is_ai_available():
        print("🤖 AI (Gemini) זמין!")
    else:
        print("⚠️ AI לא זמין - נדרש API Key")
    
    print("🌐 מפעיל שרת web על http://localhost:5000")
    
    app.run(debug=True, host='127.0.0.1', port=5000)








