"""
Outlook Email Manager - With Real Outlook Connection
מערכת ניהול מיילים חכמה עם AI + חיבור אמיתי ל-Outlook
"""
from flask import Flask, render_template, request, jsonify
import win32com.client
import json
import os
from datetime import datetime, timedelta
import sqlite3
import random

app = Flask(__name__)

class EmailManager:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.inbox = None
        self.user_preferences = {}
        self.db_path = "email_preferences.db"
        self.init_database()
        self.load_user_preferences()
        self.use_real_outlook = True  # דגל למעבר בין נתונים אמיתיים לדמה
    
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
        
        conn.commit()
        conn.close()
    
    def connect_to_outlook(self):
        """חיבור ל-Outlook"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)  # Inbox
            print("✅ חיבור ל-Outlook הצליח!")
            return True
        except Exception as e:
            print(f"❌ שגיאה בחיבור ל-Outlook: {e}")
            self.use_real_outlook = False
            return False
    
    def get_emails_from_outlook(self, limit=20):
        """קבלת מיילים אמיתיים מ-Outlook"""
        try:
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
                        'importance_score': self.calculate_importance_score(message),
                        'is_read': not message.UnRead
                    }
                    emails.append(email_data)
                    print(f"✅ מייל {i+1}: {email_data['subject'][:30]}...")
                    
                except Exception as e:
                    print(f"❌ שגיאה במייל {i+1}: {e}")
                    continue
            
            print(f"📧 הוחזרו {len(emails)} מיילים אמיתיים")
            return emails
            
        except Exception as e:
            print(f"❌ שגיאה בקבלת מיילים מ-Outlook: {e}")
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
                'is_read': False
            },
            {
                'id': 4,
                'subject': 'Weekly Newsletter - Tech Updates',
                'sender': 'TechNews',
                'sender_email': 'newsletter@technews.com',
                'received_time': str(datetime.now() - timedelta(days=2)),
                'body_preview': 'This week in tech: New AI developments, startup funding, and industry trends...',
                'importance_score': 0.3,
                'is_read': True
            },
            {
                'id': 5,
                'subject': 'Urgent: Server Maintenance Tonight',
                'sender': 'IT Department',
                'sender_email': 'it@company.com',
                'received_time': str(datetime.now() - timedelta(hours=1)),
                'body_preview': 'We will be performing server maintenance tonight from 11 PM to 3 AM. Please save your work...',
                'importance_score': 0.9,
                'is_read': False
            }
        ]
        return sample_emails
    
    def get_emails(self, limit=20):
        """קבלת מיילים - אמיתיים או דמה"""
        if self.use_real_outlook and self.connect_to_outlook():
            return self.get_emails_from_outlook(limit)
        else:
            print("📧 משתמש בנתונים דמה")
            return self.get_sample_emails()[:limit]
    
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
            time_diff = datetime.now() - message.ReceivedTime
            if time_diff.days < 1:
                score += 0.2
            elif time_diff.days < 7:
                score += 0.1
            
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

if __name__ == '__main__':
    print("🚀 מפעיל את Outlook Email Manager...")
    print("📧 מנסה להתחבר ל-Outlook...")
    
    if email_manager.connect_to_outlook():
        print("✅ חיבור ל-Outlook הצליח!")
        print("🌐 מפעיל שרת web על http://localhost:5000")
    else:
        print("⚠️ לא ניתן להתחבר ל-Outlook - משתמש בנתונים דמה")
        print("🌐 מפעיל שרת web על http://localhost:5000")
    
    app.run(debug=True, host='127.0.0.1', port=5000)





