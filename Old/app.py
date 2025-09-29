"""
Outlook Email Manager - Web Server
מערכת ניהול מיילים חכמה עם AI
"""
from flask import Flask, render_template, request, jsonify
import win32com.client
import json
import os
from datetime import datetime
import sqlite3

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
            return True
        except Exception as e:
            print(f"שגיאה בחיבור ל-Outlook: {e}")
            return False
    
    def get_emails(self, limit=50):
        """קבלת מיילים מ-Outlook"""
        if not self.inbox:
            if not self.connect_to_outlook():
                return []
        
        try:
            messages = self.inbox.Items
            emails = []
            
            for i in range(min(limit, messages.Count)):
                message = messages[i + 1]
                email_data = {
                    'id': i + 1,
                    'subject': message.Subject,
                    'sender': message.SenderName,
                    'sender_email': message.SenderEmailAddress,
                    'received_time': str(message.ReceivedTime),
                    'body_preview': message.Body[:200] + "..." if len(message.Body) > 200 else message.Body,
                    'importance_score': self.calculate_importance_score(message),
                    'is_read': message.UnRead == False
                }
                emails.append(email_data)
            
            return emails
        except Exception as e:
            print(f"שגיאה בקבלת מיילים: {e}")
            return []
    
    def calculate_importance_score(self, message):
        """חישוב ציון חשיבות למייל"""
        score = 0.5  # ציון בסיסי
        
        # בדיקת מילות מפתח חשובות
        important_keywords = ['חשוב', 'דחוף', 'urgent', 'important', 'meeting', 'פגישה']
        subject = message.Subject.lower()
        body = message.Body.lower()
        
        for keyword in important_keywords:
            if keyword in subject:
                score += 0.2
            if keyword in body:
                score += 0.1
        
        # בדיקת שולח חשוב
        important_senders = ['manager', 'boss', 'מנהל', 'hr', 'it']
        sender = message.SenderName.lower()
        
        for important_sender in important_senders:
            if important_sender in sender:
                score += 0.3
        
        # בדיקת זמן - מיילים חדשים יותר חשובים
        time_diff = datetime.now() - message.ReceivedTime
        if time_diff.days < 1:
            score += 0.2
        elif time_diff.days < 7:
            score += 0.1
        
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

# יצירת מופע של מנהל המיילים
email_manager = EmailManager()

@app.route('/')
def index():
    """דף הבית"""
    return render_template('index.html')

@app.route('/api/emails')
def get_emails():
    """API לקבלת מיילים"""
    emails = email_manager.get_emails()
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
    emails = email_manager.get_emails()
    # מיון לפי ציון חשיבות
    important_emails = sorted(emails, key=lambda x: x['importance_score'], reverse=True)
    return jsonify(important_emails[:10])  # 10 המיילים החשובים ביותר

if __name__ == '__main__':
    print("🚀 מפעיל את Outlook Email Manager...")
    print("📧 חיבור ל-Outlook...")
    
    if email_manager.connect_to_outlook():
        print("✅ חיבור ל-Outlook הצליח!")
        print("🌐 מפעיל שרת web על http://localhost:5000")
        app.run(debug=True, host='0.0.0.0', port=5000)
    else:
        print("❌ לא ניתן להתחבר ל-Outlook. ודא ש-Outlook פתוח.")


