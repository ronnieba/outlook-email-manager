"""
Outlook Email Manager - Demo Version with Sample Data
מערכת ניהול מיילים חכמה עם AI - גרסת דמו
"""
from flask import Flask, render_template, request, jsonify
import json
import os
from datetime import datetime, timedelta
import sqlite3
import random

app = Flask(__name__)

class EmailManager:
    def __init__(self):
        self.user_preferences = {}
        self.db_path = "email_preferences.db"
        self.init_database()
        self.load_user_preferences()
        
        # נתונים דמה לבדיקה
        self.sample_emails = self.create_sample_emails()
    
    def create_sample_emails(self):
        """יצירת נתונים דמה לבדיקה"""
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
            },
            {
                'id': 6,
                'subject': 'Invoice #12345 - Payment Due',
                'sender': 'Accounting Department',
                'sender_email': 'accounting@company.com',
                'received_time': str(datetime.now() - timedelta(days=3)),
                'body_preview': 'Your invoice #12345 is due for payment. Please process payment by the end of the month...',
                'importance_score': 0.6,
                'is_read': True
            },
            {
                'id': 7,
                'subject': 'Happy Birthday! 🎉',
                'sender': 'Office Manager',
                'sender_email': 'office@company.com',
                'received_time': str(datetime.now() - timedelta(hours=8)),
                'body_preview': 'Wishing you a very happy birthday! We have a small celebration planned in the break room...',
                'importance_score': 0.4,
                'is_read': False
            },
            {
                'id': 8,
                'subject': 'Project Deadline Extension',
                'sender': 'Project Manager',
                'sender_email': 'pm@company.com',
                'received_time': str(datetime.now() - timedelta(hours=12)),
                'body_preview': 'Good news! We have received approval to extend the project deadline by one week...',
                'importance_score': 0.7,
                'is_read': True
            },
            {
                'id': 9,
                'subject': 'Security Alert: Suspicious Activity',
                'sender': 'Security Team',
                'sender_email': 'security@company.com',
                'received_time': str(datetime.now() - timedelta(minutes=30)),
                'body_preview': 'We detected suspicious activity on your account. Please change your password immediately...',
                'importance_score': 0.95,
                'is_read': False
            },
            {
                'id': 10,
                'subject': 'Company Picnic Invitation',
                'sender': 'Events Team',
                'sender_email': 'events@company.com',
                'received_time': str(datetime.now() - timedelta(days=5)),
                'body_preview': 'Join us for our annual company picnic this Saturday at Central Park. Food and activities provided...',
                'importance_score': 0.2,
                'is_read': True
            }
        ]
        
        return sample_emails
    
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
        
        conn.commit()
        conn.close()
    
    def get_emails(self, limit=20):
        """קבלת מיילים (נתונים דמה)"""
        print(f"📧 מחזיר {len(self.sample_emails)} מיילים דמה")
        return self.sample_emails[:limit]
    
    def calculate_importance_score(self, email):
        """חישוב ציון חשיבות למייל"""
        return email.get('importance_score', 0.5)
    
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

if __name__ == '__main__':
    print("🚀 מפעיל את Outlook Email Manager - גרסת דמו...")
    print("📧 משתמש בנתונים דמה לבדיקה")
    print("🌐 מפעיל שרת web על http://localhost:5000")
    print("✨ המערכת מוכנה לשימוש!")
    
    app.run(debug=True, host='127.0.0.1', port=5000)








