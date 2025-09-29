# -*- coding: utf-8 -*-
"""
Outlook Email Manager - Simple Version with Console Output
"""
from flask import Flask, render_template, jsonify
import win32com.client
from datetime import datetime
import pythoncom
import threading

app = Flask(__name__)

class OutlookManager:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.inbox = None
        self.connected = False

    def connect_to_outlook(self):
        if self.connected:
            return True
        try:
            print("🔍 מתחבר ל-Outlook...")
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)  # 6 = Inbox
            self.connected = True
            print("✅ חיבור ל-Outlook הצליח!")
            return True
        except Exception as e:
            print(f"❌ שגיאה בחיבור ל-Outlook: {e}")
            self.connected = False
            return False
        finally:
            pythoncom.CoUninitialize()

    def get_emails(self, limit=20):
        if not self.connected:
            if not self.connect_to_outlook():
                print("📧 Outlook לא עובד - עובר לנתונים דמה")
                return self._get_dummy_emails(limit)

        try:
            pythoncom.CoInitialize()
            messages = self.inbox.Items
            messages.Sort("[ReceivedTime]", True)
            emails = []
            print(f"📧 נמצאו {messages.Count} מיילים ב-Outlook")

            for i in range(min(limit, messages.Count)):
                try:
                    message = messages[i + 1]
                    subject = str(message.Subject) if message.Subject else "ללא נושא"
                    sender = str(message.SenderName) if message.SenderName else "שולח לא ידוע"
                    received_time = message.ReceivedTime
                    body_preview = str(message.Body)[:200] + "..." if len(str(message.Body)) > 200 else str(message.Body)
                    is_read = not message.UnRead

                    emails.append({
                        'id': i + 1,
                        'subject': subject,
                        'sender': sender,
                        'received_time': str(received_time),
                        'body_preview': body_preview,
                        'importance_score': 0.5,
                        'is_read': is_read
                    })
                    print(f"✅ מייל {i+1}: {subject[:50]}...")
                except Exception as e:
                    print(f"שגיאה במייל {i+1}: {e}")
                    continue
            
            print(f"📧 הוחזרו {len(emails)} מיילים אמיתיים")
            return emails
        except Exception as e:
            print(f"❌ שגיאה בקבלת מיילים מ-Outlook: {e}")
            print("📧 Outlook לא עובד - עובר לנתונים דמה")
            return self._get_dummy_emails(limit)
        finally:
            pythoncom.CoUninitialize()

    def _get_dummy_emails(self, limit=20):
        print("📧 יוצר נתונים דמה...")
        dummy_emails = []
        for i in range(limit):
            dummy_emails.append({
                'id': i + 1,
                'subject': f'מייל דמה {i + 1}',
                'sender': f'שולח דמה {i + 1}',
                'received_time': str(datetime.now()),
                'body_preview': f'תוכן מייל דמה {i + 1}...',
                'importance_score': 0.3,
                'is_read': False
            })
        print(f"📧 הוחזרו {len(dummy_emails)} מיילים דמה")
        return dummy_emails

# יצירת מנהל Outlook
outlook_manager = OutlookManager()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/emails')
def get_emails():
    emails = outlook_manager.get_emails(20)
    return jsonify(emails)

@app.route('/api/stats')
def get_stats():
    emails = outlook_manager.get_emails(50)
    total_emails = len(emails)
    unread_emails = sum(1 for email in emails if not email['is_read'])
    important_emails = sum(1 for email in emails if email['importance_score'] > 0.7)
    
    stats = {
        'total_emails': total_emails,
        'unread_emails': unread_emails,
        'important_emails': important_emails,
        'read_emails': total_emails - unread_emails
    }
    return jsonify(stats)

if __name__ == '__main__':
    print("🚀 מתחיל את השרת...")
    print("🔍 בודק חיבור ל-Outlook...")
    
    if outlook_manager.connect_to_outlook():
        print("✅ חיבור ל-Outlook הצליח!")
        print("🌐 מפעיל שרת web על http://localhost:5000")
    else:
        print("⚠️ לא ניתן להתחבר ל-Outlook - משתמש בנתונים דמה")
        print("🌐 מפעיל שרת web על http://localhost:5000")
    
    print("🎯 השרת מוכן! פתח דפדפן על: http://localhost:5000")
    app.run(debug=True, host='127.0.0.1', port=5000)








