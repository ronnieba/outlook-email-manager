# -*- coding: utf-8 -*-
"""
יצירת מיילים פשוטים בכמות גדולה
"""

import win32com.client
import random

SUBJECTS = [
    "דוח שבועי - {} {}",
    "עדכון חשוב - {}",
    "פגישה לתיאום - {}",
    "בקשה לאישור {}",
    "תזכורת: {}",
    "URGENT: {}",
    "סיכום {} - {}",
    "שאלה בנוגע ל-{}",
    "מידע על {}",
    "תיאום {} השבוע"
]

TOPICS = ["Azure", "CRM", "Website", "Mobile App", "Security", "Infrastructure", 
          "Marketing", "Sales", "HR", "IT", "Finance", "Operations"]

SENDERS = [
    ("דני כהן", "danny@company.com"),
    ("שירה לוי", "shira@company.com"),
    ("יוסי אברהם", "yossi@company.com"),
    ("מיכל רוזנברג", "michal@company.com"),
    ("איתן גולד", "eitan@company.com"),
]

def create_emails(count=227):
    """יצירת מיילים פשוטים"""
    print(f"יוצר {count} מיילים...")
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    
    created = 0
    
    for i in range(count):
        try:
            mail = outlook.CreateItem(0)
            
            # נושא פשוט
            subject_template = random.choice(SUBJECTS)
            topic = random.choice(TOPICS)
            month = random.choice(["ינואר", "פברואר", "מרץ", "אפריל"])
            
            if subject_template.count("{}") == 2:
                subject = subject_template.format(topic, month)
            else:
                subject = subject_template.format(topic)
            
            # שולח
            sender_name, sender_email = random.choice(SENDERS)
            
            # תוכן פשוט
            body = f"""שלום,

עדכון בנושא {topic}:

• סטטוס: בתהליך
• התקדמות: {random.randint(30, 90)}%
• משימות: {random.randint(5, 20)}

אשמח למשוב.

בברכה,
{sender_name}"""
            
            mail.Subject = subject
            mail.Body = f"מאת: {sender_name} <{sender_email}>\n\n{body}"
            mail.Save()
            mail.Move(inbox)
            
            created += 1
            if (i + 1) % 50 == 0:
                print(f"✅ {i + 1}...")
                
        except Exception as e:
            print(f"❌ שגיאה {i + 1}: {e}")
    
    print(f"\n✅ נוצרו {created}/{count} מיילים")
    return created

if __name__ == "__main__":
    print("="*60)
    print("🎯 יצירת מיילים פשוטים")
    print("="*60)
    created = create_emails(227)  # להשלים ל-320 נוספים
    print("="*60)
    print(f"✅ הושלם! נוצרו {created} מיילים")
    print("="*60)

