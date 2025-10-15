# -*- coding: utf-8 -*-
"""
יצירת כמות גדולה של מיילים ופגישות ב-Outlook
"""

import win32com.client
from datetime import datetime, timedelta
import random

# בסיס תוכן למיילים
EMAIL_SUBJECTS = [
    "דוח שבועי - פרויקט {}",
    "URGENT: בעיה ב-{}",
    "הזמנה לפגישה - {}",
    "סיכום פגישת {} Q4",
    "תוצאות {} - חודש {}",
    "בקשה לאישור {}",
    "עדכון חשוב - {}",
    "תזכורת: {} - דדליין {}",
    "שאלה לגבי {}",
    "בדיקה נדרשת - {}",
    "הצעה חדשה - {}",
    "אישור נדרש - {}",
    "דחוף: {} דורש תשומת לב",
    "מידע חשוב על {}",
    "תיאום {} לשבוע הבא"
]

PROJECTS = ["Azure Migration", "CRM Upgrade", "Website Redesign", "Mobile App", 
            "Security Audit", "Infrastructure", "Cloud Services", "API Integration",
            "Database Optimization", "User Training", "Marketing Campaign", "Sales Process"]

SENDERS = [
    ("דני כהן", "danny.cohen@company.com"),
    ("שירה לוי", "shira.levi@company.com"),
    ("יוסי אברהם", "yossi.a@company.com"),
    ("מיכל רוזנברג", "michal.r@company.com"),
    ("איתן גולד", "eitan.gold@company.com"),
    ("רונית שפירא", "ronit.s@company.com"),
    ("אלון כהן", "alon.cohen@company.com"),
    ("טל לוי", "tal.levi@company.com"),
    ("דנה שמש", "dana.shemesh@company.com"),
    ("רוני מור", "roni.mor@company.com")
]

EMAIL_BODIES = [
    """שלום,

מצורף עדכון על הפרויקט:

📊 סטטוס:
- התקדמות: {}%
- משימות שהושלמו: {}
- משימות נותרות: {}

📅 לוח זמנים:
- יעד השלמה: {}
- פגישת מעקב: {}

בברכה,
{}""",
    """היי,

דרוש {} בנושא הבא:

• נושא: {}
• עדיפות: {}
• דדליין: {}

תודה,
{}""",
    """שלום רב,

בהמשך לשיחה שלנו, הנה הפרטים:

✓ {}
✓ {}
✓ {}

אשמח למשוב.

בברכה,
{}"""
]

MEETING_SUBJECTS = [
    "סטנדאפ צוות {}",
    "ישיבת {} - תכנון",
    "1-on-1 עם {}",
    "הדרכה: {}",
    "פגישת לקוח - {}",
    "Code Review - {}",
    "Demo - {}",
    "ברייסטורם - {}",
    "סקירת {} שבועית",
    "אישור {} וביצוע"
]

MEETING_LOCATIONS = [
    "Zoom Meeting",
    "Microsoft Teams",
    "חדר ישיבות A",
    "חדר ישיבות B",
    "משרד המנהל",
    "מעבדת מחשבים",
    "אולם ההרצאות",
    "משרדי הלקוח"
]

def create_bulk_emails(count=320):
    """יצירת מיילים בכמות גדולה"""
    print(f"\n{'='*60}")
    print(f"יוצר {count} מיילים...")
    print(f"{'='*60}\n")
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    
    created = 0
    
    for i in range(count):
        try:
            mail = outlook.CreateItem(0)
            
            # בחירת נושא אקראי
            subject_template = random.choice(EMAIL_SUBJECTS)
            project = random.choice(PROJECTS)
            month = random.choice(["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", 
                                  "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"])
            
            subject = subject_template.format(project)
            if "{}" in subject:
                subject = subject.format(month)
            
            # בחירת שולח
            sender_name, sender_email = random.choice(SENDERS)
            
            # יצירת תוכן
            body_template = random.choice(EMAIL_BODIES)
            progress = random.randint(10, 95)
            completed = random.randint(5, 20)
            remaining = random.randint(1, 10)
            
            date1 = (datetime.now() + timedelta(days=random.randint(1, 30))).strftime("%d/%m/%Y")
            date2 = (datetime.now() + timedelta(days=random.randint(1, 14))).strftime("%d/%m/%Y")
            
            priority = random.choice(["גבוהה", "בינונית", "נמוכה"])
            action = random.choice(["אישור", "עדכון", "סקירה", "תיאום", "החלטה"])
            
            if len(body_template.format("", "", "", "", "", "").split("{}")) > 6:
                body = body_template.format(progress, completed, remaining, date1, date2, sender_name)
            elif len(body_template.format("", "", "", "", "").split("{}")) > 5:
                body = body_template.format(action, project, priority, date1, sender_name)
            else:
                body = body_template.format(project, f"פרט חשוב על {project}", 
                                          f"נושא נוסף בנוגע ל-{project}", sender_name)
            
            # הגדרת המייל
            mail.Subject = subject
            mail.Body = f"מאת: {sender_name} <{sender_email}>\n\n{body}"
            
            # שמירה
            mail.Save()
            mail.Move(inbox)
            
            created += 1
            if (i + 1) % 50 == 0:
                print(f"✅ נוצרו {i + 1} מיילים...")
                
        except Exception as e:
            print(f"❌ שגיאה במייל {i + 1}: {e}")
    
    print(f"\n✅ סיימתי! נוצרו {created}/{count} מיילים")
    return created

def create_bulk_meetings(count=63):
    """יצירת פגישות בכמות גדולה"""
    print(f"\n{'='*60}")
    print(f"יוצר {count} פגישות...")
    print(f"{'='*60}\n")
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    
    base_date = datetime.now() + timedelta(days=1)
    created = 0
    
    for i in range(count):
        try:
            meeting = outlook.CreateItem(1)
            
            # בחירת נושא
            subject_template = random.choice(MEETING_SUBJECTS)
            project = random.choice(PROJECTS)
            team_member = random.choice(SENDERS)[0]
            
            subject = subject_template.format(project if "{}" in subject_template else team_member)
            
            # תאריך ושעה אקראיים
            days_ahead = random.randint(1, 60)
            hour = random.choice([9, 10, 11, 13, 14, 15, 16])
            minute = random.choice([0, 30])
            
            start_time = base_date + timedelta(days=days_ahead)
            start_time = start_time.replace(hour=hour, minute=minute, second=0, microsecond=0)
            
            duration = random.choice([30, 45, 60, 90, 120])
            location = random.choice(MEETING_LOCATIONS)
            
            # תוכן הפגישה
            body = f"""סדר יום:
1. פתיחה ומטרות ({random.randint(5, 10)} דקות)
2. {project} - סטטוס ועדכונים ({random.randint(10, 20)} דקות)
3. דיון ודיון ({random.randint(10, 30)} דקות)
4. סיכום ומשימות ({random.randint(5, 10)} דקות)

משתתפים:
{random.choice(SENDERS)[0]}, {random.choice(SENDERS)[0]}, {random.choice(SENDERS)[0]}

הערות:
נא להגיע מוכנים עם עדכונים"""
            
            # הגדרת הפגישה
            meeting.Subject = subject
            meeting.Location = location
            meeting.Body = body
            meeting.Start = start_time
            meeting.Duration = duration
            meeting.ReminderSet = True
            meeting.ReminderMinutesBeforeStart = 15
            
            # שמירה
            meeting.Save()
            
            created += 1
            if (i + 1) % 10 == 0:
                print(f"✅ נוצרו {i + 1} פגישות...")
                
        except Exception as e:
            print(f"❌ שגיאה בפגישה {i + 1}: {e}")
    
    print(f"\n✅ סיימתי! נוצרו {created}/{count} פגישות")
    return created

def main():
    print("="*60)
    print("🎯 יצירת מיילים ופגישות בכמות גדולה")
    print("="*60)
    
    # יצירת 320 מיילים
    emails = create_bulk_emails(320)
    
    # יצירת 63 פגישות
    meetings = create_bulk_meetings(63)
    
    # סיכום
    print("\n" + "="*60)
    print("✅ הושלם!")
    print("="*60)
    print(f"📧 מיילים שנוצרו: {emails}")
    print(f"📅 פגישות שנוצרו: {meetings}")
    print(f"📊 סה״כ: {emails + meetings}")
    print("="*60)

if __name__ == "__main__":
    main()

