"""
Outlook Demo Data Creator
יוצר מיילים ופגישות דמה ישירות ב-Outlook
"""

import win32com.client
from datetime import datetime, timedelta
import random
import time

class OutlookDemoDataCreator:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.inbox = None
        self.calendar = None
        
    def connect_to_outlook(self):
        """התחברות ל-Outlook"""
        try:
            print("🔌 Connecting to Outlook...")
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            self.calendar = self.namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
            print("✅ Connected to Outlook successfully!")
            return True
        except Exception as e:
            print(f"❌ Error connecting to Outlook: {e}")
            return False
    
    def create_demo_emails(self, count=30):
        """יוצר מיילים דמה ב-Outlook"""
        if not self.outlook:
            print("❌ Not connected to Outlook")
            return False
        
        print(f"📧 Creating {count} demo emails...")
        
        hebrew_subjects = [
            "פגישה חשובה מחר", "דוח רבעוני - דחוף", "הצעת מחיר חדשה", "פגישת צוות שבועית",
            "עדכון פרויקט", "הזמנה לפגישה", "מסמכים לחתימה", "תזכורת: פגישה",
            "הצעת עבודה", "סקר שביעות רצון", "הזמנה לאירוע", "עדכון מערכת",
            "בדיקת אבטחה", "הדרכה חדשה", "פגישת הערכה"
        ]
        
        english_subjects = [
            "Important Meeting Tomorrow", "Quarterly Report - Urgent", "New Price Quote",
            "Weekly Team Meeting", "Project Update", "Meeting Invitation", "Documents for Signature",
            "Reminder: Meeting", "Job Offer", "Satisfaction Survey", "Event Invitation",
            "System Update", "Security Check", "New Training", "Performance Review"
        ]
        
        hebrew_bodies = [
            "שלום, אני רוצה לתאם איתך פגישה חשובה לגבי הפרויקט החדש. האם תוכל להיפגש מחר ב-10:00?",
            "הדוח הרבעוני מוכן וצריך את החתימה שלך. אנא קרא את הקובץ המצורף.",
            "יש לנו הצעת מחיר חדשה ללקוח חשוב. צריך את האישור שלך עד סוף השבוע.",
            "פגישת הצוות השבועית תתקיים ביום רביעי ב-14:00. אנא הביאו את הדוחות שלכם.",
            "הפרויקט מתקדם יפה. יש כמה נקודות שצריך לדון בהן. מתי נוח לך?",
            "הזמנה לפגישה עם הלקוח החדש. נפגש ביום שלישי ב-15:00 במשרד.",
            "יש מסמכים שצריכים חתימה דחופה. אנא תגיע למשרד היום.",
            "תזכורת: יש לנו פגישה מחר ב-9:00. אל תשכח להביא את החומרים.",
            "יש לנו הצעת עבודה מעניינת עבורך. רוצה לשמוע פרטים?",
            "אנא מלא את סקר שביעות הרצון. זה יעזור לנו לשפר את השירות."
        ]
        
        english_bodies = [
            "Hello, I would like to schedule an important meeting about the new project. Can you meet tomorrow at 10:00?",
            "The quarterly report is ready and needs your signature. Please review the attached file.",
            "We have a new price quote for an important client. Need your approval by end of week.",
            "Weekly team meeting will be held on Wednesday at 2:00 PM. Please bring your reports.",
            "The project is progressing well. There are a few points we need to discuss. When is convenient for you?",
            "Invitation to meeting with new client. We'll meet on Tuesday at 3:00 PM in the office.",
            "There are documents that need urgent signature. Please come to the office today.",
            "Reminder: We have a meeting tomorrow at 9:00 AM. Don't forget to bring the materials.",
            "We have an interesting job offer for you. Would you like to hear the details?",
            "Please fill out the satisfaction survey. It will help us improve our service."
        ]
        
        hebrew_names = [
            "דוד כהן", "שרה לוי", "מיכל אברהם", "יוסי ישראלי", "רחל גולדברג",
            "אבי רוזן", "מיכל שטרן", "דני כהן", "נועה לוי", "אלי ישראלי"
        ]
        
        english_names = [
            "John Smith", "Sarah Johnson", "Michael Brown", "Emily Davis", "David Wilson",
            "Lisa Anderson", "Robert Taylor", "Jennifer Thomas", "William Jackson", "Maria Garcia"
        ]
        
        companies = [
            "Microsoft", "Google", "Apple", "Amazon", "Meta", "Netflix", "Tesla", "Uber",
            "חברת טכנולוגיה", "בנק הפועלים", "משרד עורכי דין", "קליניקה רפואית"
        ]
        
        email_domains = [
            "gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "company.com", "tech.co.il"
        ]
        
        for i in range(count):
            try:
                # יצירת מייל חדש
                mail = self.outlook.CreateItem(0)  # 0 = olMailItem
                
                # בחירת שם ושפה
                use_hebrew = random.choice([True, False])
                if use_hebrew:
                    sender_name = random.choice(hebrew_names)
                    subject = random.choice(hebrew_subjects)
                    body = random.choice(hebrew_bodies)
                else:
                    sender_name = random.choice(english_names)
                    subject = random.choice(english_subjects)
                    body = random.choice(english_bodies)
                
                # הגדרת פרטי המייל
                mail.Subject = subject
                mail.Body = body
                # הוספת שם השולח בתוכן המייל
                mail.Body = f"From: {sender_name}\nEmail: {sender_name.lower().replace(' ', '.')}@{random.choice(email_domains)}\n\n{body}"
                
                # יצירת זמן אקראי
                days_ago = random.randint(0, 30)
                hours_ago = random.randint(0, 23)
                minutes_ago = random.randint(0, 59)
                
                received_time = datetime.now() - timedelta(
                    days=days_ago,
                    hours=hours_ago,
                    minutes=minutes_ago
                )
                
                # הוספת המייל לתיקיית ה-Inbox
                mail.Move(self.inbox)
                
                # שמירת המייל
                mail.Save()
                
                print(f"✅ Created email {i+1}: {subject[:30]}...")
                
                # השהיה קצרה כדי לא להעמיס על Outlook
                time.sleep(0.1)
                
            except Exception as e:
                print(f"❌ Error creating email {i+1}: {e}")
                continue
        
        print(f"📧 Created {count} demo emails successfully!")
        return True
    
    def create_demo_meetings(self, count=15):
        """יוצר פגישות דמה ב-Outlook"""
        if not self.outlook:
            print("❌ Not connected to Outlook")
            return False
        
        print(f"📅 Creating {count} demo meetings...")
        
        hebrew_subjects = [
            "פגישת צוות שבועית", "פגישה עם הלקוח", "פגישת פרויקט", "פגישת הערכה",
            "פגישת תכנון", "פגישת סקירה", "פגישת הדרכה", "פגישת אסטרטגיה"
        ]
        
        english_subjects = [
            "Weekly Team Meeting", "Client Meeting", "Project Meeting", "Performance Review",
            "Planning Meeting", "Review Meeting", "Training Session", "Strategy Meeting"
        ]
        
        locations = [
            "Conference Room 1", "Conference Room 2", "Manager's Office", "Auditorium",
            "חדר ישיבות 1", "חדר ישיבות 2", "משרד המנהל", "אולם כנסים"
        ]
        
        hebrew_names = [
            "דוד כהן", "שרה לוי", "מיכל אברהם", "יוסי ישראלי", "רחל גולדברג",
            "אבי רוזן", "מיכל שטרן", "דני כהן", "נועה לוי", "אלי ישראלי"
        ]
        
        english_names = [
            "John Smith", "Sarah Johnson", "Michael Brown", "Emily Davis", "David Wilson",
            "Lisa Anderson", "Robert Taylor", "Jennifer Thomas", "William Jackson", "Maria Garcia"
        ]
        
        for i in range(count):
            try:
                # יצירת פגישה חדשה
                appointment = self.outlook.CreateItem(1)  # 1 = olAppointmentItem
                
                # בחירת שם ושפה
                use_hebrew = random.choice([True, False])
                if use_hebrew:
                    subject = random.choice(hebrew_subjects)
                    organizer = random.choice(hebrew_names)
                else:
                    subject = random.choice(english_subjects)
                    organizer = random.choice(english_names)
                
                # הגדרת פרטי הפגישה
                appointment.Subject = subject
                appointment.Location = random.choice(locations)
                appointment.Organizer = organizer
                
                # יצירת זמן אקראי
                days_ahead = random.randint(0, 30)
                hours_ahead = random.randint(9, 17)
                minutes_ahead = random.choice([0, 15, 30, 45])
                
                start_time = datetime.now() + timedelta(
                    days=days_ahead,
                    hours=hours_ahead,
                    minutes=minutes_ahead
                )
                
                duration = random.choice([30, 60, 90, 120])
                end_time = start_time + timedelta(minutes=duration)
                
                appointment.Start = start_time
                appointment.End = end_time
                appointment.Duration = duration
                
                # הוספת תיאור
                appointment.Body = f"Meeting: {subject}\\n\\nLocation: {appointment.Location}\\n\\nOrganizer: {organizer}"
                
                # שמירת הפגישה
                appointment.Save()
                
                print(f"✅ Created meeting {i+1}: {subject[:30]}...")
                
                # השהיה קצרה כדי לא להעמיס על Outlook
                time.sleep(0.1)
                
            except Exception as e:
                print(f"❌ Error creating meeting {i+1}: {e}")
                continue
        
        print(f"📅 Created {count} demo meetings successfully!")
        return True
    
    def create_all_demo_data(self):
        """יוצר את כל נתוני הדמה"""
        if not self.connect_to_outlook():
            return False
        
        print("🚀 Starting demo data creation...")
        
        # יצירת מיילים
        if self.create_demo_emails(30):
            print("✅ Demo emails created successfully!")
        else:
            print("❌ Failed to create demo emails")
        
        # יצירת פגישות
        if self.create_demo_meetings(15):
            print("✅ Demo meetings created successfully!")
        else:
            print("❌ Failed to create demo meetings")
        
        print("🎉 Demo data creation completed!")
        return True

if __name__ == "__main__":
    creator = OutlookDemoDataCreator()
    creator.create_all_demo_data()
