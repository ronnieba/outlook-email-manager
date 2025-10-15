# -*- coding: utf-8 -*-
"""
יצירת מיילים ופגישות לדוגמה ב-Outlook
"""

import win32com.client
from datetime import datetime, timedelta
import random

def create_test_emails():
    """יצירת מיילים לדוגמה"""
    print("מתחבר ל-Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
    
    # רשימת מיילים לדוגמה
    emails = [
        {
            "sender": "דני כהן <danny.cohen@company.com>",
            "subject": "דוח שבועי - פרויקט Azure Migration",
            "body": """שלום,

מצורף דוח התקדמות שבועי לפרויקט ההעברה ל-Azure:

📊 סטטוס נוכחי:
- 75% מהשרתים עברו בהצלחה
- 3 שרתים נותרו להעברה
- ביצועים משופרים ב-40%

⚠️ נושאים דורשי תשומת לב:
1. שרת DB-PROD דורש אישור מנהל IT
2. יש לתאם downtime עם צוות התמיכה
3. נדרשת הדרכה לצוות על הסביבה החדשה

📅 לוח זמנים:
- סיום צפוי: 25/10/2024
- פגישת סיכום: 30/10/2024

אשמח לתיאום פגישה להצגת הממצאים.

בברכה,
דני כהן
מנהל פרויקט
טלפון: 052-1234567"""
        },
        {
            "sender": "שירה לוי <shira.levi@hr.company.com>",
            "subject": "הזמנה לסדנת פיתוח מקצועי - 28/10",
            "body": """שלום רב,

אנו שמחים להזמינך לסדנה מקצועית בנושא:
"ניהול זמן יעיל וסדרי עדיפויות"

📅 מתי: יום שלישי, 28/10/2024
🕐 שעה: 10:00-13:00
📍 איפה: חדר ההדרכה, קומה 3
👤 מרצה: ד"ר יוסי ברק

🎯 נושאי הסדנה:
• טכניקות ניהול זמן מתקדמות
• קביעת סדרי עדיפויות נכונים
• ניהול משימות במקביל
• כלים דיגיטליים לפרודוקטיביות

☕ ארוחת בוקר קלה תוגש

נא לאשר השתתפות עד 24/10.
מספר המקומות מוגבל!

בברכה,
שירה לוי
משאבי אנוש"""
        },
        {
            "sender": "יוסי אברהם <yossi.abraham@microsoft.com>",
            "subject": "⚠️ URGENT - בעיה קריטית בסביבת הייצור",
            "body": """⚠️ דחוף - נדרשת תשומת לב מיידית! ⚠️

זוהתה בעיה קריטית בסביבת הייצור:

🔴 סוג הבעיה: שירות Authentication לא מגיב
🕐 זמן תחילת הבעיה: 14:30
📊 השפעה: כל המשתמשים לא יכולים להתחבר
⏱️ זמן השבתה: כ-2 שעות

פעולות שבוצעו עד כה:
1. ✅ Restart לשירות - לא עזר
2. ✅ בדיקת Logs - נמצאה שגיאת חיבור ל-DB
3. ⏳ פתיחת טיקט ל-DBA Team
4. ⏳ הפעלת Backup Server

נדרשות הפעולות הבאות:
• אישור מנכ"ל IT להפעלת DR Site
• עדכון ללקוחות על התקלה
• הקמת War Room

אנא התקשר אליי מיד: 054-9876543

בדחיפות,
יוסי אברהם
ראש צוות Operations"""
        },
        {
            "sender": "מיכל רוזנברג <michal.r@finance.company.com>",
            "subject": "סיכום פגישת תקציב Q4 2024",
            "body": """שלום לכולם,

להלן סיכום הפגישה מהבוקר בנושא תקציב הרבעון הרביעי:

💰 תקציב מאושר: 2.5M ₪

חלוקה לפי מחלקות:
• R&D: 1M ₪ (40%)
• Marketing: 600K ₪ (24%)
• Operations: 500K ₪ (20%)
• HR: 400K ₪ (16%)

📌 החלטות:
1. העברת 100K מMarketing ל-R&D בהסכמת המנכ"ל
2. הקפאת גיוס עד ינואר 2025
3. אישור רכישת ציוד IT חדש - עד 150K

📅 לוח זמנים:
- 20/10 - הגשת בקשות תקציב סופיות
- 25/10 - אישור סופי מההנהלה
- 01/11 - תחילת ביצוע

נא לשלוח הערות עד סוף השבוע.

בברכה,
מיכל רוזנברג
מנהלת כספים"""
        },
        {
            "sender": "איתן גולד <eitan.gold@marketing.company.com>",
            "subject": "תוצאות קמפיין הפרסום - ספטמבר 2024",
            "body": """היי צוות!

הנתונים נכנסו והם מעולים! 🎉

📈 תוצאות הקמפיין:
• Impressions: 450,000 (↑ 35% מהיעד)
• Clicks: 12,500 (CTR: 2.78%)
• Conversions: 850 (↑ 42%)
• ROI: 320% 💪

🎯 הביצועים הטובים ביותר:
1. Facebook Ads - 380 conversions
2. Google Search - 290 conversions
3. LinkedIn - 180 conversions

💡 תובנות מרכזיות:
• קהל היעד בגילאי 25-34 הכי מעורב
• שעות הערב (19:00-22:00) הכי אפקטיביות
• תוכן וידאו מניב פי 3 יותר engagement

📅 השלבים הבאים:
- 18/10: פגישת תכנון לקמפיין אוקטובר
- 20/10: הגשת תקציב לחודש הבא
- 25/10: השקת Landing Page חדש

Best,
איתן גולד
Digital Marketing Manager"""
        },
        {
            "sender": "רונית שפירא <ronit.shapira@legal.company.com>",
            "subject": "בקשה לאישור חוזה עם ספק חדש",
            "body": """שלום,

אנא מצא בצירוף טיוטת חוזה עם הספק CloudTech Solutions.

📄 פרטי החוזה:
• שם ספק: CloudTech Solutions Ltd.
• שירות: Cloud Infrastructure Management
• תקופה: 24 חודשים
• עלות שנתית: 180,000 ₪
• מועד התחלה: 01/11/2024

⚖️ נקודות משפטיות לבדיקה:
1. סעיף 8 - אחריות וביטוח
2. סעיף 12 - סיום מוקדם של ההסכם
3. נספח ג' - רמות שירות (SLA)

דרושים אישורים:
✅ מנהל IT - אושר
✅ מנהל כספים - אושר
⏳ יועץ משפטי - ממתין
⏳ מנכ"ל - ממתין

⏰ דחוף: נדרש אישור עד 22/10 על מנת לעמוד בלוח הזמנים.

בברכה,
רו"ח רונית שפירא
יועצת משפטית"""
        },
        {
            "sender": "AWS Notifications <no-reply@aws.amazon.com>",
            "subject": "AWS Monthly Bill - October 2024",
            "body": """Dear Customer,

Your AWS monthly bill for October 2024 is now available.

Account ID: 123456789012
Billing Period: Oct 1, 2024 - Oct 31, 2024
Total Amount Due: $2,847.52

Service Breakdown:
• Amazon EC2: $1,245.00
• Amazon S3: $456.30
• Amazon RDS: $823.22
• Amazon CloudFront: $323.00

Payment Due Date: October 25, 2024

To view your detailed bill, please log in to the AWS Billing Console.

Thank you for using Amazon Web Services.

Best regards,
AWS Billing Team"""
        },
        {
            "sender": "LinkedIn <messages-noreply@linkedin.com>",
            "subject": "יש לך 3 הודעות חדשות מחיבורים שלך",
            "body": """שלום,

יש לך הודעות חדשות ב-LinkedIn:

👤 אלון כהן שלח לך הודעה
"היי, ראיתי את הפרופיל שלך ואשמח לשוחח על הזדמנות תעסוקתית..."

👤 טל לוי רוצה להתחבר אליך
"שלום, אני עובד בחברת Microsoft ואשמח להתחבר..."

👤 דנה שמש המליצה עליך
"ממליצה בחום על רוני - מקצוען מעולה!"

📊 סטטיסטיקות הפרופיל שלך השבוע:
• 47 צפיות בפרופיל (↑ 12%)
• 8 חיפושים שהופעת בהם
• 15 חיבורים חדשים

💼 משרות שעשויות להתאים לך:
• Senior Developer - Microsoft
• Team Lead - Google
• Tech Manager - Amazon

בברכה,
צוות LinkedIn"""
        },
        {
            "sender": "GitHub <notifications@github.com>",
            "subject": "[Project] New Pull Request requires your review",
            "body": """Hi there,

A new pull request has been opened and requires your review:

PR #234: Implement user authentication system
Author: @johndoe
Repository: company/main-project

Changes:
• Added JWT authentication
• Implemented password hashing
• Created login/logout endpoints
• Added unit tests (85% coverage)

Files changed: 12 files (+456 -123)

Please review and approve or request changes.

View Pull Request:
https://github.com/company/main-project/pull/234

Best,
GitHub Team"""
        },
        {
            "sender": "Microsoft Teams <no-reply@microsoft.com>",
            "subject": "You were mentioned in 'Project Updates' channel",
            "body": """Hi,

You were mentioned in a conversation:

Team: Engineering
Channel: Project Updates

Sarah Johnson:
@רוני - can you please review the deployment plan for next week?
We need your approval before moving forward.

Reply in Teams:
https://teams.microsoft.com/l/message/...

Microsoft Teams"""
        },
        {
            "sender": "DocuSign <dse@docusign.net>",
            "subject": "Please sign: Annual Contract Renewal 2024",
            "body": """Action Required: Please Sign Document

You have been sent a document to sign:

Document: Annual Contract Renewal 2024
From: Legal Department
Status: Awaiting your signature
Expires: October 25, 2024

Important Information:
• This is a time-sensitive document
• Your signature is required to complete the process
• The document will expire in 5 days

To review and sign the document, please click below:
[Review Document]

If you have any questions, please contact:
legal@company.com

DocuSign - The Global Standard for Digital Transaction Management"""
        },
        {
            "sender": "Azure DevOps <noreply@dev.azure.com>",
            "subject": "Build Failed: main-project-CI #145",
            "body": """Build Failed

Repository: company/main-project
Branch: main
Build: #145
Triggered by: Auto (commit)

Errors:
1. Unit test failure in AuthenticationTests.cs
2. Code coverage below threshold (75% required, got 68%)
3. SonarQube quality gate failed

Details:
• Duration: 5m 32s
• Tests run: 127
• Tests passed: 124
• Tests failed: 3

Failed Tests:
- TestUserLogin_InvalidPassword
- TestTokenExpiration
- TestRefreshToken

Please fix the issues and push again.

View build details:
https://dev.azure.com/company/main-project/_build/results?buildId=145

Azure DevOps"""
        }
    ]
    
    print(f"\nיוצר {len(emails)} מיילים...")
    created_count = 0
    
    for i, email_data in enumerate(emails, 1):
        try:
            # יצירת מייל חדש
            mail = outlook.CreateItem(0)  # 0 = MailItem
            
            # הגדרת פרטי המייל
            mail.Subject = email_data["subject"]
            mail.Body = f"מאת: {email_data['sender']}\n\n{email_data['body']}"
            
            # שמירה ישירות ב-Inbox
            mail.Save()
            mail.Move(inbox)
            
            created_count += 1
            print(f"✅ {i}. נוצר מייל: {email_data['subject'][:60]}...")
            
        except Exception as e:
            print(f"❌ שגיאה ביצירת מייל {i}: {e}")
    
    print(f"\n✅ סיימתי! נוצרו {created_count} מיילים בהצלחה")
    return created_count

def create_test_meetings():
    """יצירת פגישות לדוגמה"""
    print("\n" + "="*60)
    print("מתחבר ל-Outlook לפגישות...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    
    # בסיס לתאריכים - מהשבוע הבא
    base_date = datetime.now() + timedelta(days=7)
    base_date = base_date.replace(hour=9, minute=0, second=0, microsecond=0)
    
    # רשימת פגישות לדוגמה
    meetings = [
        {
            "subject": "סטנדאפ צוות - עדכון שבועי",
            "location": "Zoom - https://zoom.us/j/123456789",
            "start": base_date.replace(hour=9, minute=0),
            "duration": 30,
            "body": """📅 פגישת סטנדאפ שבועית

סדר יום:
1. עדכוני פרויקטים (10 דקות)
2. חסמים ובעיות (10 דקות)
3. תכנון השבוע (10 דקות)

משתתפים:
• דני כהן - מנהל פרויקט
• שירה לוי - Team Lead
• יוסי אברהם - Developer

קישור Zoom: https://zoom.us/j/123456789"""
        },
        {
            "subject": "ישיבת הנהלה - Q4 Planning",
            "location": "חדר ישיבות A, קומה 5",
            "start": base_date.replace(hour=14, minute=0),
            "duration": 120,
            "body": """🎯 ישיבת תכנון רבעונית

נושאים לדיון:
1. סקירת ביצועים Q3 (20 דקות)
2. יעדים ל-Q4 (30 דקות)
3. תקציב ומשאבים (30 דקות)
4. פרויקטים חדשים (20 דקות)

משתתפים:
• מנכ"ל
• סמנכ"ל כספים
• מנהלי מחלקות"""
        },
        {
            "subject": "Demo - מערכת CRM החדשה",
            "location": "Microsoft Teams",
            "start": base_date.replace(hour=10, minute=30) + timedelta(days=1),
            "duration": 45,
            "body": """🎬 הדגמת מערכת CRM החדשה

מה נראה:
• ממשק משתמש מחודש
• ניהול לידים משופר
• אינטגרציה עם Outlook
• דוחות ואנליטיקה

Teams Link: https://teams.microsoft.com/l/meetup-join/..."""
        },
        {
            "subject": "1-on-1 עם מנהל - ביקורת ביצועים",
            "location": "משרד המנהל, קומה 3",
            "start": base_date.replace(hour=15, minute=0) + timedelta(days=2),
            "duration": 60,
            "body": """👤 פגישה אישית - ביקורת רבעונית

נושאים לשיחה:
1. סקירת הישגים ברבעון
2. אתגרים וקשיים
3. מטרות לרבעון הבא
4. משוב דו-כיווני"""
        },
        {
            "subject": "הדרכה טכנית - Azure DevOps",
            "location": "מעבדת מחשבים, קומה 2",
            "start": base_date.replace(hour=13, minute=0) + timedelta(days=3),
            "duration": 180,
            "body": """🎓 הדרכה: Azure DevOps Fundamentals

תוכנית ההדרכה:
13:00-14:00 | חלק 1: מבוא
14:00-15:00 | חלק 2: Hands-on
15:00-15:15 | הפסקה
15:15-16:00 | חלק 3: מתקדם

מרצה: יוסי אברהם"""
        },
        {
            "subject": "פגישת לקוח - חברת TechCorp",
            "location": "משרדי הלקוח, תל אביב",
            "start": base_date.replace(hour=10, minute=0) + timedelta(days=4),
            "duration": 90,
            "body": """🤝 פגישה עם לקוח: TechCorp Solutions

מטרת הפגישה:
• סקירת פרויקט ה-AI שלהם
• הצגת הצעת המחיר שלנו
• דיון בלוח זמנים וציפיות

מיקום: TechCorp Tower, רחוב הארבעה 4, תל אביב"""
        },
        {
            "subject": "Code Review Session - Sprint 12",
            "location": "Zoom Meeting",
            "start": base_date.replace(hour=16, minute=0) + timedelta(days=4),
            "duration": 60,
            "body": """👨‍💻 Code Review - Sprint 12

Pull Requests לסקירה:
1. PR #145 - User Authentication Module
2. PR #146 - API Rate Limiting
3. PR #147 - UI/UX Updates

Zoom Link: https://zoom.us/j/code-review-123"""
        },
        {
            "subject": "All-Hands Meeting - חגיגת הצלחות",
            "location": "אולם האירועים, קומת קרקע",
            "start": base_date.replace(hour=17, minute=0) + timedelta(days=5),
            "duration": 120,
            "body": """🎉 All-Hands Meeting + חגיגה!

סדר היום:
17:00 | ברכה ופתיחה
17:20 | הצגות צוותים
18:00 | הכרה והוקרה
18:20 | עדכונים
18:40 | חגיגה חופשית

נושא השנה: Together We Achieve More"""
        }
    ]
    
    print(f"\nיוצר {len(meetings)} פגישות...")
    created_count = 0
    
    for i, meeting_data in enumerate(meetings, 1):
        try:
            # יצירת פגישה חדשה
            meeting = outlook.CreateItem(1)  # 1 = AppointmentItem
            
            # הגדרת פרטי הפגישה
            meeting.Subject = meeting_data["subject"]
            meeting.Location = meeting_data["location"]
            meeting.Body = meeting_data["body"]
            meeting.Start = meeting_data["start"]
            meeting.Duration = meeting_data["duration"]
            
            # הוספת תזכורת
            meeting.ReminderSet = True
            meeting.ReminderMinutesBeforeStart = 15
            
            # שמירה
            meeting.Save()
            
            created_count += 1
            start_time = meeting_data["start"].strftime("%d/%m %H:%M")
            print(f"✅ {i}. נוצרה פגישה: {meeting_data['subject'][:50]}... ({start_time})")
            
        except Exception as e:
            print(f"❌ שגיאה ביצירת פגישה {i}: {e}")
    
    print(f"\n✅ סיימתי! נוצרו {created_count} פגישות בהצלחה")
    return created_count

def main():
    """פונקציה ראשית"""
    print("="*60)
    print("🎯 יצירת מיילים ופגישות לדוגמה ב-Outlook")
    print("="*60)
    
    try:
        # יצירת מיילים
        emails_created = create_test_emails()
        
        # יצירת פגישות
        meetings_created = create_test_meetings()
        
        # סיכום
        print("\n" + "="*60)
        print("✅ הושלם בהצלחה!")
        print("="*60)
        print(f"📧 מיילים שנוצרו: {emails_created}")
        print(f"📅 פגישות שנוצרו: {meetings_created}")
        print(f"📊 סהכ פריטים: {emails_created + meetings_created}")
        print("\n💡 עכשיו אפשר לנתח אותם עם:")
        print("   python working_email_analyzer.py")
        print("="*60)
        
    except Exception as e:
        print(f"\n❌ שגיאה כללית: {e}")
        import traceback
        traceback.print_exc()
    
    input("\nלחץ Enter לסגירה...")

if __name__ == "__main__":
    main()
