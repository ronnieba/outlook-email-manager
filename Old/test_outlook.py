"""
בדיקת חיבור ל-Outlook באמצעות COM Object
"""
import win32com.client
import sys

def test_outlook_connection():
    try:
        print("מתחבר ל-Outlook...")
        
        # חיבור ל-Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("✅ חיבור ל-Outlook הצליח!")
        
        # קבלת namespace
        namespace = outlook.GetNamespace("MAPI")
        print("✅ קבלת Namespace הצליחה!")
        
        # קבלת תיקיית Inbox
        inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
        print("✅ קבלת תיקיית Inbox הצליחה!")
        
        # ספירת מיילים
        messages = inbox.Items
        message_count = messages.Count
        print(f"📧 נמצאו {message_count} מיילים בתיקיית Inbox")
        
        # הצגת 3 מיילים אחרונים
        print("\n📋 3 המיילים האחרונים:")
        for i in range(min(3, message_count)):
            message = messages[i + 1]  # Outlook מתחיל מ-1
            subject = message.Subject
            sender = message.SenderName
            received_time = message.ReceivedTime
            print(f"  {i+1}. {subject[:50]}... - {sender} ({received_time})")
        
        return True
        
    except Exception as e:
        print(f"❌ שגיאה בחיבור ל-Outlook: {e}")
        return False

if __name__ == "__main__":
    print("🔍 בודק חיבור ל-Outlook...")
    success = test_outlook_connection()
    
    if success:
        print("\n🎉 החיבור ל-Outlook עובד! אפשר להמשיך עם הפרויקט.")
    else:
        print("\n⚠️ יש בעיה בחיבור. נצטרך לפתור את זה קודם.")
