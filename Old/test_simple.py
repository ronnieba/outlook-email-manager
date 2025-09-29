"""
בדיקה פשוטה של חיבור ל-Outlook
"""
import win32com.client

def test_simple():
    try:
        print("מתחבר ל-Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # Inbox
        
        print("✅ חיבור הצליח!")
        print(f"📧 מספר מיילים: {inbox.Items.Count}")
        
        # נסה לקבל מייל אחד
        if inbox.Items.Count > 0:
            message = inbox.Items[1]  # המייל הראשון
            print(f"📧 נושא: {message.Subject}")
            print(f"👤 שולח: {message.SenderName}")
            print(f"🕒 זמן: {message.ReceivedTime}")
            print(f"📖 נקרא: {not message.UnRead}")
        else:
            print("❌ אין מיילים בתיקייה")
            
    except Exception as e:
        print(f"❌ שגיאה: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_simple()








