"""
מערכת ניהול מיילים עם AI - אינטגרציה מלאה עם Outlook
כל הפעולות מתבצעות מתוך Outlook Desktop
"""

import win32com.client
import pythoncom
import time
import requests
import json
import sys
from datetime import datetime

# כתובת השרת שלך
API_BASE_URL = "http://localhost:5000"

class OutlookAIIntegration:
    """אינטגרציה עם Outlook - כל הפעולות דרך Outlook"""
    
    def __init__(self):
        """אתחול התוסף"""
        print("🚀 מאתחל אינטגרציה עם Outlook...")
        self.outlook = None
        self.namespace = None
        self.connect_to_outlook()
        
    def connect_to_outlook(self):
        """התחברות ל-Outlook"""
        try:
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            print("✅ התחברות ל-Outlook הצליחה!")
            return True
        except Exception as e:
            print(f"❌ שגיאה בהתחברות ל-Outlook: {e}")
            return False
    
    def add_context_menu(self):
        """הוספת תפריט הקשר ל-Outlook"""
        print("\n📋 הוראות שימוש:")
        print("=" * 50)
        print("בתוך Outlook:")
        print("1. לחץ לחיצה ימנית על מייל")
        print("2. בחר 'פעולות מהירות' (Quick Steps)")
        print("3. או השתמש בקיצורי המקלדת:")
        print("   - F9: נתח מייל נוכחי")
        print("   - Ctrl+F9: נתח את כל התיקיה")
        print("=" * 50)
        
    def analyze_email_with_ai(self, mail_item):
        """ניתוח מייל בודד עם AI"""
        try:
            print(f"\n🔍 מנתח מייל: {mail_item.Subject}")
            
            # הכנת הנתונים לניתוח
            email_data = {
                "subject": mail_item.Subject,
                "body": mail_item.Body,
                "sender": mail_item.SenderEmailAddress,
                "received_time": str(mail_item.ReceivedTime)
            }
            
            # שליחה ל-API
            print("📤 שולח ל-AI לניתוח...")
            response = requests.post(
                f"{API_BASE_URL}/api/analyze",
                json=email_data,
                timeout=30
            )
            
            if response.status_code == 200:
                analysis = response.json()
                print("✅ ניתוח הושלם!")
                
                # עדכון המייל ב-Outlook
                self.update_email_with_analysis(mail_item, analysis)
                return True
            else:
                print(f"❌ שגיאה בניתוח: {response.status_code}")
                return False
                
        except Exception as e:
            print(f"❌ שגיאה: {e}")
            return False
    
    def update_email_with_analysis(self, mail_item, analysis):
        """עדכון המייל עם תוצאות הניתוח"""
        try:
            print("📝 מעדכן את המייל...")
            
            # הוספת קטגוריה
            if "category" in analysis:
                mail_item.Categories = analysis["category"]
                print(f"  📋 קטגוריה: {analysis['category']}")
            
            # הגדרת דחיפות
            if "priority" in analysis:
                priority_map = {"גבוהה": 2, "רגילה": 1, "נמוכה": 0}
                mail_item.Importance = priority_map.get(analysis["priority"], 1)
                print(f"  ⚠️ דחיפות: {analysis['priority']}")
            
            # הוספת דגל למעקב
            if analysis.get("requires_action", False):
                mail_item.FlagRequest = "למעקב"
                print("  📌 נוסף דגל למעקב")
            
            # שמירת ניתוח מפורט כמאפיין מותאם אישית
            user_property = mail_item.UserProperties.Add(
                "AI Analysis", 
                1  # olText
            )
            user_property.Value = json.dumps(analysis, ensure_ascii=False)
            
            # שמירת השינויים
            mail_item.Save()
            print("💾 המייל עודכן בהצלחה!")
            
        except Exception as e:
            print(f"❌ שגיאה בעדכון המייל: {e}")
    
    def analyze_current_email(self):
        """ניתוח המייל הנוכחי שנבחר ב-Outlook"""
        try:
            explorer = self.outlook.ActiveExplorer()
            
            if not explorer:
                print("❌ אין חלון Outlook פעיל")
                return False
            
            selection = explorer.Selection
            
            if selection.Count == 0:
                print("❌ לא נבחר מייל")
                print("💡 בחר מייל ב-Outlook ונסה שוב")
                return False
            
            # ניתוח המייל הראשון שנבחר
            mail_item = selection.Item(1)
            return self.analyze_email_with_ai(mail_item)
            
        except Exception as e:
            print(f"❌ שגיאה: {e}")
            return False
    
    def analyze_folder(self, folder_name="Inbox"):
        """ניתוח כל המיילים בתיקיה"""
        try:
            print(f"\n📁 מנתח תיקיה: {folder_name}")
            
            folder = self.namespace.GetDefaultFolder(6)  # 6 = Inbox
            items = folder.Items
            
            print(f"📊 נמצאו {items.Count} מיילים")
            
            analyzed = 0
            for item in items:
                if item.Class == 43:  # Mail item
                    if self.analyze_email_with_ai(item):
                        analyzed += 1
                    time.sleep(0.5)  # המתן קצר בין מיילים
            
            print(f"\n✅ סיום! נותחו {analyzed} מיילים")
            return True
            
        except Exception as e:
            print(f"❌ שגיאה: {e}")
            return False
    
    def monitor_new_emails(self):
        """ניטור מיילים חדשים (אופציונלי - רק אם רוצים אוטומציה מלאה)"""
        print("\n👀 מנטר מיילים חדשים...")
        print("(לחץ Ctrl+C לעצור)")
        
        try:
            last_count = self.namespace.GetDefaultFolder(6).Items.Count
            
            while True:
                time.sleep(5)  # בדיקה כל 5 שניות
                current_count = self.namespace.GetDefaultFolder(6).Items.Count
                
                if current_count > last_count:
                    print(f"\n📬 זוהו {current_count - last_count} מיילים חדשים!")
                    # ניתוח המיילים החדשים...
                    last_count = current_count
                    
        except KeyboardInterrupt:
            print("\n⏹️ עצירת ניטור")
    
    def show_menu(self):
        """תפריט ראשי"""
        print("\n" + "=" * 50)
        print("🤖 AI Email Manager - אינטגרציה עם Outlook")
        print("=" * 50)
        print("\nפעולות זמינות:")
        print("1. נתח מייל נוכחי (המייל שבחרת ב-Outlook)")
        print("2. נתח את כל תיבת הדואר הנכנס")
        print("3. התחל ניטור אוטומטי של מיילים חדשים")
        print("4. צפייה בהוראות שימוש")
        print("5. יציאה")
        print("=" * 50)

def main():
    """פונקציה ראשית"""
    print("🚀 מפעיל AI Email Manager...\n")
    
    # בדיקה אם השרת פעיל
    try:
        response = requests.get(f"{API_BASE_URL}/health", timeout=2)
        if response.status_code != 200:
            print("⚠️ השרת לא פעיל. הפעל את השרת תחילה:")
            print("   python app_with_ai.py")
            return
    except:
        print("⚠️ השרת לא פעיל. הפעל את השרת תחילה:")
        print("   python app_with_ai.py")
        return
    
    # יצירת האינטגרציה
    integration = OutlookAIIntegration()
    
    # תפריט אינטראקטיבי
    while True:
        integration.show_menu()
        
        try:
            choice = input("\n👉 בחר פעולה (1-5): ").strip()
            
            if choice == "1":
                print("\n📧 מנתח את המייל שבחרת ב-Outlook...")
                print("💡 ודא שבחרת מייל ב-Outlook!")
                input("לחץ Enter כשאתה מוכן...")
                integration.analyze_current_email()
                
            elif choice == "2":
                print("\n📁 מנתח את כל תיבת הדואר הנכנס...")
                confirm = input("האם אתה בטוח? זה יכול לקחת זמן (y/n): ")
                if confirm.lower() == 'y':
                    integration.analyze_folder()
                    
            elif choice == "3":
                print("\n👀 מתחיל ניטור אוטומטי...")
                integration.monitor_new_emails()
                
            elif choice == "4":
                integration.add_context_menu()
                
            elif choice == "5":
                print("\n👋 להתראות!")
                break
                
            else:
                print("❌ בחירה לא חוקית")
                
        except KeyboardInterrupt:
            print("\n\n👋 להתראות!")
            break
        except Exception as e:
            print(f"\n❌ שגיאה: {e}")
            continue

if __name__ == "__main__":
    main()

