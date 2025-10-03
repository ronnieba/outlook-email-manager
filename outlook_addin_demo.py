"""
Outlook Add-in Demo - תוסף Outlook לדוגמה
מחבר את המערכת הקיימת ל-Outlook ומציג את הניתוח
"""

import win32com.client
import requests
import json
import time
from datetime import datetime

class OutlookAddinDemo:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.server_url = "http://localhost:5000"
        
    def connect_to_outlook(self):
        """חיבור ל-Outlook"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            print("חובר ל-Outlook בהצלחה!")
            return True
        except Exception as e:
            print(f"שגיאה בחיבור ל-Outlook: {e}")
            return False
    
    def create_ai_column(self):
        """יצירת עמודה AI אוטומטית ב-Outlook"""
        try:
            # קבלת התצוגה הנוכחית
            explorer = self.outlook.ActiveExplorer()
            if not explorer:
                print("לא ניתן לגשת לתצוגת Outlook")
                return False
            
            # קבלת התצוגה הנוכחית
            view = explorer.CurrentView
            
            # ניסיון ליצור עמודה באמצעות VBA
            try:
                # יצירת עמודה מותאמת אישית באמצעות VBA
                vba_code = """
                Sub CreateAIColumn()
                    Dim objView As Outlook.View
                    Dim objViewField As Outlook.ViewField
                    
                    Set objView = Application.ActiveExplorer.CurrentView
                    
                    ' יצירת שדה מותאם אישית
                    Set objViewField = objView.ViewFields.Add("AI_Score")
                    objViewField.ColumnFormat = 1  ' Text format
                    
                    ' שמירת התצוגה
                    objView.Save
                    
                    ' רענון התצוגה
                    Application.ActiveExplorer.Refresh
                End Sub
                """
                
                # הפעלת ה-VBA
                self.outlook.Application.Run("CreateAIColumn")
                
                print("עמודה AI_Score נוצרה בהצלחה באמצעות VBA!")
                return True
                
            except Exception as e:
                print(f"VBA לא עבד: {e}")
                
                # ניסיון חלופי - שימוש בשדה קיים
                try:
                    # ניסיון להוסיף שדה קיים
                    view.ViewFields.Add("Subject")  # שדה קיים
                    print("נוסף שדה Subject כעמודה זמנית")
                    
                    # הוספת הודעה למשתמש
                    print("\n" + "="*50)
                    print("הוראות ליצירת עמודה AI ידנית:")
                    print("="*50)
                    print("1. פתח את Outlook")
                    print("2. לחץ על 'תצוגה' (View)")
                    print("3. לחץ על 'הגדרות תצוגה' (View Settings)")
                    print("4. לחץ על 'עמודות' (Columns)")
                    print("5. לחץ על 'חדש...' (New...)")
                    print("6. הזן שם: AI_Score")
                    print("7. בחר סוג: טקסט (Text)")
                    print("8. לחץ 'אישור'")
                    print("9. גרור את השדה החדש לתצוגה")
                    print("10. לחץ 'אישור'")
                    print("="*50)
                    
                    return False
                    
                except Exception as e2:
                    print(f"גם זה לא עבד: {e2}")
                    return False
                    
        except Exception as e:
            print(f"שגיאה ביצירת עמודה AI: {e}")
            return False
    
    def test_server_connection(self):
        """בדיקת חיבור לשרת"""
        try:
            response = requests.get(f"{self.server_url}/api/status", timeout=5)
            if response.status_code == 200:
                print("השרת זמין!")
                return True
            else:
                print(f"השרת לא זמין: {response.status_code}")
                return False
        except Exception as e:
            print(f"שגיאה בחיבור לשרת: {e}")
            return False
    
    def analyze_email_with_server(self, email_data):
        """ניתוח מייל עם השרת"""
        try:
            response = requests.post(
                f"{self.server_url}/api/analyze-email",
                json=email_data,
                timeout=10
            )
            
            if response.status_code == 200:
                return response.json()
            else:
                print(f"שגיאה בניתוח: {response.status_code}")
                return None
                
        except Exception as e:
            print(f"שגיאה בניתוח מייל: {e}")
            return None
    
    def add_analysis_to_email(self, mail_item, analysis):
        """הוספת הניתוח למייל"""
        try:
            # לא משנים את גוף המייל - רק מוסיפים metadata
            
            # הוספת מידע נוסף (Custom Properties) - זה יאפשר ליצור עמודה מותאמת
            try:
                # יצירת Custom Property בשם "AI_Score" עם הציון
                importance_percent = int(analysis['importance_score'] * 100)
                
                # ניסיון ליצור Custom Property
                try:
                    mail_item.UserProperties.Add("AI_Score", 1, True)  # 1 = Text
                except:
                    pass  # אם כבר קיים
                
                mail_item.UserProperties("AI_Score").Value = f"{importance_percent}%"
                
                # הוספת קטגוריה
                mail_item.UserProperties.Add("AI_Category", 1, True)
                mail_item.UserProperties("AI_Category").Value = analysis['category']
                
                print(f"נוסף ציון AI: {importance_percent}% למייל: {mail_item.Subject}")
                
            except Exception as e:
                print(f"שגיאה בהוספת Custom Properties: {e}")
            
            # הוספת דגל לפי חשיבות (רק אם רוצים)
            try:
                if analysis['importance_score'] >= 0.8:
                    mail_item.FlagRequest = "Follow up"
                elif analysis['importance_score'] >= 0.6:
                    mail_item.FlagRequest = "No Response Necessary"
            except:
                pass
            
            # שמירה
            mail_item.Save()
            
            print("הניתוח נוסף למייל בהצלחה!")
            return True
            
        except Exception as e:
            print(f"שגיאה בהוספת הניתוח: {e}")
            return False
    
    def analyze_current_email(self):
        """ניתוח המייל הנוכחי"""
        try:
            # קבלת המייל הנוכחי
            selection = self.outlook.ActiveExplorer().Selection
            if selection.Count == 0:
                print("לא נבחר מייל")
                return False
            
            mail_item = selection[0]
            
            # הכנת הנתונים לניתוח
            email_data = {
                'subject': mail_item.Subject or '',
                'sender': mail_item.SenderName or '',
                'body_preview': mail_item.Body[:500] if mail_item.Body else '',  # 500 תווים ראשונים
                'received_time': mail_item.ReceivedTime.isoformat() if hasattr(mail_item, 'ReceivedTime') else '',
                'has_attachments': mail_item.Attachments.Count > 0
            }
            
            print(f"מנתח מייל: {email_data['subject']}")
            
            # ניתוח עם השרת
            analysis = self.analyze_email_with_server(email_data)
            
            if analysis:
                # הוספת הניתוח למייל
                return self.add_analysis_to_email(mail_item, analysis)
            else:
                print("לא ניתן לנתח את המייל")
                return False
                
        except Exception as e:
            print(f"שגיאה בניתוח המייל הנוכחי: {e}")
            return False
    
    def analyze_selected_emails(self):
        """ניתוח כל המיילים הנבחרים"""
        try:
            selection = self.outlook.ActiveExplorer().Selection
            if selection.Count == 0:
                print("לא נבחרו מיילים")
                return False
            
            print(f"מנתח {selection.Count} מיילים...")
            
            success_count = 0
            for i in range(selection.Count):
                try:
                    mail_item = selection[i]
                    
                    # הכנת הנתונים לניתוח
                    email_data = {
                        'subject': mail_item.Subject or '',
                        'sender': mail_item.SenderName or '',
                        'body_preview': mail_item.Body[:500] if mail_item.Body else '',
                        'received_time': mail_item.ReceivedTime.isoformat() if hasattr(mail_item, 'ReceivedTime') else '',
                        'has_attachments': mail_item.Attachments.Count > 0
                    }
                    
                    print(f"מנתח מייל {i+1}/{selection.Count}: {email_data['subject']}")
                    
                    # ניתוח עם השרת
                    analysis = self.analyze_email_with_server(email_data)
                    
                    if analysis:
                        # הוספת הניתוח למייל
                        if self.add_analysis_to_email(mail_item, analysis):
                            success_count += 1
                    
                    # המתנה קצרה בין מיילים
                    time.sleep(0.5)
                    
                except Exception as e:
                    print(f"שגיאה במייל {i+1}: {e}")
                    continue
            
            print(f"נותחו בהצלחה {success_count} מתוך {selection.Count} מיילים")
            return success_count > 0
            
        except Exception as e:
            print(f"שגיאה בניתוח המיילים: {e}")
            return False

def main():
    """פונקציה ראשית"""
    print("מתחיל תוסף Outlook Demo...")
    
    # יצירת התוסף
    addin = OutlookAddinDemo()
    
    # חיבור ל-Outlook
    if not addin.connect_to_outlook():
        return
    
    # בדיקת חיבור לשרת
    if not addin.test_server_connection():
        print("השרת לא זמין. ודא שהשרת רץ על localhost:5000")
        return
    
    print("\nאפשרויות:")
    print("1. ניתוח המייל הנוכחי")
    print("2. ניתוח כל המיילים הנבחרים")
    print("3. יצירת עמודה AI אוטומטית")
    print("4. יציאה")
    
    while True:
        try:
            choice = input("\nבחר אפשרות (1-4): ").strip()
            
            if choice == "1":
                addin.analyze_current_email()
            elif choice == "2":
                addin.analyze_selected_emails()
            elif choice == "3":
                addin.create_ai_column()
            elif choice == "4":
                print("תודה לשימוש!")
                break
            else:
                print("בחירה לא תקינה")
                
        except KeyboardInterrupt:
            print("\nתודה לשימוש!")
            break
        except Exception as e:
            print(f"שגיאה: {e}")

if __name__ == "__main__":
    main()
