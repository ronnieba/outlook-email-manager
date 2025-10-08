# -*- coding: utf-8 -*-
"""
AI Email Manager - תוסף Outlook ללא COM
גישה פשוטה יותר - תוסף Python עצמאי
"""

import win32com.client
import requests
import json
import time
from datetime import datetime
import os

class StandaloneOutlookAddin:
    """תוסף Outlook עצמאי ללא COM"""
    
    def __init__(self):
        self.outlook = None
        self.server_url = "http://localhost:5000"
        self.log_file = os.path.join(os.environ.get('TEMP', os.getcwd()), 'standalone_addin.log')
        
    def log_message(self, message):
        """רישום הודעות"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"{timestamp} - {message}\n"
        
        try:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(log_entry)
        except:
            pass
        
        print(f"[{timestamp}] {message}")
    
    def connect_to_outlook(self):
        """חיבור ל-Outlook"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.log_message("חובר ל-Outlook בהצלחה")
            return True
        except Exception as e:
            self.log_message(f"שגיאה בחיבור ל-Outlook: {e}")
            return False
    
    def test_server_connection(self):
        """בדיקת חיבור לשרת"""
        try:
            response = requests.get(f"{self.server_url}/api/status", timeout=5)
            if response.status_code == 200:
                self.log_message("השרת זמין")
                return True
            else:
                self.log_message(f"השרת לא זמין: {response.status_code}")
                return False
        except Exception as e:
            self.log_message(f"שגיאה בחיבור לשרת: {e}")
            return False
    
    def analyze_email_with_server(self, email_data):
        """ניתוח מייל עם השרת"""
        try:
            response = requests.post(
                f"{self.server_url}/api/outlook-addin/analyze-email",
                json=email_data,
                timeout=30
            )
            
            if response.status_code == 200:
                return response.json()
            else:
                self.log_message(f"שגיאה בניתוח: {response.status_code}")
                return None
                
        except Exception as e:
            self.log_message(f"שגיאה בניתוח מייל: {e}")
            return None
    
    def add_analysis_to_email(self, mail_item, analysis):
        """הוספת הניתוח למייל"""
        try:
            # הוספת Custom Properties
            importance_percent = int(analysis.get('importance_score', 0) * 100)
            
            # ציון חשיבות
            try:
                mail_item.UserProperties.Add("AI_Score", 1, True)  # 1 = Text
            except:
                pass  # אם כבר קיים
            
            mail_item.UserProperties("AI_Score").Value = f"{importance_percent}%"
            
            # קטגוריה
            try:
                mail_item.UserProperties.Add("AI_Category", 1, True)
            except:
                pass
            
            mail_item.UserProperties("AI_Category").Value = analysis.get('category', 'לא סווג')
            
            # סיכום
            try:
                mail_item.UserProperties.Add("AI_Summary", 1, True)
            except:
                pass
            
            mail_item.UserProperties("AI_Summary").Value = analysis.get('summary', '')[:255]
            
            # תאריך ניתוח
            try:
                mail_item.UserProperties.Add("AI_Analyzed", 1, True)
            except:
                pass
            
            mail_item.UserProperties("AI_Analyzed").Value = datetime.now().strftime("%Y-%m-%d %H:%M")
            
            # הוספת דגל לפי חשיבות
            if importance_percent >= 80:
                mail_item.FlagRequest = "Follow up"
            elif importance_percent >= 60:
                mail_item.FlagRequest = "No Response Necessary"
            
            # שמירה
            mail_item.Save()
            
            self.log_message(f"ניתוח נוסף למייל: {mail_item.Subject}")
            return True
            
        except Exception as e:
            self.log_message(f"שגיאה בהוספת הניתוח: {e}")
            return False
    
    def analyze_current_email(self):
        """ניתוח המייל הנוכחי"""
        try:
            # קבלת המייל הנוכחי
            selection = self.outlook.ActiveExplorer().Selection
            if selection.Count == 0:
                self.log_message("לא נבחר מייל")
                return False
            
            mail_item = selection[0]
            
            # הכנת הנתונים לניתוח
            email_data = {
                'subject': mail_item.Subject or '',
                'sender': mail_item.SenderName or '',
                'body': mail_item.Body or '',
                'sender_email': mail_item.SenderEmailAddress or '',
                'received_time': mail_item.ReceivedTime.isoformat() if hasattr(mail_item, 'ReceivedTime') else '',
                'has_attachments': mail_item.Attachments.Count > 0
            }
            
            self.log_message(f"מנתח מייל: {email_data['subject']}")
            
            # ניתוח עם השרת
            analysis = self.analyze_email_with_server(email_data)
            
            if analysis and analysis.get("success"):
                # הוספת הניתוח למייל
                if self.add_analysis_to_email(mail_item, analysis):
                    score = int(analysis.get('importance_score', 0) * 100)
                    category = analysis.get('category', 'לא סווג')
                    summary = analysis.get('summary', 'לא נמצא סיכום')
                    
                    print(f"\n{'='*50}")
                    print(f"ניתוח הושלם בהצלחה!")
                    print(f"{'='*50}")
                    print(f"📊 ציון חשיבות: {score}%")
                    print(f"🏷️ קטגוריה: {category}")
                    print(f"📝 סיכום: {summary}")
                    print(f"{'='*50}\n")
                    
                    return True
                else:
                    self.log_message("לא ניתן להוסיף את הניתוח למייל")
                    return False
            else:
                self.log_message("לא ניתן לנתח את המייל")
                return False
                
        except Exception as e:
            self.log_message(f"שגיאה בניתוח המייל הנוכחי: {e}")
            return False
    
    def analyze_selected_emails(self):
        """ניתוח כל המיילים הנבחרים"""
        try:
            selection = self.outlook.ActiveExplorer().Selection
            if selection.Count == 0:
                self.log_message("לא נבחרו מיילים")
                return False
            
            count = selection.Count
            self.log_message(f"מנתח {count} מיילים...")
            
            success_count = 0
            for i in range(count):
                try:
                    mail_item = selection[i]
                    
                    # הכנת הנתונים לניתוח
                    email_data = {
                        'subject': mail_item.Subject or '',
                        'sender': mail_item.SenderName or '',
                        'body': mail_item.Body or '',
                        'sender_email': mail_item.SenderEmailAddress or '',
                        'received_time': mail_item.ReceivedTime.isoformat() if hasattr(mail_item, 'ReceivedTime') else '',
                        'has_attachments': mail_item.Attachments.Count > 0
                    }
                    
                    self.log_message(f"מנתח מייל {i+1}/{count}: {email_data['subject']}")
                    
                    # ניתוח עם השרת
                    analysis = self.analyze_email_with_server(email_data)
                    
                    if analysis and analysis.get("success"):
                        # הוספת הניתוח למייל
                        if self.add_analysis_to_email(mail_item, analysis):
                            success_count += 1
                    
                    # המתנה קצרה בין מיילים
                    time.sleep(0.5)
                    
                except Exception as e:
                    self.log_message(f"שגיאה במייל {i+1}: {e}")
                    continue
            
            self.log_message(f"נותחו בהצלחה {success_count} מתוך {count} מיילים")
            print(f"\n{'='*50}")
            print(f"ניתוח הושלם!")
            print(f"נותחו בהצלחה {success_count} מתוך {count} מיילים")
            print(f"{'='*50}\n")
            
            return success_count > 0
            
        except Exception as e:
            self.log_message(f"שגיאה בניתוח המיילים: {e}")
            return False
    
    def show_stats(self):
        """הצגת סטטיסטיקות"""
        try:
            response = requests.get(f"{self.server_url}/api/stats", timeout=5)
            if response.status_code == 200:
                stats = response.json()
                print(f"\n{'='*50}")
                print(f"סטטיסטיקות ניתוח:")
                print(f"{'='*50}")
                print(f"מיילים נותחים: {stats.get('total_emails', 0)}")
                print(f"פגישות נותחות: {stats.get('total_meetings', 0)}")
                print(f"ניתוחים היום: {stats.get('today_analyses', 0)}")
                print(f"{'='*50}\n")
            else:
                print("לא ניתן לקבל סטטיסטיקות מהשרת")
        except Exception as e:
            self.log_message(f"שגיאה בקבלת סטטיסטיקות: {e}")
            print("שגיאה בקבלת סטטיסטיקות")

def main():
    """פונקציה ראשית"""
    print("="*60)
    print("AI Email Manager - תוסף Outlook עצמאי")
    print("="*60)
    
    # יצירת התוסף
    addin = StandaloneOutlookAddin()
    
    # חיבור ל-Outlook
    if not addin.connect_to_outlook():
        print("לא ניתן להתחבר ל-Outlook. ודא ש-Outlook פתוח.")
        return
    
    # בדיקת חיבור לשרת
    if not addin.test_server_connection():
        print("השרת לא זמין. ודא שהשרת רץ על localhost:5000")
        print("הפעל: python app_with_ai.py")
        return
    
    print("\nאפשרויות:")
    print("1. ניתוח המייל הנוכחי")
    print("2. ניתוח כל המיילים הנבחרים")
    print("3. הצגת סטטיסטיקות")
    print("4. יציאה")
    
    while True:
        try:
            choice = input("\nבחר אפשרות (1-4): ").strip()
            
            if choice == "1":
                addin.analyze_current_email()
            elif choice == "2":
                addin.analyze_selected_emails()
            elif choice == "3":
                addin.show_stats()
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


