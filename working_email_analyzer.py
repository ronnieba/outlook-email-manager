# -*- coding: utf-8 -*-
"""
AI Email Manager - מנתח מיילים שעובד
גרסה שעובדת עם Outlook
"""

import win32com.client
import requests
import json
from datetime import datetime
import os
import time

def analyze_current_email():
    """ניתוח המייל הנוכחי"""
    try:
        print("מתחבר ל-Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        print("בודק חלון פעיל...")
        active_explorer = outlook.ActiveExplorer()
        
        if not active_explorer:
            print("אין חלון Outlook פעיל. אנא פתח חלון Outlook ונסה שוב.")
            return False
        
        print("בודק מיילים נבחרים...")
        selection = active_explorer.Selection
        
        if selection.Count == 0:
            print("לא נבחר מייל. אנא בחר מייל ב-Outlook ונסה שוב.")
            return False
        
        mail_item = selection[0]
        subject = mail_item.Subject or "ללא נושא"
        
        print(f"מנתח מייל: {subject}")
        
        # הכנת הנתונים לניתוח
        email_data = {
            'subject': subject,
            'sender': mail_item.SenderName or '',
            'body': mail_item.Body or '',
            'sender_email': mail_item.SenderEmailAddress or '',
            'received_time': mail_item.ReceivedTime.isoformat() if hasattr(mail_item, 'ReceivedTime') else '',
            'has_attachments': mail_item.Attachments.Count > 0
        }
        
        # שליחה לשרת
        print("שולח לניתוח AI...")
        response = requests.post(
            "http://localhost:5000/api/outlook-addin/analyze-email",
            json=email_data,
            timeout=30
        )
        
        if response.status_code == 200:
            analysis = response.json()
            
            if analysis.get("success"):
                # הוספת הניתוח למייל
                importance_percent = int(analysis.get('importance_score', 0) * 100)
                
                # ציון חשיבות
                try:
                    mail_item.UserProperties.Add("AI_Score", 1, True)
                except:
                    pass
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
                
                # עדכון קטגוריה של Outlook
                category_name = f"AI: {importance_percent}%"
                try:
                    # הוסף קטגוריה (אם יש כבר קטגוריות, נוסיף אליהן)
                    existing_categories = mail_item.Categories
                    if existing_categories:
                        # מחק קטגוריות AI קיימות
                        categories_list = [cat.strip() for cat in existing_categories.split(',') if not cat.strip().startswith('AI:')]
                        categories_list.append(category_name)
                        mail_item.Categories = ', '.join(categories_list)
                    else:
                        mail_item.Categories = category_name
                except Exception as e:
                    print(f"שגיאה בעדכון קטגוריה: {e}")
                
                # עדכון PRIORITYNUM (שדה מספרי)
                try:
                    # נסה למצוא את השדה, אם לא קיים - צור אותו
                    priority_field = None
                    try:
                        priority_field = mail_item.UserProperties("PRIORITYNUM")
                    except:
                        # צור שדה חדש מסוג מספר (olNumber = 3)
                        priority_field = mail_item.UserProperties.Add("PRIORITYNUM", 3, True)
                    
                    # עדכן את הערך
                    priority_field.Value = importance_percent
                except Exception as e:
                    print(f"שגיאה בעדכון PRIORITYNUM: {e}")
                
                # שמירה
                mail_item.Save()
                
                # הצגת התוצאות
                score = importance_percent
                category = analysis.get('category', 'לא סווג')
                summary = analysis.get('summary', 'לא נמצא סיכום')
                
                print(f"\n{'='*50}")
                print(f"ניתוח הושלם בהצלחה!")
                print(f"{'='*50}")
                print(f"ציון חשיבות: {score}%")
                print(f"קטגוריה: {category}")
                print(f"סיכום: {summary}")
                print(f"{'='*50}\n")
                
                print("הניתוח נוסף למייל בהצלחה!")
                print("ניתן לראות את התוצאות ב-Custom Properties של המייל.")
                
                return True
            else:
                error_msg = analysis.get('error', 'שגיאה לא ידועה')
                print(f"שגיאה בניתוח: {error_msg}")
                return False
        else:
            print(f"שגיאת שרת: {response.status_code}")
            return False
            
    except requests.exceptions.RequestException as e:
        print(f"שגיאת רשת: {e}")
        print("ודא שהשרת רץ: python app_with_ai.py")
        return False
    except Exception as e:
        print(f"שגיאה: {e}")
        return False

def main():
    """פונקציה ראשית"""
    print("="*60)
    print("AI Email Manager - מנתח מיילים שעובד")
    print("="*60)
    print()
    print("הוראות:")
    print("1. ודא ש-Outlook פתוח")
    print("2. ודא שהשרת רץ: python app_with_ai.py")
    print("3. פתח חלון Outlook (למשל Inbox)")
    print("4. בחר מייל ברשימה")
    print("5. המתן 3 שניות להתחלת הניתוח...")
    print()
    
    # המתנה של 3 שניות
    for i in range(3, 0, -1):
        print(f"מתחיל בעוד {i} שניות...")
        time.sleep(1)
    
    print("\nמתחיל ניתוח...")
    
    success = analyze_current_email()
    
    if success:
        print("\nהניתוח הושלם בהצלחה!")
        print("המייל עודכן עם הניתוח AI")
    else:
        print("\nהניתוח נכשל")
        print("בדוק את ההוראות למעלה")
    
    print("\nהמתן 5 שניות לסגירה...")
    time.sleep(5)

if __name__ == "__main__":
    main()














