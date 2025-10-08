# -*- coding: utf-8 -*-
"""
AI Email Manager - Outlook COM Add-in פשוט ועובד
גרסה פשוטה ומתקדמת של תוסף Outlook לניתוח מיילים עם AI
"""

import win32com.client
from win32com.client import constants
import pythoncom
import os
import sys
import logging
import requests
import json
from datetime import datetime

# הגדרת לוגים
LOG_FILE = os.path.join(os.environ.get('TEMP', os.getcwd()), 'ai_email_manager.log')
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

def log_info(message):
    """רישום הודעות מידע"""
    logging.info(message)
    print(f"INFO: {message}")

def log_error(message, exc_info=False):
    """רישום שגיאות"""
    logging.error(message, exc_info=exc_info)
    print(f"ERROR: {message}")

# הגדרות
SERVER_URL = "http://localhost:5000"

class AIEmailManagerAddin:
    """תוסף AI Email Manager ל-Outlook"""
    
    # הגדרות COM
    _public_methods_ = [
        'OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown',
        'GetCustomUI', 'OnRibbonLoad', 'OnAnalyzeEmailPress', 'OnOpenWebUIPress',
        'OnAnalyzeSelectedEmailsPress', 'OnShowStatsPress'
    ]
    _reg_clsid_ = "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"
    _reg_progid_ = "AIEmailManager.Addin"
    _reg_desc_ = "AI Email Manager for Outlook"
    _reg_ver_ = "2.0"
    _reg_threading_ = "Apartment"
    _reg_interfaces_ = [pythoncom.IID_IDispatch, '{000C0396-0000-0000-C000-000000000046}']

    def __init__(self):
        log_info("AI Email Manager Add-in initialized")
        self.Application = None
        self.addin_loaded = False
        self.ribbon = None

    def _co_initialize(self):
        """אתחול COM"""
        try:
            pythoncom.CoInitialize()
        except pythoncom.com_error:
            pass

    def OnConnection(self, application, connectMode, addin, custom):
        """חיבור ל-Outlook"""
        try:
            log_info(f"Connecting to Outlook. Mode: {connectMode}")
            self.Application = application
            self._co_initialize()
            self.addin_loaded = True
            log_info("Successfully connected to Outlook")
        except Exception as e:
            log_error(f"Error in OnConnection: {e}", exc_info=True)
            self.addin_loaded = False

    def OnDisconnection(self, removeMode, custom):
        """ניתוק מ-Outlook"""
        self._co_initialize()
        log_info(f"Disconnecting from Outlook. Mode: {removeMode}")
        self.addin_loaded = False
        self.Application = None

    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        self._co_initialize()
        log_info("Outlook startup complete")
        if self.addin_loaded:
            log_info("Add-in ready for use")

    def OnBeginShutdown(self, custom):
        """תחילת סגירת Outlook"""
        self._co_initialize()
        log_info("Outlook shutdown beginning")

    def GetCustomUI(self, RibbonID):
        """החזרת XML של ה-Ribbon"""
        self._co_initialize()
        log_info(f"GetCustomUI called for: {RibbonID}")
        
        ribbon_xml = """
        <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
          <ribbon>
            <tabs>
              <tab id="aiEmailManagerTab" label="AI Email Manager">
                <group id="aiAnalysisGroup" label="ניתוח AI">
                  <button id="analyzeEmailButton"
                          label="נתח מייל נוכחי"
                          size="large"
                          onAction="OnAnalyzeEmailPress"
                          imageMso="AnalyzeScenario"
                          screentip="נתח את המייל הנבחר עם AI"
                          supertip="מנתח את המייל הנבחר ומציג ציון חשיבות, קטגוריה וסיכום" />
                  <button id="analyzeSelectedEmailsButton"
                          label="נתח מיילים נבחרים"
                          size="large"
                          onAction="OnAnalyzeSelectedEmailsPress"
                          imageMso="AnalyzeScenario"
                          screentip="נתח את כל המיילים הנבחרים"
                          supertip="מנתח את כל המיילים הנבחרים בבת אחת" />
                </group>
                <group id="aiToolsGroup" label="כלים">
                  <button id="openWebUIButton"
                          label="פתח ממשק Web"
                          size="large"
                          onAction="OnOpenWebUIPress"
                          imageMso="World"
                          screentip="פתח את הממשק המקוון"
                          supertip="פותח את הממשק המקוון של AI Email Manager בדפדפן" />
                  <button id="showStatsButton"
                          label="הצג סטטיסטיקות"
                          size="large"
                          onAction="OnShowStatsPress"
                          imageMso="Chart"
                          screentip="הצג סטטיסטיקות ניתוח"
                          supertip="מציג סטטיסטיקות על הניתוחים שבוצעו" />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>
        """
        
        return ribbon_xml

    def OnRibbonLoad(self, ribbon):
        """טעינת ה-Ribbon"""
        self._co_initialize()
        log_info("Ribbon loaded successfully")
        self.ribbon = ribbon

    def OnAnalyzeEmailPress(self, control):
        """ניתוח מייל נוכחי"""
        log_info("Analyze Email button pressed")
        self._co_initialize()
        
        try:
            if not self.Application:
                self._show_message("שגיאה: לא ניתן לגשת ל-Outlook", "שגיאה")
                return

            # קבלת המייל הנבחר
            active_explorer = self.Application.ActiveExplorer()
            if not active_explorer or active_explorer.Selection.Count == 0:
                self._show_message("אנא בחר מייל לניתוח", "הודעה")
                return

            selected_item = active_explorer.Selection.Item(1)
            
            if selected_item.Class != constants.olMail:
                self._show_message("הפריט שנבחר אינו מייל", "הודעה")
                return

            # ניתוח המייל
            self._analyze_single_email(selected_item)
            
        except Exception as e:
            log_error(f"Error in OnAnalyzeEmailPress: {e}", exc_info=True)
            self._show_message(f"שגיאה בניתוח המייל: {e}", "שגיאה")

    def OnAnalyzeSelectedEmailsPress(self, control):
        """ניתוח מיילים נבחרים"""
        log_info("Analyze Selected Emails button pressed")
        self._co_initialize()
        
        try:
            if not self.Application:
                self._show_message("שגיאה: לא ניתן לגשת ל-Outlook", "שגיאה")
                return

            active_explorer = self.Application.ActiveExplorer()
            if not active_explorer or active_explorer.Selection.Count == 0:
                self._show_message("אנא בחר מיילים לניתוח", "הודעה")
                return

            selection = active_explorer.Selection
            count = selection.Count
            
            self._show_message(f"מנתח {count} מיילים...", "מידע")
            
            success_count = 0
            for i in range(count):
                try:
                    mail_item = selection.Item(i + 1)
                    if mail_item.Class == constants.olMail:
                        if self._analyze_single_email(mail_item, show_message=False):
                            success_count += 1
                except Exception as e:
                    log_error(f"Error analyzing email {i+1}: {e}")
                    continue
            
            self._show_message(f"נותחו בהצלחה {success_count} מתוך {count} מיילים", "תוצאות")
            
        except Exception as e:
            log_error(f"Error in OnAnalyzeSelectedEmailsPress: {e}", exc_info=True)
            self._show_message(f"שגיאה בניתוח המיילים: {e}", "שגיאה")

    def OnOpenWebUIPress(self, control):
        """פתיחת ממשק Web"""
        log_info("Open Web UI button pressed")
        self._co_initialize()
        
        try:
            import webbrowser
            webbrowser.open(SERVER_URL)
            log_info("Web UI opened successfully")
        except Exception as e:
            log_error(f"Error opening Web UI: {e}", exc_info=True)
            self._show_message(f"לא ניתן לפתוח את הממשק המקוון: {e}", "שגיאה")

    def OnShowStatsPress(self, control):
        """הצגת סטטיסטיקות"""
        log_info("Show Stats button pressed")
        self._co_initialize()
        
        try:
            # ניסיון לקבל סטטיסטיקות מהשרת
            response = requests.get(f"{SERVER_URL}/api/stats", timeout=5)
            if response.status_code == 200:
                stats = response.json()
                message = f"סטטיסטיקות ניתוח:\n\n"
                message += f"מיילים נותחים: {stats.get('total_emails', 0)}\n"
                message += f"פגישות נותחות: {stats.get('total_meetings', 0)}\n"
                message += f"ניתוחים היום: {stats.get('today_analyses', 0)}\n"
                self._show_message(message, "סטטיסטיקות")
            else:
                self._show_message("לא ניתן לקבל סטטיסטיקות מהשרת", "הודעה")
        except Exception as e:
            log_error(f"Error getting stats: {e}", exc_info=True)
            self._show_message("שגיאה בקבלת סטטיסטיקות", "שגיאה")

    def _analyze_single_email(self, mail_item, show_message=True):
        """ניתוח מייל בודד"""
        try:
            subject = mail_item.Subject or "ללא נושא"
            log_info(f"Analyzing email: {subject}")
            
            # הכנת נתוני המייל
            email_data = {
                "subject": subject,
                "body": mail_item.Body or "",
                "sender_name": mail_item.SenderName or "",
                "sender": mail_item.SenderEmailAddress or "",
                "date": mail_item.ReceivedTime.isoformat() if hasattr(mail_item, 'ReceivedTime') else "",
                "has_attachments": mail_item.Attachments.Count > 0
            }
            
            # שליחה לשרת
            response = requests.post(
                f"{SERVER_URL}/api/outlook-addin/analyze-email",
                json=email_data,
                timeout=30
            )
            
            if response.status_code == 200:
                analysis = response.json()
                
                if analysis.get("success"):
                    # הוספת הניתוח למייל
                    self._add_analysis_to_email(mail_item, analysis)
                    
                    if show_message:
                        score = int(analysis.get('importance_score', 0) * 100)
                        category = analysis.get('category', 'לא סווג')
                        summary = analysis.get('summary', 'לא נמצא סיכום')
                        
                        message = f"ניתוח הושלם:\n\n"
                        message += f"📊 ציון חשיבות: {score}%\n"
                        message += f"🏷️ קטגוריה: {category}\n\n"
                        message += f"📝 סיכום:\n{summary}"
                        
                        self._show_message(message, "תוצאות ניתוח AI")
                    
                    return True
                else:
                    error_msg = analysis.get('error', 'שגיאה לא ידועה')
                    log_error(f"Server error: {error_msg}")
                    if show_message:
                        self._show_message(f"שגיאה בניתוח: {error_msg}", "שגיאה")
                    return False
            else:
                log_error(f"HTTP error: {response.status_code}")
                if show_message:
                    self._show_message(f"שגיאת שרת: {response.status_code}", "שגיאה")
                return False
                
        except requests.exceptions.RequestException as e:
            log_error(f"Network error: {e}", exc_info=True)
            if show_message:
                self._show_message(
                    f"שגיאת רשת בחיבור לשרת.\n\nודא שהשרת הראשי פועל בכתובת:\n{SERVER_URL}",
                    "שגיאת חיבור"
                )
            return False
        except Exception as e:
            log_error(f"Error analyzing email: {e}", exc_info=True)
            if show_message:
                self._show_message(f"שגיאה בניתוח המייל: {e}", "שגיאה")
            return False

    def _add_analysis_to_email(self, mail_item, analysis):
        """הוספת הניתוח למייל"""
        try:
            # הוספת Custom Properties
            importance_percent = int(analysis.get('importance_score', 0) * 100)
            
            # ציון חשיבות - מספר שלם (לתצוגה בעמודה)
            try:
                mail_item.UserProperties.Add("AISCORE", 3, True)  # 3 = olNumber
            except:
                pass  # אם כבר קיים
            
            try:
                mail_item.UserProperties("AISCORE").Value = importance_percent
            except Exception as e:
                log_error(f"Error setting AISCORE: {e}")
            
            # ציון חשיבות - טקסט (גיבוי)
            try:
                mail_item.UserProperties.Add("AI_Score", 1, True)  # 1 = Text
            except:
                pass
            
            mail_item.UserProperties("AI_Score").Value = f"{importance_percent}%"
            
            # קטגוריה
            try:
                mail_item.UserProperties.Add("AI_Category", 1, True)
            except:
                pass
            
            category = analysis.get('category', 'לא סווג')
            mail_item.UserProperties("AI_Category").Value = category
            
            # סיכום
            try:
                mail_item.UserProperties.Add("AI_Summary", 1, True)
            except:
                pass
            
            mail_item.UserProperties("AI_Summary").Value = analysis.get('summary', '')[:255]  # מוגבל ל-255 תווים
            
            # תאריך ניתוח
            try:
                mail_item.UserProperties.Add("AI_Analyzed", 1, True)
            except:
                pass
            
            mail_item.UserProperties("AI_Analyzed").Value = datetime.now().strftime("%Y-%m-%d %H:%M")
            
            # הוספת קטגוריה של Outlook לתצוגה ויזואלית
            try:
                category_name = f"AI: {importance_percent}%"
                current_categories = mail_item.Categories
                if current_categories:
                    if category_name not in current_categories:
                        mail_item.Categories = current_categories + ", " + category_name
                else:
                    mail_item.Categories = category_name
            except Exception as e:
                log_error(f"Error setting Outlook category: {e}")
            
            # הוספת דגל לפי חשיבות
            if importance_percent >= 80:
                mail_item.FlagRequest = "Follow up"
                mail_item.Importance = 2  # High importance
            elif importance_percent >= 60:
                mail_item.FlagRequest = "No Response Necessary"
                mail_item.Importance = 1  # Normal importance
            else:
                mail_item.Importance = 0  # Low importance
            
            # שמירה
            mail_item.Save()
            
            log_info(f"Analysis added to email: {mail_item.Subject} - Score: {importance_percent}%")
            
        except Exception as e:
            log_error(f"Error adding analysis to email: {e}", exc_info=True)

    def _show_message(self, text, title):
        """הצגת הודעה למשתמש"""
        try:
            if self.Application:
                self.Application.Session.MessageBox(text, constants.olOkOnly | constants.olInformation, title)
        except Exception as e:
            log_error(f"Error showing message: {e}")


def RegisterAddin(klass):
    """רישום התוסף ב-COM"""
    import win32com.server.register
    win32com.server.register.UseCommandLine(klass)

def UnregisterAddin(klass):
    """ביטול רישום התוסף"""
    import win32com.server.register
    win32com.server.register.UseCommandLine(klass, unregister=True)


if __name__ == '__main__':
    """הפעלה מהשורת פקודה"""
    import sys
    
    if len(sys.argv) > 1:
        if sys.argv[1] == '--register':
            try:
                RegisterAddin(AIEmailManagerAddin)
                print("התוסף נרשם בהצלחה!")
            except Exception as e:
                print(f"שגיאה ברישום: {e}")
        elif sys.argv[1] == '--unregister':
            try:
                UnregisterAddin(AIEmailManagerAddin)
                print("התוסף בוטל בהצלחה!")
            except Exception as e:
                print(f"שגיאה בביטול רישום: {e}")
        else:
            print("שימוש: python outlook_com_addin_final.py --register או --unregister")
    else:
        print("תוסף AI Email Manager ל-Outlook")
        print("שימוש: python outlook_com_addin_final.py --register או --unregister")

log_info("AI Email Manager Add-in script finished")