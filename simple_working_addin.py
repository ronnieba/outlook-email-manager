# -*- coding: utf-8 -*-
"""
AI Email Manager - תוסף COM פשוט שעובד בוודאות
גרסה מינימלית ללא שגיאות
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
LOG_FILE = os.path.join(os.environ.get('TEMP', os.getcwd()), 'simple_addin_working.log')
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

def log_info(message):
    logging.info(message)

def log_error(message, exc_info=False):
    logging.error(message, exc_info=exc_info)

# הגדרות
SERVER_URL = "http://localhost:5000"

class SimpleWorkingAddin:
    """תוסף פשוט שעובד בוודאות"""
    
    # הגדרות COM מינימליות
    _public_methods_ = [
        'OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown',
        'GetCustomUI', 'OnRibbonLoad', 'OnAnalyzeEmailPress'
    ]
    _reg_clsid_ = "{12345678-1234-1234-1234-123456789012}"
    _reg_progid_ = "SimpleWorkingAddin.Addin"
    _reg_desc_ = "Simple Working Addin"
    _reg_ver_ = "1.0"
    _reg_threading_ = "Apartment"
    _reg_interfaces_ = [pythoncom.IID_IDispatch, '{000C0396-0000-0000-C000-000000000046}']

    def __init__(self):
        log_info("Simple Working Add-in initialized")
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
            log_info(f"OnConnection called with connectMode: {connectMode}")
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
        log_info(f"OnDisconnection called with removeMode: {removeMode}")
        self.addin_loaded = False
        self.Application = None

    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        self._co_initialize()
        log_info("OnStartupComplete called")
        if self.addin_loaded:
            log_info("Add-in ready for use")

    def OnBeginShutdown(self, custom):
        """תחילת סגירת Outlook"""
        self._co_initialize()
        log_info("OnBeginShutdown called")

    def GetCustomUI(self, RibbonID):
        """החזרת XML של ה-Ribbon"""
        self._co_initialize()
        log_info(f"GetCustomUI called for: {RibbonID}")
        
        ribbon_xml = """
        <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
          <ribbon>
            <tabs>
              <tab id="simpleWorkingTab" label="AI Email Manager">
                <group id="simpleAnalysisGroup" label="ניתוח AI">
                  <button id="analyzeEmailButton"
                          label="נתח מייל נוכחי"
                          size="large"
                          onAction="OnAnalyzeEmailPress"
                          imageMso="AnalyzeScenario"
                          screentip="נתח את המייל הנבחר עם AI"
                          supertip="מנתח את המייל הנבחר ומציג ציון חשיבות, קטגוריה וסיכום" />
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

    def _analyze_single_email(self, mail_item):
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
                    
                    score = int(analysis.get('importance_score', 0) * 100)
                    category = analysis.get('category', 'לא סווג')
                    summary = analysis.get('summary', 'לא נמצא סיכום')
                    
                    message = f"ניתוח הושלם:\n\n"
                    message += f"ציון חשיבות: {score}%\n"
                    message += f"קטגוריה: {category}\n\n"
                    message += f"סיכום:\n{summary}"
                    
                    self._show_message(message, "תוצאות ניתוח AI")
                    
                    return True
                else:
                    error_msg = analysis.get('error', 'שגיאה לא ידועה')
                    log_error(f"Server error: {error_msg}")
                    self._show_message(f"שגיאה בניתוח: {error_msg}", "שגיאה")
                    return False
            else:
                log_error(f"HTTP error: {response.status_code}")
                self._show_message(f"שגיאת שרת: {response.status_code}", "שגיאה")
                return False
                
        except requests.exceptions.RequestException as e:
            log_error(f"Network error: {e}", exc_info=True)
            self._show_message(
                f"שגיאת רשת בחיבור לשרת.\n\nודא שהשרת הראשי פועל בכתובת:\n{SERVER_URL}",
                "שגיאת חיבור"
            )
            return False
        except Exception as e:
            log_error(f"Error analyzing email: {e}", exc_info=True)
            self._show_message(f"שגיאה בניתוח המייל: {e}", "שגיאה")
            return False

    def _add_analysis_to_email(self, mail_item, analysis):
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
            
            log_info(f"Analysis added to email: {mail_item.Subject}")
            
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
                RegisterAddin(SimpleWorkingAddin)
                print("התוסף הפשוט נרשם בהצלחה!")
            except Exception as e:
                print(f"שגיאה ברישום: {e}")
        elif sys.argv[1] == '--unregister':
            try:
                UnregisterAddin(SimpleWorkingAddin)
                print("התוסף הפשוט בוטל בהצלחה!")
            except Exception as e:
                print(f"שגיאה בביטול רישום: {e}")
        else:
            print("שימוש: python simple_working_addin.py --register או --unregister")
    else:
        print("תוסף פשוט שעובד בוודאות")
        print("שימוש: python simple_working_addin.py --register או --unregister")

log_info("Simple Working Add-in script finished")
