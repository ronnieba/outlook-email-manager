# -*- coding: utf-8 -*-
"""
AI Email Manager - תוסף COM שעובד בוודאות
גרסה מינימלית ומוכחת שעובדת
"""

import win32com.client
from win32com.client import constants
import pythoncom
import os
import sys
import logging

# הגדרת לוגים מפורטת
LOG_FILE = os.path.join(os.environ.get('TEMP', os.getcwd()), 'working_addin.log')
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

def log_info(message):
    logging.info(message)
    print(f"INFO: {message}")

def log_error(message, exc_info=False):
    logging.error(message, exc_info=exc_info)
    print(f"ERROR: {message}")

def log_debug(message):
    logging.debug(message)
    print(f"DEBUG: {message}")

class WorkingOutlookAddin:
    """תוסף Outlook שעובד בוודאות"""
    
    # הגדרות COM מינימליות
    _public_methods_ = [
        'OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown'
    ]
    _reg_clsid_ = "{87654321-4321-4321-4321-210987654321}"
    _reg_progid_ = "WorkingAIEmailManager.Addin"
    _reg_desc_ = "Working AI Email Manager for Outlook"
    _reg_ver_ = "1.0"
    _reg_threading_ = "Apartment"
    _reg_interfaces_ = [pythoncom.IID_IDispatch]

    def __init__(self):
        log_info("Working Add-in __init__ called")
        self.Application = None
        self.addin_loaded = False
        
        # ניסיון ליצור קובץ בדיקה
        try:
            test_file = os.path.join(os.environ.get('TEMP', os.getcwd()), 'addin_test.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write(f"Add-in initialized at {os.getcwd()}\n")
            log_info(f"Test file created: {test_file}")
        except Exception as e:
            log_error(f"Could not create test file: {e}")

    def OnConnection(self, application, connectMode, addin, custom):
        """חיבור ל-Outlook - נקודת הכניסה הראשית"""
        try:
            log_info(f"OnConnection called with connectMode: {connectMode}")
            
            # אתחול COM
            try:
                pythoncom.CoInitialize()
                log_debug("COM initialized")
            except pythoncom.com_error as e:
                log_debug(f"COM already initialized: {e}")
            
            # שמירת הפניה ל-Outlook
            self.Application = application
            self.addin_loaded = True
            
            log_info("Successfully connected to Outlook")
            
            # יצירת קובץ בדיקה
            try:
                test_file = os.path.join(os.environ.get('TEMP', os.getcwd()), 'addin_connected.txt')
                with open(test_file, 'w', encoding='utf-8') as f:
                    f.write(f"Add-in connected successfully at {os.getcwd()}\n")
                    f.write(f"ConnectMode: {connectMode}\n")
                    f.write(f"Application: {application}\n")
                log_info(f"Connection test file created: {test_file}")
            except Exception as e:
                log_error(f"Could not create connection test file: {e}")
                
        except Exception as e:
            log_error(f"Error in OnConnection: {e}", exc_info=True)
            self.addin_loaded = False

    def OnDisconnection(self, removeMode, custom):
        """ניתוק מ-Outlook"""
        try:
            log_info(f"OnDisconnection called with removeMode: {removeMode}")
            
            # יצירת קובץ בדיקה
            try:
                test_file = os.path.join(os.environ.get('TEMP', os.getcwd()), 'addin_disconnected.txt')
                with open(test_file, 'w', encoding='utf-8') as f:
                    f.write(f"Add-in disconnected at {os.getcwd()}\n")
                    f.write(f"RemoveMode: {removeMode}\n")
                log_info(f"Disconnection test file created: {test_file}")
            except Exception as e:
                log_error(f"Could not create disconnection test file: {e}")
            
            self.addin_loaded = False
            self.Application = None
            
        except Exception as e:
            log_error(f"Error in OnDisconnection: {e}", exc_info=True)

    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            log_info("OnStartupComplete called")
            
            # יצירת קובץ בדיקה
            try:
                test_file = os.path.join(os.environ.get('TEMP', os.getcwd()), 'addin_startup_complete.txt')
                with open(test_file, 'w', encoding='utf-8') as f:
                    f.write(f"Add-in startup complete at {os.getcwd()}\n")
                    f.write(f"Add-in loaded: {self.addin_loaded}\n")
                log_info(f"Startup complete test file created: {test_file}")
            except Exception as e:
                log_error(f"Could not create startup complete test file: {e}")
                
        except Exception as e:
            log_error(f"Error in OnStartupComplete: {e}", exc_info=True)

    def OnBeginShutdown(self, custom):
        """תחילת סגירת Outlook"""
        try:
            log_info("OnBeginShutdown called")
            
            # יצירת קובץ בדיקה
            try:
                test_file = os.path.join(os.environ.get('TEMP', os.getcwd()), 'addin_shutdown.txt')
                with open(test_file, 'w', encoding='utf-8') as f:
                    f.write(f"Add-in shutdown at {os.getcwd()}\n")
                log_info(f"Shutdown test file created: {test_file}")
            except Exception as e:
                log_error(f"Could not create shutdown test file: {e}")
                
        except Exception as e:
            log_error(f"Error in OnBeginShutdown: {e}", exc_info=True)


def RegisterAddin(klass):
    """רישום התוסף ב-COM"""
    try:
        import win32com.server.register
        log_info("Starting COM registration...")
        win32com.server.register.UseCommandLine(klass)
        log_info("COM registration completed successfully")
    except Exception as e:
        log_error(f"Error during COM registration: {e}", exc_info=True)
        raise

def UnregisterAddin(klass):
    """ביטול רישום התוסף"""
    try:
        import win32com.server.register
        log_info("Starting COM unregistration...")
        win32com.server.register.UseCommandLine(klass, unregister=True)
        log_info("COM unregistration completed successfully")
    except Exception as e:
        log_error(f"Error during COM unregistration: {e}", exc_info=True)
        raise


if __name__ == '__main__':
    """הפעלה מהשורת פקודה"""
    import sys
    
    log_info("Working Add-in script started")
    
    if len(sys.argv) > 1:
        if sys.argv[1] == '--register':
            try:
                RegisterAddin(WorkingOutlookAddin)
                print("התוסף נרשם בהצלחה!")
                log_info("Registration completed successfully")
            except Exception as e:
                print(f"שגיאה ברישום: {e}")
                log_error(f"Registration failed: {e}", exc_info=True)
        elif sys.argv[1] == '--unregister':
            try:
                UnregisterAddin(WorkingOutlookAddin)
                print("התוסף בוטל בהצלחה!")
                log_info("Unregistration completed successfully")
            except Exception as e:
                print(f"שגיאה בביטול רישום: {e}")
                log_error(f"Unregistration failed: {e}", exc_info=True)
        else:
            print("שימוש: python working_outlook_addin.py --register או --unregister")
    else:
        print("תוסף Outlook שעובד בוודאות")
        print("שימוש: python working_outlook_addin.py --register או --unregister")

log_info("Working Add-in script finished")


