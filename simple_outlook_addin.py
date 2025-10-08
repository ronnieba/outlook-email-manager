# -*- coding: utf-8 -*-
"""
AI Email Manager - תוסף Outlook פשוט ועובד
גרסה מינימלית שתעבוד בוודאות
"""

import win32com.client
from win32com.client import constants
import pythoncom
import os
import sys
import logging

# הגדרת לוגים פשוטה
LOG_FILE = os.path.join(os.environ.get('TEMP', os.getcwd()), 'simple_addin.log')
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    encoding='utf-8'
)

def log_info(message):
    logging.info(message)
    print(f"INFO: {message}")

def log_error(message):
    logging.error(message)
    print(f"ERROR: {message}")

class SimpleOutlookAddin:
    """תוסף Outlook פשוט ועובד"""
    
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_clsid_ = "{12345678-1234-1234-1234-123456789012}"
    _reg_progid_ = "SimpleAIEmailManager.Addin"
    _reg_desc_ = "Simple AI Email Manager"
    _reg_ver_ = "1.0"
    _reg_threading_ = "Apartment"

    def __init__(self):
        log_info("Simple Add-in initialized")
        self.Application = None

    def OnConnection(self, application, connectMode, addin, custom):
        """חיבור ל-Outlook"""
        try:
            log_info("OnConnection called")
            self.Application = application
            log_info("Successfully connected to Outlook")
        except Exception as e:
            log_error(f"Error in OnConnection: {e}")

    def OnDisconnection(self, removeMode, custom):
        """ניתוק מ-Outlook"""
        log_info("OnDisconnection called")
        self.Application = None

    def OnStartupComplete(self, custom):
        """השלמת אתחול"""
        log_info("OnStartupComplete called")

    def OnBeginShutdown(self, custom):
        """תחילת סגירה"""
        log_info("OnBeginShutdown called")


def RegisterAddin(klass):
    """רישום התוסף"""
    import win32com.server.register
    win32com.server.register.UseCommandLine(klass)

def UnregisterAddin(klass):
    """ביטול רישום התוסף"""
    import win32com.server.register
    win32com.server.register.UseCommandLine(klass, unregister=True)


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) > 1:
        if sys.argv[1] == '--register':
            try:
                RegisterAddin(SimpleOutlookAddin)
                print("התוסף הפשוט נרשם בהצלחה!")
            except Exception as e:
                print(f"שגיאה ברישום: {e}")
        elif sys.argv[1] == '--unregister':
            try:
                UnregisterAddin(SimpleOutlookAddin)
                print("התוסף הפשוט בוטל בהצלחה!")
            except Exception as e:
                print(f"שגיאה בביטול רישום: {e}")
        else:
            print("שימוש: python simple_outlook_addin.py --register או --unregister")
    else:
        print("תוסף Outlook פשוט")
        print("שימוש: python simple_outlook_addin.py --register או --unregister")

log_info("Simple Add-in script finished")


