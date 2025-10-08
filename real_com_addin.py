# -*- coding: utf-8 -*-
"""
תוסף COM אמיתי עם קובץ EXE
"""

import win32com.server.register
import win32com.server.util
import win32com.client
import os
import sys
import traceback
import pythoncom
import win32api
import win32con

# הגדרת CLSID ייחודי
_reg_clsid_ = "{CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC}"

class RealCOMAddin:
    """תוסף COM אמיתי עם קובץ EXE"""
    
    _reg_clsid_ = _reg_clsid_
    _reg_progid_ = "RealCOMAddin.Addin"
    _reg_desc_ = "Real COM Addin"
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_threading_ = "Apartment"
    
    def __init__(self):
        """אתחול התוסף"""
        try:
            # יצירת קובץ טסט
            test_file = os.path.join(os.environ['TEMP'], 'real_com_addin_loaded.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Real COM Addin Loaded Successfully!")
            print(f"Real COM Addin initialized - test file created: {test_file}")
        except Exception as e:
            print(f"Error in RealCOMAddin.__init__: {e}")
            traceback.print_exc()
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """חיבור ל-Outlook"""
        try:
            print("RealCOMAddin.OnConnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'real_com_addin_connected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Real COM Addin Connected Successfully!")
            print(f"Real COM Addin connected - test file created: {test_file}")
            return True
        except Exception as e:
            print(f"Error in RealCOMAddin.OnConnection: {e}")
            traceback.print_exc()
            return False
    
    def OnDisconnection(self, remove_mode):
        """ניתוק מ-Outlook"""
        try:
            print("RealCOMAddin.OnDisconnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'real_com_addin_disconnected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Real COM Addin Disconnected Successfully!")
            print(f"Real COM Addin disconnected - test file created: {test_file}")
        except Exception as e:
            print(f"Error in RealCOMAddin.OnDisconnection: {e}")
            traceback.print_exc()
    
    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            print("RealCOMAddin.OnStartupComplete called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'real_com_addin_startup_complete.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Real COM Addin Startup Complete Successfully!")
            print(f"Real COM Addin startup complete - test file created: {test_file}")
        except Exception as e:
            print(f"Error in RealCOMAddin.OnStartupComplete: {e}")
            traceback.print_exc()
    
    def OnBeginShutdown(self, custom):
        """התחלת סגירת Outlook"""
        try:
            print("RealCOMAddin.OnBeginShutdown called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'real_com_addin_shutdown.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Real COM Addin Shutdown Successfully!")
            print(f"Real COM Addin shutdown - test file created: {test_file}")
        except Exception as e:
            print(f"Error in RealCOMAddin.OnBeginShutdown: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    print("Registering Real COM Addin...")
    win32com.server.register.UseCommandLine(RealCOMAddin)
    print("Real COM Addin registered successfully!")


