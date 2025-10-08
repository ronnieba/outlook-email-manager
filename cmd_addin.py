# -*- coding: utf-8 -*-
"""
תוסף COM עם קובץ CMD
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
_reg_clsid_ = "{11111111-0000-0000-0000-000000000000}"

class CMDAddin:
    """תוסף COM עם קובץ CMD"""
    
    _reg_clsid_ = _reg_clsid_
    _reg_progid_ = "CMDAddin.Addin"
    _reg_desc_ = "CMD Addin"
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_threading_ = "Apartment"
    
    def __init__(self):
        """אתחול התוסף"""
        try:
            # יצירת קובץ טסט
            test_file = os.path.join(os.environ['TEMP'], 'cmd_addin_loaded.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("CMD Addin Loaded Successfully!")
            print(f"CMD Addin initialized - test file created: {test_file}")
        except Exception as e:
            print(f"Error in CMDAddin.__init__: {e}")
            traceback.print_exc()
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """חיבור ל-Outlook"""
        try:
            print("CMDAddin.OnConnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'cmd_addin_connected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("CMD Addin Connected Successfully!")
            print(f"CMD Addin connected - test file created: {test_file}")
            return True
        except Exception as e:
            print(f"Error in CMDAddin.OnConnection: {e}")
            traceback.print_exc()
            return False
    
    def OnDisconnection(self, remove_mode):
        """ניתוק מ-Outlook"""
        try:
            print("CMDAddin.OnDisconnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'cmd_addin_disconnected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("CMD Addin Disconnected Successfully!")
            print(f"CMD Addin disconnected - test file created: {test_file}")
        except Exception as e:
            print(f"Error in CMDAddin.OnDisconnection: {e}")
            traceback.print_exc()
    
    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            print("CMDAddin.OnStartupComplete called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'cmd_addin_startup_complete.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("CMD Addin Startup Complete Successfully!")
            print(f"CMD Addin startup complete - test file created: {test_file}")
        except Exception as e:
            print(f"Error in CMDAddin.OnStartupComplete: {e}")
            traceback.print_exc()
    
    def OnBeginShutdown(self, custom):
        """התחלת סגירת Outlook"""
        try:
            print("CMDAddin.OnBeginShutdown called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'cmd_addin_shutdown.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("CMD Addin Shutdown Successfully!")
            print(f"CMD Addin shutdown - test file created: {test_file}")
        except Exception as e:
            print(f"Error in CMDAddin.OnBeginShutdown: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    print("Registering CMD Addin...")
    win32com.server.register.UseCommandLine(CMDAddin)
    print("CMD Addin registered successfully!")


