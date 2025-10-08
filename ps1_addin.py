# -*- coding: utf-8 -*-
"""
תוסף COM עם קובץ PS1
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
_reg_clsid_ = "{55555555-4444-3333-2222-111111111111}"

class PS1Addin:
    """תוסף COM עם קובץ PS1"""
    
    _reg_clsid_ = _reg_clsid_
    _reg_progid_ = "PS1Addin.Addin"
    _reg_desc_ = "PS1 Addin"
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_threading_ = "Apartment"
    
    def __init__(self):
        """אתחול התוסף"""
        try:
            # יצירת קובץ טסט
            test_file = os.path.join(os.environ['TEMP'], 'ps1_addin_loaded.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("PS1 Addin Loaded Successfully!")
            print(f"PS1 Addin initialized - test file created: {test_file}")
        except Exception as e:
            print(f"Error in PS1Addin.__init__: {e}")
            traceback.print_exc()
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """חיבור ל-Outlook"""
        try:
            print("PS1Addin.OnConnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'ps1_addin_connected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("PS1 Addin Connected Successfully!")
            print(f"PS1 Addin connected - test file created: {test_file}")
            return True
        except Exception as e:
            print(f"Error in PS1Addin.OnConnection: {e}")
            traceback.print_exc()
            return False
    
    def OnDisconnection(self, remove_mode):
        """ניתוק מ-Outlook"""
        try:
            print("PS1Addin.OnDisconnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'ps1_addin_disconnected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("PS1 Addin Disconnected Successfully!")
            print(f"PS1 Addin disconnected - test file created: {test_file}")
        except Exception as e:
            print(f"Error in PS1Addin.OnDisconnection: {e}")
            traceback.print_exc()
    
    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            print("PS1Addin.OnStartupComplete called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'ps1_addin_startup_complete.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("PS1 Addin Startup Complete Successfully!")
            print(f"PS1 Addin startup complete - test file created: {test_file}")
        except Exception as e:
            print(f"Error in PS1Addin.OnStartupComplete: {e}")
            traceback.print_exc()
    
    def OnBeginShutdown(self, custom):
        """התחלת סגירת Outlook"""
        try:
            print("PS1Addin.OnBeginShutdown called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'ps1_addin_shutdown.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("PS1 Addin Shutdown Successfully!")
            print(f"PS1 Addin shutdown - test file created: {test_file}")
        except Exception as e:
            print(f"Error in PS1Addin.OnBeginShutdown: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    print("Registering PS1 Addin...")
    win32com.server.register.UseCommandLine(PS1Addin)
    print("PS1 Addin registered successfully!")


