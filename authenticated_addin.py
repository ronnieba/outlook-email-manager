# -*- coding: utf-8 -*-
"""
תוסף COM עם אימות מלא
"""

import win32com.server.register
import win32com.server.util
import win32com.client
import os
import sys
import traceback
import pythoncom

# הגדרת CLSID ייחודי
_reg_clsid_ = "{87654321-4321-4321-4321-210987654321}"

class AuthenticatedAddin:
    """תוסף COM עם אימות מלא"""
    
    _reg_clsid_ = _reg_clsid_
    _reg_progid_ = "AuthenticatedAddin.Addin"
    _reg_desc_ = "Authenticated Addin"
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_threading_ = "Apartment"
    
    def __init__(self):
        """אתחול התוסף"""
        try:
            # יצירת קובץ טסט
            test_file = os.path.join(os.environ['TEMP'], 'authenticated_addin_loaded.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Authenticated Addin Loaded Successfully!")
            print(f"Authenticated Addin initialized - test file created: {test_file}")
        except Exception as e:
            print(f"Error in AuthenticatedAddin.__init__: {e}")
            traceback.print_exc()
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """חיבור ל-Outlook"""
        try:
            print("AuthenticatedAddin.OnConnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'authenticated_addin_connected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Authenticated Addin Connected Successfully!")
            print(f"Authenticated Addin connected - test file created: {test_file}")
            return True
        except Exception as e:
            print(f"Error in AuthenticatedAddin.OnConnection: {e}")
            traceback.print_exc()
            return False
    
    def OnDisconnection(self, remove_mode):
        """ניתוק מ-Outlook"""
        try:
            print("AuthenticatedAddin.OnDisconnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'authenticated_addin_disconnected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Authenticated Addin Disconnected Successfully!")
            print(f"Authenticated Addin disconnected - test file created: {test_file}")
        except Exception as e:
            print(f"Error in AuthenticatedAddin.OnDisconnection: {e}")
            traceback.print_exc()
    
    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            print("AuthenticatedAddin.OnStartupComplete called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'authenticated_addin_startup_complete.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Authenticated Addin Startup Complete Successfully!")
            print(f"Authenticated Addin startup complete - test file created: {test_file}")
        except Exception as e:
            print(f"Error in AuthenticatedAddin.OnStartupComplete: {e}")
            traceback.print_exc()
    
    def OnBeginShutdown(self, custom):
        """התחלת סגירת Outlook"""
        try:
            print("AuthenticatedAddin.OnBeginShutdown called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'authenticated_addin_shutdown.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Authenticated Addin Shutdown Successfully!")
            print(f"Authenticated Addin shutdown - test file created: {test_file}")
        except Exception as e:
            print(f"Error in AuthenticatedAddin.OnBeginShutdown: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    print("Registering Authenticated Addin...")
    win32com.server.register.UseCommandLine(AuthenticatedAddin)
    print("Authenticated Addin registered successfully!")


