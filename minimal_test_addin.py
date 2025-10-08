# -*- coding: utf-8 -*-
"""
תוסף COM מינימלי לבדיקה
"""

import win32com.server.register
import win32com.server.util
import win32com.client
import os
import sys
import traceback

# הגדרת CLSID ייחודי
_reg_clsid_ = "{12345678-1234-1234-1234-123456789012}"

class MinimalTestAddin:
    """תוסף COM מינימלי לבדיקה"""
    
    _reg_clsid_ = _reg_clsid_
    _reg_progid_ = "MinimalTestAddin.Addin"
    _reg_desc_ = "Minimal Test Addin"
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    
    def __init__(self):
        """אתחול התוסף"""
        try:
            # יצירת קובץ טסט
            test_file = os.path.join(os.environ['TEMP'], 'minimal_test_addin_loaded.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Minimal Test Addin Loaded Successfully!")
            print(f"Minimal Test Addin initialized - test file created: {test_file}")
        except Exception as e:
            print(f"Error in MinimalTestAddin.__init__: {e}")
            traceback.print_exc()
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """חיבור ל-Outlook"""
        try:
            print("MinimalTestAddin.OnConnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'minimal_test_addin_connected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Minimal Test Addin Connected Successfully!")
            print(f"Minimal Test Addin connected - test file created: {test_file}")
            return True
        except Exception as e:
            print(f"Error in MinimalTestAddin.OnConnection: {e}")
            traceback.print_exc()
            return False
    
    def OnDisconnection(self, remove_mode):
        """ניתוק מ-Outlook"""
        try:
            print("MinimalTestAddin.OnDisconnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'minimal_test_addin_disconnected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Minimal Test Addin Disconnected Successfully!")
            print(f"Minimal Test Addin disconnected - test file created: {test_file}")
        except Exception as e:
            print(f"Error in MinimalTestAddin.OnDisconnection: {e}")
            traceback.print_exc()
    
    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            print("MinimalTestAddin.OnStartupComplete called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'minimal_test_addin_startup_complete.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Minimal Test Addin Startup Complete Successfully!")
            print(f"Minimal Test Addin startup complete - test file created: {test_file}")
        except Exception as e:
            print(f"Error in MinimalTestAddin.OnStartupComplete: {e}")
            traceback.print_exc()
    
    def OnBeginShutdown(self, custom):
        """התחלת סגירת Outlook"""
        try:
            print("MinimalTestAddin.OnBeginShutdown called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'minimal_test_addin_shutdown.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("Minimal Test Addin Shutdown Successfully!")
            print(f"Minimal Test Addin shutdown - test file created: {test_file}")
        except Exception as e:
            print(f"Error in MinimalTestAddin.OnBeginShutdown: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    print("Registering Minimal Test Addin...")
    win32com.server.register.UseCommandLine(MinimalTestAddin)
    print("Minimal Test Addin registered successfully!")
