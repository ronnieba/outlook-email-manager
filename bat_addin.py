# -*- coding: utf-8 -*-
"""
תוסף COM עם קובץ BAT
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
_reg_clsid_ = "{33333333-2222-1111-0000-999999999999}"

class BATAddin:
    """תוסף COM עם קובץ BAT"""
    
    _reg_clsid_ = _reg_clsid_
    _reg_progid_ = "BATAddin.Addin"
    _reg_desc_ = "BAT Addin"
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_threading_ = "Apartment"
    
    def __init__(self):
        """אתחול התוסף"""
        try:
            # יצירת קובץ טסט
            test_file = os.path.join(os.environ['TEMP'], 'bat_addin_loaded.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("BAT Addin Loaded Successfully!")
            print(f"BAT Addin initialized - test file created: {test_file}")
        except Exception as e:
            print(f"Error in BATAddin.__init__: {e}")
            traceback.print_exc()
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """חיבור ל-Outlook"""
        try:
            print("BATAddin.OnConnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'bat_addin_connected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("BAT Addin Connected Successfully!")
            print(f"BAT Addin connected - test file created: {test_file}")
            return True
        except Exception as e:
            print(f"Error in BATAddin.OnConnection: {e}")
            traceback.print_exc()
            return False
    
    def OnDisconnection(self, remove_mode):
        """ניתוק מ-Outlook"""
        try:
            print("BATAddin.OnDisconnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'bat_addin_disconnected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("BAT Addin Disconnected Successfully!")
            print(f"BAT Addin disconnected - test file created: {test_file}")
        except Exception as e:
            print(f"Error in BATAddin.OnDisconnection: {e}")
            traceback.print_exc()
    
    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            print("BATAddin.OnStartupComplete called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'bat_addin_startup_complete.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("BAT Addin Startup Complete Successfully!")
            print(f"BAT Addin startup complete - test file created: {test_file}")
        except Exception as e:
            print(f"Error in BATAddin.OnStartupComplete: {e}")
            traceback.print_exc()
    
    def OnBeginShutdown(self, custom):
        """התחלת סגירת Outlook"""
        try:
            print("BATAddin.OnBeginShutdown called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'bat_addin_shutdown.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("BAT Addin Shutdown Successfully!")
            print(f"BAT Addin shutdown - test file created: {test_file}")
        except Exception as e:
            print(f"Error in BATAddin.OnBeginShutdown: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    print("Registering BAT Addin...")
    win32com.server.register.UseCommandLine(BATAddin)
    print("BAT Addin registered successfully!")


