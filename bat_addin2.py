# -*- coding: utf-8 -*-
"""
תוסף COM עם קובץ BAT נפרד
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
_reg_clsid_ = "{55555555-5555-5555-5555-555555555555}"

class BATAddin2:
    """תוסף COM עם קובץ BAT נפרד"""
    
    _reg_clsid_ = _reg_clsid_
    _reg_progid_ = "BATAddin2.Addin"
    _reg_desc_ = "BAT Addin 2"
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_threading_ = "Apartment"
    
    def __init__(self):
        """אתחול התוסף"""
        try:
            # יצירת קובץ טסט
            test_file = os.path.join(os.environ['TEMP'], 'bat_addin2_loaded.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("BAT Addin 2 Loaded Successfully!")
            print(f"BAT Addin 2 initialized - test file created: {test_file}")
        except Exception as e:
            print(f"Error in BATAddin2.__init__: {e}")
            traceback.print_exc()
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """חיבור ל-Outlook"""
        try:
            print("BATAddin2.OnConnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'bat_addin2_connected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("BAT Addin 2 Connected Successfully!")
            print(f"BAT Addin 2 connected - test file created: {test_file}")
            return True
        except Exception as e:
            print(f"Error in BATAddin2.OnConnection: {e}")
            traceback.print_exc()
            return False
    
    def OnDisconnection(self, remove_mode):
        """ניתוק מ-Outlook"""
        try:
            print("BATAddin2.OnDisconnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'bat_addin2_disconnected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("BAT Addin 2 Disconnected Successfully!")
            print(f"BAT Addin 2 disconnected - test file created: {test_file}")
        except Exception as e:
            print(f"Error in BATAddin2.OnDisconnection: {e}")
            traceback.print_exc()
    
    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            print("BATAddin2.OnStartupComplete called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'bat_addin2_startup_complete.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("BAT Addin 2 Startup Complete Successfully!")
            print(f"BAT Addin 2 startup complete - test file created: {test_file}")
        except Exception as e:
            print(f"Error in BATAddin2.OnStartupComplete: {e}")
            traceback.print_exc()
    
    def OnBeginShutdown(self, custom):
        """התחלת סגירת Outlook"""
        try:
            print("BATAddin2.OnBeginShutdown called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'bat_addin2_shutdown.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("BAT Addin 2 Shutdown Successfully!")
            print(f"BAT Addin 2 shutdown - test file created: {test_file}")
        except Exception as e:
            print(f"Error in BATAddin2.OnBeginShutdown: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    print("Registering BAT Addin 2...")
    win32com.server.register.UseCommandLine(BATAddin2)
    print("BAT Addin 2 registered successfully!")


