# -*- coding: utf-8 -*-
"""
תוסף COM עם קובץ VBS נפרד
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
_reg_clsid_ = "{99999999-9999-9999-9999-999999999999}"

class VBSAddin3:
    """תוסף COM עם קובץ VBS נפרד"""
    
    _reg_clsid_ = _reg_clsid_
    _reg_progid_ = "VBSAddin3.Addin"
    _reg_desc_ = "VBS Addin 3"
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_threading_ = "Apartment"
    
    def __init__(self):
        """אתחול התוסף"""
        try:
            # יצירת קובץ טסט
            test_file = os.path.join(os.environ['TEMP'], 'vbs_addin3_loaded.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("VBS Addin 3 Loaded Successfully!")
            print(f"VBS Addin 3 initialized - test file created: {test_file}")
        except Exception as e:
            print(f"Error in VBSAddin3.__init__: {e}")
            traceback.print_exc()
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """חיבור ל-Outlook"""
        try:
            print("VBSAddin3.OnConnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'vbs_addin3_connected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("VBS Addin 3 Connected Successfully!")
            print(f"VBS Addin 3 connected - test file created: {test_file}")
            return True
        except Exception as e:
            print(f"Error in VBSAddin3.OnConnection: {e}")
            traceback.print_exc()
            return False
    
    def OnDisconnection(self, remove_mode):
        """ניתוק מ-Outlook"""
        try:
            print("VBSAddin3.OnDisconnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'vbs_addin3_disconnected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("VBS Addin 3 Disconnected Successfully!")
            print(f"VBS Addin 3 disconnected - test file created: {test_file}")
        except Exception as e:
            print(f"Error in VBSAddin3.OnDisconnection: {e}")
            traceback.print_exc()
    
    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            print("VBSAddin3.OnStartupComplete called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'vbs_addin3_startup_complete.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("VBS Addin 3 Startup Complete Successfully!")
            print(f"VBS Addin 3 startup complete - test file created: {test_file}")
        except Exception as e:
            print(f"Error in VBSAddin3.OnStartupComplete: {e}")
            traceback.print_exc()
    
    def OnBeginShutdown(self, custom):
        """התחלת סגירת Outlook"""
        try:
            print("VBSAddin3.OnBeginShutdown called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'vbs_addin3_shutdown.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("VBS Addin 3 Shutdown Successfully!")
            print(f"VBS Addin 3 shutdown - test file created: {test_file}")
        except Exception as e:
            print(f"Error in VBSAddin3.OnBeginShutdown: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    print("Registering VBS Addin 3...")
    win32com.server.register.UseCommandLine(VBSAddin3)
    print("VBS Addin 3 registered successfully!")


