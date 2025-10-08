# -*- coding: utf-8 -*-
"""
תוסף COM עם קובץ VBS
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
_reg_clsid_ = "{77777777-6666-5555-4444-333333333333}"

class VBSAddin:
    """תוסף COM עם קובץ VBS"""
    
    _reg_clsid_ = _reg_clsid_
    _reg_progid_ = "VBSAddin.Addin"
    _reg_desc_ = "VBS Addin"
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_threading_ = "Apartment"
    
    def __init__(self):
        """אתחול התוסף"""
        try:
            # יצירת קובץ טסט
            test_file = os.path.join(os.environ['TEMP'], 'vbs_addin_loaded.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("VBS Addin Loaded Successfully!")
            print(f"VBS Addin initialized - test file created: {test_file}")
        except Exception as e:
            print(f"Error in VBSAddin.__init__: {e}")
            traceback.print_exc()
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """חיבור ל-Outlook"""
        try:
            print("VBSAddin.OnConnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'vbs_addin_connected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("VBS Addin Connected Successfully!")
            print(f"VBS Addin connected - test file created: {test_file}")
            return True
        except Exception as e:
            print(f"Error in VBSAddin.OnConnection: {e}")
            traceback.print_exc()
            return False
    
    def OnDisconnection(self, remove_mode):
        """ניתוק מ-Outlook"""
        try:
            print("VBSAddin.OnDisconnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'vbs_addin_disconnected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("VBS Addin Disconnected Successfully!")
            print(f"VBS Addin disconnected - test file created: {test_file}")
        except Exception as e:
            print(f"Error in VBSAddin.OnDisconnection: {e}")
            traceback.print_exc()
    
    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            print("VBSAddin.OnStartupComplete called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'vbs_addin_startup_complete.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("VBS Addin Startup Complete Successfully!")
            print(f"VBS Addin startup complete - test file created: {test_file}")
        except Exception as e:
            print(f"Error in VBSAddin.OnStartupComplete: {e}")
            traceback.print_exc()
    
    def OnBeginShutdown(self, custom):
        """התחלת סגירת Outlook"""
        try:
            print("VBSAddin.OnBeginShutdown called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'vbs_addin_shutdown.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("VBS Addin Shutdown Successfully!")
            print(f"VBS Addin shutdown - test file created: {test_file}")
        except Exception as e:
            print(f"Error in VBSAddin.OnBeginShutdown: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    print("Registering VBS Addin...")
    win32com.server.register.UseCommandLine(VBSAddin)
    print("VBS Addin registered successfully!")


