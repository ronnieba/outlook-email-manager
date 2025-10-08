# -*- coding: utf-8 -*-
"""
AI Email Manager - תוסף COM שעובד בוודאות
גרסה אולטרה-פשוטה ללא שגיאות
"""

import win32com.client
from win32com.client import constants
import pythoncom
import os
import sys

class UltraSimpleAddin:
    """תוסף Outlook אולטרה-פשוט"""
    
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_clsid_ = "{11111111-1111-1111-1111-111111111111}"
    _reg_progid_ = "UltraSimpleAddin.Addin"
    _reg_desc_ = "Ultra Simple Addin"
    _reg_ver_ = "1.0"
    _reg_threading_ = "Apartment"

    def __init__(self):
        # יצירת קובץ בדיקה פשוט
        try:
            with open(os.path.join(os.environ.get('TEMP', os.getcwd()), 'ultra_simple_init.txt'), 'w') as f:
                f.write("UltraSimpleAddin initialized\n")
        except:
            pass

    def OnConnection(self, application, connectMode, addin, custom):
        """חיבור ל-Outlook"""
        try:
            # יצירת קובץ בדיקה
            with open(os.path.join(os.environ.get('TEMP', os.getcwd()), 'ultra_simple_connected.txt'), 'w') as f:
                f.write("UltraSimpleAddin connected successfully\n")
                f.write(f"ConnectMode: {connectMode}\n")
        except:
            pass

    def OnDisconnection(self, removeMode, custom):
        """ניתוק מ-Outlook"""
        try:
            with open(os.path.join(os.environ.get('TEMP', os.getcwd()), 'ultra_simple_disconnected.txt'), 'w') as f:
                f.write("UltraSimpleAddin disconnected\n")
        except:
            pass

    def OnStartupComplete(self, custom):
        """השלמת אתחול"""
        try:
            with open(os.path.join(os.environ.get('TEMP', os.getcwd()), 'ultra_simple_startup.txt'), 'w') as f:
                f.write("UltraSimpleAddin startup complete\n")
        except:
            pass

    def OnBeginShutdown(self, custom):
        """תחילת סגירה"""
        try:
            with open(os.path.join(os.environ.get('TEMP', os.getcwd()), 'ultra_simple_shutdown.txt'), 'w') as f:
                f.write("UltraSimpleAddin shutdown\n")
        except:
            pass


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
                RegisterAddin(UltraSimpleAddin)
                print("UltraSimpleAddin registered successfully!")
            except Exception as e:
                print(f"Registration error: {e}")
        elif sys.argv[1] == '--unregister':
            try:
                UnregisterAddin(UltraSimpleAddin)
                print("UltraSimpleAddin unregistered successfully!")
            except Exception as e:
                print(f"Unregistration error: {e}")
        else:
            print("Usage: python ultra_simple_addin.py --register or --unregister")
    else:
        print("Ultra Simple Outlook Add-in")
        print("Usage: python ultra_simple_addin.py --register or --unregister")


