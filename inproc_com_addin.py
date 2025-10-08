
import sys
import os
import tempfile
import win32com.server.register
import win32com.server.util
import win32com.client
import traceback
import time

# הגדרת CLSID ייחודי
_reg_clsid_ = "{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"

class InprocCOMAddin:
    """תוסף COM שירוץ כ-InprocServer32"""
    
    _reg_clsid_ = _reg_clsid_
    _reg_progid_ = "InprocCOMAddin.Addin"
    _reg_desc_ = "Inproc COM Addin"
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_threading_ = "Apartment"
    
    def __init__(self):
        """אתחול התוסף"""
        try:
            # יצירת קובץ טסט
            test_file = os.path.join(os.environ['TEMP'], 'inproc_com_addin_loaded.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write(f"Inproc COM Addin Loaded Successfully!\nTime: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Inproc COM Addin initialized - test file created: {test_file}")
        except Exception as e:
            print(f"Error in InprocCOMAddin.__init__: {e}")
            traceback.print_exc()
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """חיבור ל-Outlook"""
        try:
            print("InprocCOMAddin.OnConnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'inproc_com_addin_connected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write(f"Inproc COM Addin Connected Successfully!\nTime: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Inproc COM Addin connected - test file created: {test_file}")
            return True
        except Exception as e:
            print(f"Error in InprocCOMAddin.OnConnection: {e}")
            traceback.print_exc()
            return False
    
    def OnDisconnection(self, remove_mode):
        """ניתוק מ-Outlook"""
        try:
            print("InprocCOMAddin.OnDisconnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'inproc_com_addin_disconnected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write(f"Inproc COM Addin Disconnected Successfully!\nTime: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Inproc COM Addin disconnected - test file created: {test_file}")
        except Exception as e:
            print(f"Error in InprocCOMAddin.OnDisconnection: {e}")
            traceback.print_exc()
    
    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            print("InprocCOMAddin.OnStartupComplete called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'inproc_com_addin_startup_complete.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write(f"Inproc COM Addin Startup Complete Successfully!\nTime: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Inproc COM Addin startup complete - test file created: {test_file}")
        except Exception as e:
            print(f"Error in InprocCOMAddin.OnStartupComplete: {e}")
            traceback.print_exc()
    
    def OnBeginShutdown(self, custom):
        """התחלת סגירת Outlook"""
        try:
            print("InprocCOMAddin.OnBeginShutdown called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'inproc_com_addin_shutdown.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write(f"Inproc COM Addin Shutdown Successfully!\nTime: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Inproc COM Addin shutdown - test file created: {test_file}")
        except Exception as e:
            print(f"Error in InprocCOMAddin.OnBeginShutdown: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    print("Registering Inproc COM Addin...")
    win32com.server.register.UseCommandLine(InprocCOMAddin)
    print("Inproc COM Addin registered successfully!")
