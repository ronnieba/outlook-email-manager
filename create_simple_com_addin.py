import sys
import os
import subprocess
import tempfile

def create_simple_com_addin():
    """יוצר תוסף COM פשוט יותר"""
    
    # קוד Python פשוט יותר
    python_code = '''
import sys
import os
import tempfile
import win32com.server.register
import win32com.server.util
import win32com.client
import traceback
import time

# הגדרת CLSID ייחודי
_reg_clsid_ = "{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}"

class SimpleCOMAddin:
    """תוסף COM פשוט יותר"""
    
    _reg_clsid_ = _reg_clsid_
    _reg_progid_ = "SimpleCOMAddin.Addin"
    _reg_desc_ = "Simple COM Addin"
    _public_methods_ = ['OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown']
    _reg_threading_ = "Apartment"
    
    def __init__(self):
        """אתחול התוסף"""
        try:
            # יצירת קובץ טסט
            test_file = os.path.join(os.environ['TEMP'], 'simple_com_addin_loaded.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write(f"Simple COM Addin Loaded Successfully!\\nTime: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Simple COM Addin initialized - test file created: {test_file}")
        except Exception as e:
            print(f"Error in SimpleCOMAddin.__init__: {e}")
            traceback.print_exc()
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """חיבור ל-Outlook"""
        try:
            print("SimpleCOMAddin.OnConnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'simple_com_addin_connected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write(f"Simple COM Addin Connected Successfully!\\nTime: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Simple COM Addin connected - test file created: {test_file}")
            return True
        except Exception as e:
            print(f"Error in SimpleCOMAddin.OnConnection: {e}")
            traceback.print_exc()
            return False
    
    def OnDisconnection(self, remove_mode):
        """ניתוק מ-Outlook"""
        try:
            print("SimpleCOMAddin.OnDisconnection called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'simple_com_addin_disconnected.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write(f"Simple COM Addin Disconnected Successfully!\\nTime: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Simple COM Addin disconnected - test file created: {test_file}")
        except Exception as e:
            print(f"Error in SimpleCOMAddin.OnDisconnection: {e}")
            traceback.print_exc()
    
    def OnStartupComplete(self, custom):
        """השלמת אתחול Outlook"""
        try:
            print("SimpleCOMAddin.OnStartupComplete called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'simple_com_addin_startup_complete.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write(f"Simple COM Addin Startup Complete Successfully!\\nTime: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Simple COM Addin startup complete - test file created: {test_file}")
        except Exception as e:
            print(f"Error in SimpleCOMAddin.OnStartupComplete: {e}")
            traceback.print_exc()
    
    def OnBeginShutdown(self, custom):
        """התחלת סגירת Outlook"""
        try:
            print("SimpleCOMAddin.OnBeginShutdown called")
            # יצירת קובץ טסט נוסף
            test_file = os.path.join(os.environ['TEMP'], 'simple_com_addin_shutdown.txt')
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write(f"Simple COM Addin Shutdown Successfully!\\nTime: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Simple COM Addin shutdown - test file created: {test_file}")
        except Exception as e:
            print(f"Error in SimpleCOMAddin.OnBeginShutdown: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    print("Registering Simple COM Addin...")
    win32com.server.register.UseCommandLine(SimpleCOMAddin)
    print("Simple COM Addin registered successfully!")
'''
    
    # כתיבת הקוד לקובץ
    with open('simple_com_addin.py', 'w', encoding='utf-8') as f:
        f.write(python_code)
    
    print("Created simple_com_addin.py")
    return True

if __name__ == '__main__':
    print("Creating Simple COM Addin...")
    create_simple_com_addin()
    print("Simple COM Addin source created!")


