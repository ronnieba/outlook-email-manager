# -*- coding: utf-8 -*-
"""
AI Email Manager - Outlook COM Add-in
This is the main file for the Outlook COM Add-in, now with Ribbon UI support.
It adds a custom tab to the Outlook Ribbon and handles button clicks.
"""

import win32com.client
from win32com.client import constants
import pythoncom
import os
from datetime import datetime

# --- Safe Import and Logging Setup ---
# This is critical for COM add-ins, as errors during startup can be silent.
try:
    import logging
    import requests
    # Log to a safe, user-writable location
    LOG_FILE = os.path.join(os.environ.get('TEMP', os.getcwd()), 'outlook_addin.log')
except ImportError as e:
    # If basic imports fail, we can't even log. This is a fatal setup error.
    # We'll try to show a message box as a last resort.
    import ctypes
    ctypes.windll.user32.MessageBoxW(0, f"A critical Python library is missing: {e}. Please run install.bat.", "Add-in Load Error", 0x10)


# --- Basic Logging Setup ---
# This helps debug issues during the add-in's startup.
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

# --- Configuration ---
SERVER_URL = "http://localhost:5000"

def log_info(message):
    """Log informational messages."""
    logging.info(message)
    print(f"INFO: {message}")

def log_error(message, exc_info=False):
    """Log error messages."""
    logging.error(message, exc_info=exc_info)
    print(f"ERROR: {message}")

log_info("Outlook Add-in script started.")

# --- Ribbon XML Definition ---
RIBBON_XML = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="aiEmailManagerTab" label="AI Email Manager">
        <group id="aiAnalysisGroup" label="ניתוח AI">
          <button id="analyzeEmailButton"
                  label="נתח מייל נוכחי"
                  size="large"
                  onAction="OnAnalyzeEmailPress"
                  imageMso="AnalyzeScenario" />
          <button id="openWebUIButton"
                  label="פתח ממשק Web"
                  size="large"
                  onAction="OnOpenWebUIPress"
                  imageMso="World" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""

class OutlookAddin:
    # --- COM Registration ---
    _public_methods_ = [
        'OnConnection', 'OnDisconnection', 'OnStartupComplete', 'OnBeginShutdown',
        'GetCustomUI', 'OnRibbonLoad', 'OnAnalyzeEmailPress', 'OnOpenWebUIPress'
    ]
    _reg_clsid_ = "{B6A7C267-343D-46A6-A655-3515A4252C16}" # Make sure this is unique
    _reg_progid_ = "AIEmailManager.Addin"
    _reg_desc_ = "AI Email Manager for Outlook"
    _reg_ver_ = "1.1" # Version updated
    _reg_threading_ = "Apartment"
    _reg_interfaces_ = [pythoncom.IID_IDispatch, '{000C0396-0000-0000-C000-000000000046}'] # IRibbonExtensibility

    def __init__(self):
        log_info("Add-in __init__ called.")
        self.Application = None
        self.addin_loaded = False

    def _co_initialize(self):
        """Ensures COM is initialized on the current thread."""
        try:
            pythoncom.CoInitialize()
        except pythoncom.com_error:
            pass # Already initialized

    def OnConnection(self, application, connectMode, addin, custom):
        """
        This is the primary entry point for the Add-in.
        """
        try:
            log_info(f"OnConnection event fired. ConnectMode: {connectMode}")
            self.Application = application
            self._co_initialize()
            self.addin_loaded = True
            log_info("Successfully connected to Outlook Application.")
        except Exception as e:
            log_error("Error in OnConnection", exc_info=True)
            self.addin_loaded = False

    def OnDisconnection(self, removeMode, custom):
        """
        Called when the Add-in is being unloaded.
        """
        self._co_initialize()
        log_info(f"OnDisconnection event fired. RemoveMode: {removeMode}")
        self.addin_loaded = False
        self.Application = None

    def OnStartupComplete(self, custom):
        """
        Called when Outlook has finished its startup process.
        """
        self._co_initialize()
        log_info("OnStartupComplete event fired.")
        if self.addin_loaded:
            log_info("Add-in is loaded and ready.")
            # The Ribbon is loaded via GetCustomUI, so no need to create it here.
        else:
            log_error("Startup complete, but add-in failed to load properly.")

    def OnBeginShutdown(self, custom):
        """
        Called when Outlook is beginning its shutdown process.
        """
        self._co_initialize()
        log_info("OnBeginShutdown event fired.")

    # --- IRibbonExtensibility Methods ---

    def GetCustomUI(self, RibbonID):
        """
        Called by Outlook to get the Ribbon XML.
        """
        self._co_initialize()
        log_info(f"GetCustomUI called for RibbonID: {RibbonID}")
        return RIBBON_XML

    # --- Ribbon Callbacks ---

    def OnRibbonLoad(self, ribbon):
        """
        Called when the Ribbon is loaded. We can save a reference to it.
        """
        self._co_initialize()
        log_info("Ribbon loaded successfully.")
        self.ribbon = ribbon

    def OnAnalyzeEmailPress(self, control):
        """
        Callback for the 'Analyze Email' button.
        """
        log_info(f"Button '{control.Id}' pressed.")
        self._co_initialize()
        try:
            if not self.Application:
                log_error("Outlook Application object not available.")
                return

            # Get the currently selected email
            active_explorer = self.Application.ActiveExplorer()
            if not active_explorer or active_explorer.Selection.Count == 0:
                log_info("No item selected in Outlook.")
                self.show_message_box("לא נבחר מייל לניתוח. אנא בחר מייל מהרשימה ונסה שוב.", "שגיאה", constants.olOkOnly | constants.olCritical)
                return

            selected_item = active_explorer.Selection.Item(1)

            # Check if it's a mail item
            if selected_item.Class == constants.olMail:
                subject = selected_item.Subject
                log_info(f"Preparing to analyze email: '{subject}'")
                # self.show_message_box(f"שולח לניתוח את המייל:\n'{subject}'", "ניתוח AI", constants.olOkOnly | constants.olInformation) # Removed for a quieter experience

                # --- Call the Flask server for analysis ---
                email_data = {
                    "subject": selected_item.Subject,
                    "body": selected_item.Body,
                    "sender_name": selected_item.SenderName,
                    "sender": selected_item.SenderEmailAddress,
                    "date": selected_item.ReceivedTime.isoformat()
                }

                try:
                    api_url = f"{SERVER_URL}/api/outlook-addin/analyze-email"
                    log_info(f"Sending request to {api_url}")
                    response = requests.post(api_url, json=email_data, timeout=60)
                    response.raise_for_status() # Will raise an exception for 4xx/5xx errors

                    analysis_result = response.json()
                    log_info(f"Received analysis: {analysis_result}")

                    if analysis_result.get("success"):
                        score = int(analysis_result.get('importance_score', 0) * 100)
                        summary = analysis_result.get('summary', 'לא נמצא סיכום.')
                        category = analysis_result.get('category', 'לא סווג')
                        
                        result_message = (f"ניתוח AI הושלם:\n\n"
                                          f"📊 ציון חשיבות: {score}%\n"
                                          f"🏷️ קטגוריה: {category}\n\n"
                                          f"📝 סיכום:\n{summary}")
                        self.show_message_box(result_message, "תוצאות ניתוח AI")
                    else:
                        error_details = analysis_result.get('error', 'שגיאה לא ידועה מהשרת')
                        log_error(f"Server returned an error: {error_details}")
                        self.show_message_box(f"השרת החזיר שגיאה:\n{error_details}", "שגיאת ניתוח", constants.olOkOnly | constants.olWarning)
                except requests.exceptions.RequestException as req_err:
                    log_error(f"Network error connecting to server: {req_err}", exc_info=True)
                    self.show_message_box(f"שגיאת רשת בחיבור לשרת.\n\nודא שהשרת הראשי (app_with_ai.py) פועל בכתובת:\n{SERVER_URL}", "שגיאת חיבור", constants.olOkOnly | constants.olCritical)
            else:
                log_info("Selected item is not a mail item.")
                self.show_message_box("הפריט שנבחר אינו מייל.", "שגיאה", constants.olOkOnly | constants.olWarning)

        except Exception as e:
            log_error("Error in OnAnalyzeEmailPress", exc_info=True)
            self.show_message_box(f"אירעה שגיאה בניתוח המייל:\n{e}", "שגיאה חמורה", constants.olOkOnly | constants.olCritical)

    def OnOpenWebUIPress(self, control):
        """
        Callback for the 'Open Web UI' button.
        """
        self._co_initialize()
        log_info(f"Button '{control.Id}' pressed. Opening web UI.")
        try:
            os.startfile(SERVER_URL)
        except Exception as e:
            log_error(f"Failed to open web UI: {e}", exc_info=True)
            self.show_message_box(f"לא ניתן לפתוח את כתובת ה-URL:\n{SERVER_URL}", "שגיאה", constants.olOkOnly | constants.olCritical)

    def show_message_box(self, text, title, style=None):
        """Shows a message box. Defers constant access to runtime."""
        if style is None:
            # Default style: OK button with Information icon
            style = constants.olOkOnly | constants.olInformation
        self.Application.Session.MessageBox(text, style, title)


def RegisterAddin(klass):
    """
    Registers the Add-in with COM.
    This is typically called by an installation script (e.g., PowerShell or Batch).
    """
    import win32com.server.register
    win32com.server.register.UseCommandLine(klass)

def UnregisterAddin(klass):
    """
    Unregisters the Add-in.
    """
    import win32com.server.register
    win32com.server.register.UseCommandLine(klass, unregister=True)


if __name__ == '__main__':
    """
    This part allows the script to be registered from the command line.
    Example: python outlook_com_addin.py --register or --unregister
    """
    import sys
    if len(sys.argv) > 1:
        if sys.argv[1] == '--register' or sys.argv[1] == '--unregister':
            try:
                import win32com.server.register
                win32com.server.register.UseCommandLine(OutlookAddin)
            except Exception as e:
                print(f"Error during registration process: {e}")
        else:
            print(f"Unknown command: {sys.argv[1]}")
    else:
        print("This script is an Outlook Add-in. Use --register or --unregister.")

log_info("Add-in script finished execution.")