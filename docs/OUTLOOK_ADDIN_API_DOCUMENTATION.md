# 🌐 תיעוד API - תוסף Outlook

תיעוד מפורט של ה-API של תוסף AI Email Manager ל-Microsoft Outlook.

## 🔌 תוסף COM API

### מבנה התוסף
```python
class OutlookAddin(win32com.server.policy.DesignatedWrapPolicy):
    """Outlook COM Add-in"""
    
    # COM registration
    _reg_clsid_ = "{12345678-1234-1234-1234-123456789012}"
    _reg_progid_ = "AIEmailManager.Addin"
    _reg_desc_ = "AI Email Manager Outlook Add-in"
    _reg_threading_ = "Apartment"
```

### מתודות התוסף

#### OnConnection
```python
def OnConnection(self, application, connect_mode, add_in_inst, custom):
    """Called when add-in connects to Outlook"""
    self.outlook = application
    self.namespace = self.outlook.GetNamespace("MAPI")
    return True
```

**פרמטרים:**
- `application`: אובייקט Outlook Application
- `connect_mode`: מצב החיבור (ext_ConnectMode)
- `add_in_inst`: אובייקט התוסף
- `custom`: פרמטרים מותאמים אישית

**החזרה:**
- `True` - חיבור מוצלח
- `False` - חיבור נכשל

#### OnDisconnection
```python
def OnDisconnection(self, remove_mode, custom):
    """Called when add-in disconnects from Outlook"""
    self.outlook = None
    self.namespace = None
    return True
```

**פרמטרים:**
- `remove_mode`: מצב ההסרה (ext_DisconnectMode)
- `custom`: פרמטרים מותאמים אישית

**החזרה:**
- `True` - ניתוק מוצלח
- `False` - ניתוק נכשל

#### OnStartupComplete
```python
def OnStartupComplete(self, custom):
    """Called when add-in startup is complete"""
    return True
```

**פרמטרים:**
- `custom`: פרמטרים מותאמים אישית

**החזרה:**
- `True` - הפעלה מוצלחת
- `False` - הפעלה נכשלת

#### OnBeginShutdown
```python
def OnBeginShutdown(self, custom):
    """Called when add-in begins shutdown"""
    return True
```

**פרמטרים:**
- `custom`: פרמטרים מותאמים אישית

**החזרה:**
- `True` - סגירה מוצלחת
- `False` - סגירה נכשלת

## 🌐 תוסף Office API

### מבנה המניפסט
```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
    <Id>12345678-1234-1234-1234-123456789012</Id>
    <Version>1.0.0.0</Version>
    <ProviderName>AI Email Manager</ProviderName>
    <DisplayName DefaultValue="AI Email Manager" />
    <Description DefaultValue="AI-powered email and meeting analysis for Outlook"/>
    
    <Hosts>
        <Host Name="Mailbox" />
    </Hosts>
    
    <DefaultSettings>
        <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
    </DefaultSettings>
    
    <Permissions>ReadWriteMailbox</Permissions>
</OfficeApp>
```

### JavaScript API

#### Office.onReady
```javascript
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('Connected to Outlook');
    }
});
```

#### ניתוח מייל נוכחי
```javascript
async function analyzeCurrentEmail() {
    const item = Office.context.mailbox.item;
    const emailData = {
        subject: item.subject,
        body: item.body.getAsync({ coercionType: Office.CoercionType.Text }),
        sender: item.from.emailAddress,
        receivedTime: item.dateTimeCreated
    };
    
    const response = await fetch('http://localhost:5000/analyze_email', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(emailData)
    });
    
    const result = await response.json();
    showAnalysisResult(result);
}
```

#### ניתוח מיילים נבחרים
```javascript
async function analyzeSelectedEmails() {
    const selection = Office.context.mailbox.item;
    // Process selected emails
}
```

#### ניתוח פגישה נוכחית
```javascript
async function analyzeCurrentMeeting() {
    const item = Office.context.mailbox.item;
    // Process current meeting
}
```

#### פתיחת ממשק אינטרנט
```javascript
function openWebInterface() {
    window.open('http://localhost:5000', '_blank');
}
```

## 🔗 API Backend

### ניתוח מייל
```http
POST /analyze_email
Content-Type: application/json

{
    "subject": "Subject of the email",
    "body": "Body content of the email",
    "sender": "sender@example.com",
    "receivedTime": "2024-01-01T10:00:00Z"
}
```

**תגובה:**
```json
{
    "analysis": "Analysis of the email",
    "importance": 8,
    "category": "work",
    "action_items": ["Action 1", "Action 2"],
    "estimated_time": "15 minutes"
}
```

### ניתוח פגישה
```http
POST /analyze_meeting
Content-Type: application/json

{
    "subject": "Meeting subject",
    "body": "Meeting description",
    "organizer": "organizer@example.com",
    "attendees": ["attendee1@example.com", "attendee2@example.com"],
    "startTime": "2024-01-01T10:00:00Z",
    "endTime": "2024-01-01T11:00:00Z"
}
```

**תגובה:**
```json
{
    "analysis": "Analysis of the meeting",
    "importance": 7,
    "category": "work",
    "goals": ["Goal 1", "Goal 2"],
    "preparation": ["Prepare agenda", "Review documents"]
}
```

## 📊 לוגים ו-API

### לוגי הצלחה
```python
# כתיבה ללוג הצלחה
with open("outlook_addin_success.log", "w", encoding="utf-8") as f:
    f.write("AI Email Manager add-in connected to Outlook successfully!\n")
    f.write(f"Connect Mode: {connect_mode}\n")
```

### לוגי שגיאות
```python
# כתיבה ללוג שגיאות
with open("outlook_addin_error.log", "w", encoding="utf-8") as f:
    f.write(f"Error connecting to Outlook: {e}\n")
    f.write(f"Error Type: {type(e).__name__}\n")
    f.write(f"Error Details: {str(e)}\n")
```

### קריאת לוגים
```bash
# קריאת לוג הצלחה
type outlook_addin_success.log

# קריאת לוג שגיאות
type outlook_addin_error.log
```

## 🔧 רישום ו-API

### רישום COM
```python
# רישום התוסף
if __name__ == "__main__":
    win32com.server.register.UseCommandLine(OutlookAddin)
```

### רישום Registry
```bash
# רישום התוסף ב-Outlook
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "FriendlyName" /t REG_SZ /d "AI Email Manager" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "Description" /t REG_SZ /d "AI-powered email and meeting analysis for Outlook" /f
```

### בדיקת רישום
```bash
# בדיקת רישום התוסף
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin"
```

## 🚀 התקנה ו-API

### התקנת תוסף COM
```bash
# התקנה אוטומטית
.\install_final_com_addin.bat

# התקנה ידנית
python outlook_com_addin_final.py --register
```

### התקנת תוסף Office
```bash
# התקנה אוטומטית
.\install_office_addin.bat

# התקנה ידנית
# 1. Start web server
python -m http.server 3000 --directory outlook_addin
# 2. Install add-in in Outlook
# File → Options → Add-ins → Web Add-ins → Add → Select manifest.xml
```

## 🔍 דיבוג ו-API

### בדיקת פעילות
```bash
# בדיקת לוגים
type outlook_addin_success.log
type outlook_addin_error.log

# בדיקת שרת אינטרנט
curl http://localhost:3000/taskpane.html

# בדיקת שרת AI
curl http://localhost:5000/health
```

### בדיקת רישום
```bash
# בדיקת רישום COM
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin"

# בדיקת רישום Office Add-in
# File → Options → Add-ins → Web Add-ins → Go...
```

## 📚 משאבים נוספים

### תיעוד Microsoft
- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Outlook COM Add-ins](https://docs.microsoft.com/en-us/office/vba/outlook/concepts/getting-started-with-vba-in-outlook)
- [Office.js API](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office)

### כלים לפיתוח
- [Office Add-in Tools](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/develop-add-ins-visual-studio)
- [Office Add-in Debugger](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Office Add-in Validator](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/validate-office-add-in-manifest)

---

**בהצלחה בפיתוח! 🚀**
















