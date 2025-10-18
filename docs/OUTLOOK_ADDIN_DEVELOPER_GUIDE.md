# 🔧 מדריך מפתח - תוסף Outlook

מדריך מפורט לפיתוח ותחזוקת תוסף AI Email Manager ל-Microsoft Outlook.

## 🏗️ ארכיטקטורה

### מבנה התוסף
```
outlook_email_manager/
├── 🔌 outlook_com_addin_final.py    # תוסף COM ראשי
├── 📁 outlook_addin/                # תוסף Office (Web Add-in)
│   ├── manifest.xml                 # מניפסט התוסף
│   ├── taskpane.html               # ממשק משתמש
│   └── assets/                     # קבצי עזר
├── 📄 install_final_com_addin.bat  # סקריפט התקנה COM
├── 📄 install_office_addin.bat     # סקריפט התקנה Office
└── 📄 outlook_addin_success.log    # לוגים
```

### סוגי תוספים

#### 1. תוסף COM
- **קובץ**: `outlook_com_addin_final.py`
- **טכנולוגיה**: Python + win32com
- **אינטגרציה**: ישירה עם Outlook
- **רישום**: COM registration + Registry

#### 2. תוסף Office (Web Add-in)
- **קבצים**: `outlook_addin/manifest.xml`, `outlook_addin/taskpane.html`
- **טכנולוגיה**: HTML/JavaScript + Office.js
- **אינטגרציה**: דרך Office Add-in framework
- **רישום**: Office Add-in registration

## 🔌 פיתוח תוסף COM

### מבנה הקוד
```python
class OutlookAddin(win32com.server.policy.DesignatedWrapPolicy):
    """Outlook COM Add-in"""
    
    # COM registration
    _reg_clsid_ = "{12345678-1234-1234-1234-123456789012}"
    _reg_progid_ = "AIEmailManager.Addin"
    _reg_desc_ = "AI Email Manager Outlook Add-in"
    _reg_threading_ = "Apartment"
    
    def OnConnection(self, application, connect_mode, add_in_inst, custom):
        """Called when add-in connects to Outlook"""
        pass
    
    def OnDisconnection(self, remove_mode, custom):
        """Called when add-in disconnects from Outlook"""
        pass
    
    def OnStartupComplete(self, custom):
        """Called when add-in startup is complete"""
        pass
    
    def OnBeginShutdown(self, custom):
        """Called when add-in begins shutdown"""
        pass
```

### רישום התוסף
```python
# Register the add-in
if __name__ == "__main__":
    win32com.server.register.UseCommandLine(OutlookAddin)
```

### התקנה
```bash
# Register COM add-in
python outlook_com_addin_final.py --register

# Unregister COM add-in
python outlook_com_addin_final.py --unregister
```

## 🌐 פיתוח תוסף Office

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

### מבנה ה-HTML
```html
<!DOCTYPE html>
<html>
<head>
    <title>AI Email Manager</title>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
</head>
<body>
    <div class="container">
        <h1>AI Email Manager</h1>
        <button onclick="analyzeCurrentEmail()">Analyze Current Email</button>
        <button onclick="analyzeSelectedEmails()">Analyze Selected Emails</button>
        <button onclick="analyzeCurrentMeeting()">Analyze Current Meeting</button>
        <button onclick="openWebInterface()">Open Web Interface</button>
    </div>
</body>
</html>
```

### JavaScript API
```javascript
// Initialize Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('Connected to Outlook');
    }
});

// Analyze current email
async function analyzeCurrentEmail() {
    const item = Office.context.mailbox.item;
    const emailData = {
        subject: item.subject,
        body: item.body.getAsync({ coercionType: Office.CoercionType.Text }),
        sender: item.from.emailAddress,
        receivedTime: item.dateTimeCreated
    };
    
    // Send to AI server
    const response = await fetch('http://localhost:5000/analyze_email', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(emailData)
    });
    
    const result = await response.json();
    showAnalysisResult(result);
}
```

## 🚀 התקנה ופריסה

### התקנת תוסף COM
```bash
# התקנה אוטומטית
.\install_final_com_addin.bat

# התקנה ידנית
python outlook_com_addin_final.py --register
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f
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

## 🔍 דיבוג וניטור

### לוגים
```python
# Success logs
with open("outlook_addin_success.log", "w", encoding="utf-8") as f:
    f.write("AI Email Manager add-in connected to Outlook successfully!\n")

# Error logs
with open("outlook_addin_error.log", "w", encoding="utf-8") as f:
    f.write(f"Error connecting to Outlook: {e}\n")
```

### בדיקת רישום
```bash
# Check COM registration
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin"

# Check Office Add-in registration
# File → Options → Add-ins → Web Add-ins → Go...
```

### בדיקת פעילות
```bash
# Check logs
type outlook_addin_success.log
type outlook_addin_error.log

# Check web server
curl http://localhost:3000/taskpane.html
```

## 🛠️ תחזוקה

### עדכון התוסף
1. עדכן את הקוד
2. בדוק שהתוסף עובד
3. עדכן את הלוגים
4. בדוק תאימות לגרסאות Outlook

### גיבוי
```bash
# Backup add-in files
copy outlook_com_addin_final.py backup/
copy outlook_addin/ backup/
copy *.log backup/
```

### ניקוי
```bash
# Clean logs
del outlook_addin_success.log
del outlook_addin_error.log

# Unregister add-in
python outlook_com_addin_final.py --unregister
```

## 🐛 פתרון בעיות

### בעיות נפוצות

#### תוסף לא נטען
- בדוק שהתוסף נרשם ב-COM
- ודא ש-LoadBehavior = 3
- בדוק את הלוגים
- נסה להפעיל את Outlook כמנהל

#### שגיאות זמן ריצה
- בדוק שהתוסף תואם לגרסת Outlook
- ודא שכל התלויות מותקנות
- בדוק את הלוגים
- נסה להסיר ולהוסיף מחדש

#### בעיות חיבור
- ודא שהשרת רץ
- בדוק את החיבור לאינטרנט
- ודא שה-API Key תקין
- בדוק את הלוגים

### כלי דיבוג
```python
# Debug mode
import logging
logging.basicConfig(level=logging.DEBUG)

# Error handling
try:
    # Add-in code
    pass
except Exception as e:
    logging.error(f"Add-in error: {e}")
    with open("outlook_addin_error.log", "a") as f:
        f.write(f"{datetime.now()}: {e}\n")
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













