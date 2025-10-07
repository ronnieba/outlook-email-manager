# 🚀 מדריך יצירת VSTO Outlook Add-in עם Ribbon

## שלב 1️⃣: יצירת הפרויקט

### ב-Visual Studio 2022:

1. **פתח Visual Studio 2022**
   - לחץ על: `Create a new project`

2. **חפש "Outlook"**
   - בשורת החיפוש הקלד: `Outlook`
   - בחר: **Outlook VSTO Add-in**
   - (אם לא רואה, וודא שהתקנת "Office/SharePoint development")

3. **הגדרות הפרויקט:**
   ```
   Project name: AIEmailManagerAddin
   Location: C:\Users\ronni\outlook_email_manager\VSTO
   Solution name: AIEmailManagerAddin
   Framework: .NET Framework 4.8 (או הגבוה ביותר)
   ```

4. **לחץ Create**

---

## שלב 2️⃣: הוספת Ribbon

1. **לחיצה ימנית על הפרויקט**
   - `Add` → `New Item...`

2. **בחר Ribbon**
   - חפש: `Ribbon (Visual Designer)`
   - שם: `AIEmailRibbon`
   - לחץ `Add`

---

## שלב 3️⃣: עיצוב ה-Ribbon

### נוסיף:
- ✅ **Tab חדש**: "AI Email Manager"
- ✅ **Group**: "ניתוח מיילים"
- ✅ **כפתורים**:
  - 🔍 "נתח מייל נוכחי"
  - 📁 "נתח תיקיה"
  - ⚙️ "הגדרות"
  - 🌐 "פתח ממשק Web"

---

## שלב 4️⃣: הקוד

### `ThisAddIn.cs` - נקודת הכניסה
```csharp
using System;
using System.Net.Http;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;

namespace AIEmailManagerAddin
{
    public partial class ThisAddIn
    {
        private const string API_BASE_URL = "http://localhost:5000";
        private static readonly HttpClient client = new HttpClient();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // קוד אתחול
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // קוד סגירה
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
```

### `AIEmailRibbon.cs` - לוגיקת הכפתורים
```csharp
using System;
using System.Net.Http;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;

namespace AIEmailManagerAddin
{
    public partial class AIEmailRibbon
    {
        private const string API_BASE_URL = "http://localhost:5000";
        private static readonly HttpClient client = new HttpClient();

        private void AIEmailRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // אתחול
        }

        private void btnAnalyzeCurrent_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer.Selection.Count > 0)
                {
                    var mailItem = explorer.Selection[1] as Outlook.MailItem;
                    if (mailItem != null)
                    {
                        AnalyzeEmail(mailItem);
                    }
                }
                else
                {
                    MessageBox.Show("אנא בחר מייל לניתוח", "AI Email Manager", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "AI Email Manager", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void AnalyzeEmail(Outlook.MailItem mailItem)
        {
            try
            {
                // הכנת הנתונים
                var emailData = new
                {
                    subject = mailItem.Subject,
                    body = mailItem.Body,
                    sender = mailItem.SenderEmailAddress,
                    received_time = mailItem.ReceivedTime.ToString()
                };

                // שליחה ל-API
                var json = JsonConvert.SerializeObject(emailData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                
                var response = await client.PostAsync($"{API_BASE_URL}/api/analyze", content);
                
                if (response.IsSuccessStatusCode)
                {
                    var resultJson = await response.Content.ReadAsStringAsync();
                    dynamic analysis = JsonConvert.DeserializeObject(resultJson);
                    
                    // עדכון המייל
                    UpdateEmailWithAnalysis(mailItem, analysis);
                    
                    MessageBox.Show("המייל נותח בהצלחה!", "AI Email Manager", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"שגיאה בניתוח: {response.StatusCode}", 
                        "AI Email Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "AI Email Manager", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateEmailWithAnalysis(Outlook.MailItem mailItem, dynamic analysis)
        {
            try
            {
                // הוספת קטגוריה
                if (analysis.category != null)
                {
                    mailItem.Categories = analysis.category.ToString();
                }

                // הגדרת דחיפות
                if (analysis.priority != null)
                {
                    string priority = analysis.priority.ToString();
                    if (priority == "גבוהה")
                        mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                    else if (priority == "נמוכה")
                        mailItem.Importance = Outlook.OlImportance.olImportanceLow;
                }

                // הוספת דגל
                if (analysis.requires_action != null && (bool)analysis.requires_action)
                {
                    mailItem.FlagRequest = "למעקב";
                }

                // שמירת ניתוח מפורט
                var userProperty = mailItem.UserProperties.Add(
                    "AI Analysis", 
                    Outlook.OlUserPropertyType.olText);
                userProperty.Value = JsonConvert.SerializeObject(analysis);

                mailItem.Save();
            }
            catch (Exception ex)
            {
                throw new Exception($"שגיאה בעדכון המייל: {ex.Message}");
            }
        }

        private void btnAnalyzeFolder_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("ניתוח תיקיה - בפיתוח", "AI Email Manager", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("הגדרות - בפיתוח", "AI Email Manager", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnOpenWeb_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("http://localhost:5000");
        }
    }
}
```

---

## שלב 5️⃣: התקנת Packages

### ב-Visual Studio:
1. **Tools** → **NuGet Package Manager** → **Manage NuGet Packages for Solution**
2. התקן:
   - `Newtonsoft.Json` (לעבודה עם JSON)
   - `System.Net.Http` (לקריאות ל-API)

---

## שלב 6️⃣: בנייה והרצה

1. **לחץ F5** או **Debug → Start Debugging**
2. Outlook ייפתח אוטומטית עם התוסף
3. חפש את ה-Tab החדש: **"AI Email Manager"**

---

## 🎨 עיצוב ה-Ribbon (Visual Designer)

### בקובץ `AIEmailRibbon.cs` [Design]:

1. **הוסף Tab חדש:**
   - גרור `Tab` מה-Toolbox
   - שנה `Label`: "AI Email Manager"

2. **הוסף Group:**
   - גרור `Group` לתוך ה-Tab
   - שנה `Label`: "ניתוח מיילים"

3. **הוסף Buttons:**
   
   **כפתור 1:**
   - גרור `Button` לתוך ה-Group
   - `Name`: `btnAnalyzeCurrent`
   - `Label`: "נתח מייל נוכחי"
   - `ScreenTip`: "ניתוח המייל שבחרת עם AI"
   - `Size`: `Large`
   - double-click → יצירת event handler

   **כפתור 2:**
   - `Name`: `btnAnalyzeFolder`
   - `Label`: "נתח תיקיה"
   - `ScreenTip`: "ניתוח כל המיילים בתיקיה"
   - `Size`: `Large`

   **כפתור 3:**
   - `Name`: `btnSettings`
   - `Label`: "הגדרות"
   - `Size`: `Normal`

   **כפתור 4:**
   - `Name`: `btnOpenWeb`
   - `Label`: "פתח ממשק Web"
   - `Size`: `Normal`

---

## 🔧 פתרון בעיות

### בעיה: "Office/SharePoint development" לא מותקן
**פתרון:**
1. פתח **Visual Studio Installer**
2. לחץ **Modify**
3. סמן **Office/SharePoint development**
4. לחץ **Install**

### בעיה: Outlook לא פותח עם התוסף
**פתרון:**
1. סגור את כל חלונות Outlook
2. ב-Visual Studio: **Debug → Start Debugging** (F5)

### בעיה: שגיאת "API לא זמין"
**פתרון:**
1. ודא שהשרת Python רץ:
   ```
   python app_with_ai.py
   ```
2. בדוק: http://localhost:5000

---

## ✅ צ'קליסט

- [ ] Visual Studio 2022 מותקן
- [ ] Office/SharePoint development מותקן
- [ ] פרויקט VSTO נוצר
- [ ] Ribbon נוסף
- [ ] כפתורים נוספו
- [ ] NuGet packages הותקנו
- [ ] השרת Python רץ
- [ ] F5 - בנייה והרצה
- [ ] ה-Ribbon נראה ב-Outlook
- [ ] הכפתורים עובדים

---

## 🎉 סיימנו!

**עכשיו יש לך:**
- ✅ Ribbon מקצועי ב-Outlook
- ✅ כפתורים בעברית
- ✅ חיבור ל-Python API
- ✅ עדכון אוטומטי של מיילים

**להתחיל לעבוד:**
1. הפעל את השרת: `python app_with_ai.py`
2. הפעל את ה-Add-in: F5 ב-Visual Studio
3. פתח מייל ב-Outlook
4. לחץ על "נתח מייל נוכחי"
5. תהנה! 🎊

