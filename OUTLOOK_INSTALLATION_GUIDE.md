# AI Email Manager - התקנה ב-Outlook
# מדריך שלב אחר שלב להתקנת התוסף ב-Outlook

## 🎯 שלב 1: הכנת התוסף

### א. יצירת קובץ DLL (אופציונלי)
אם אתה רוצה תוסף "אמיתי" ב-Outlook, צריך ליצור קובץ DLL. אבל יש דרך פשוטה יותר:

### ב. שימוש ב-COM Add-in דרך Python
התוסף שלנו עובד דרך Python ו-COM, אז אנחנו יכולים להתקין אותו ישירות.

## 🔧 שלב 2: התקנה ידנית ב-Outlook

### א. פתיחת חלון תוספות COM
1. פתח את **Microsoft Outlook**
2. לחץ על **File** (קובץ)
3. לחץ על **Options** (אפשרויות)
4. לחץ על **Add-ins** (תוספים)
5. בתחתית החלון, ליד **Manage**, בחר **COM Add-ins**
6. לחץ על **Go...**

### ב. הוספת התוסף
1. בחלון **COM Add-ins** שיפתח, לחץ על **Add...**
2. נווט לתיקיית הפרויקט: `C:\Users\ronni\outlook_email_manager`
3. בחר את הקובץ: `outlook_com_addin.py`
4. לחץ על **Open**
5. לחץ על **OK**

## 🚀 שלב 3: התקנה אוטומטית (מומלץ)

### א. שימוש בסקריפט ההתקנה
```powershell
# הפעל PowerShell כמנהל
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\install_com_addin.ps1
```

### ב. שימוש בסקריפט Batch
```cmd
# הפעל Command Prompt כמנהל
install_com_addin.bat
```

## 📋 שלב 4: בדיקת ההתקנה

### א. בדיקה ב-Outlook
1. פתח את Outlook
2. לחץ על **File** → **Options** → **Add-ins**
3. בדוק שהתוסף **AI Email Manager** מופיע ברשימה
4. ודא שהוא מסומן ב-V (מופעל)

### ב. בדיקת פונקציונליות
1. פתח את Outlook
2. בחר מייל
3. לחץ לחיצה ימנית
4. בדוק אם יש אפשרות **"נתח עם AI"**

## 🎯 שלב 5: שימוש בתוסף

### א. דרך Context Menu (לחיצה ימנית)
1. בחר מייל ב-Outlook
2. לחץ לחיצה ימנית
3. בחר **"נתח עם AI"**

### ב. דרך Ribbon (אם הותקן)
1. פתח את Outlook
2. חפש Tab חדש: **"AI Email Manager"**
3. השתמש בכפתורים השונים

### ג. דרך Python ישירות
```bash
# הפעל את התוסף
python outlook_com_addin.py
```

## 🔧 שלב 6: פתרון בעיות

### בעיה: התוסף לא מופיע ב-Outlook
**פתרון:**
1. ודא שהתוסף נרשם ב-Registry
2. הפעל מחדש את Outlook
3. בדוק ב-Outlook: File → Options → Add-ins

### בעיה: שגיאת COM
**פתרון:**
1. ודא ש-pywin32 מותקן: `pip install pywin32`
2. הפעל מחדש את Outlook
3. בדוק שאין חסימות אנטי-וירוס

### בעיה: השרת לא זמין
**פתרון:**
1. ודא שהשרת רץ: `python app_with_ai.py`
2. בדוק את הפורט: `http://localhost:5000`
3. בדוק את הגדרות Firewall

## 🎉 סיכום

התוסף מותקן ב-Outlook ויכול לנתח מיילים עם AI. אתה יכול להשתמש בו דרך:
- **Context Menu** (לחיצה ימנית)
- **Ribbon** (אם הותקן)
- **Python ישירות**

האם תרצה שאני אסביר משהו ספציפי על ההתקנה?





