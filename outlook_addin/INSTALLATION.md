# Outlook Add-in - הוראות התקנה

## 📋 דרישות מקדימות

1. **Microsoft Outlook** (גרסה 2016 ומעלה)
2. **השרת שלנו** פועל על `https://localhost:5000`
3. **SSL Certificate** מותקן (לצורך HTTPS)

## 🚀 שלבי התקנה

### שלב 1: הכנת השרת
1. ודא שהשרת פועל על HTTPS
2. ודא שכל הקבצים ב-`outlook_addin/` זמינים
3. בדוק שה-API endpoints עובדים:
   - `https://localhost:5000/api/outlook-addin/analyze-email`
   - `https://localhost:5000/api/outlook-addin/get-profile`
   - `https://localhost:5000/api/outlook-addin/update-profile`

### שלב 2: התקנת ה-Add-in ב-Outlook

#### דרך 1: התקנה ידנית (מומלץ לפיתוח)
1. פתח את **Outlook**
2. לחץ על **File** → **Manage Add-ins**
3. לחץ על **My add-ins** → **Add a custom add-in** → **Add from file**
4. בחר את הקובץ `outlook_addin/manifest.xml`
5. לחץ **Add**

#### דרך 2: התקנה דרך Registry (למשתמשים מתקדמים)
1. פתח **Registry Editor** (regedit)
2. נווט ל: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer`
3. צור מפתח חדש עם השם: `12345678-1234-1234-1234-123456789012`
4. הוסף ערך `string` עם השם `WebApplicationManifest` והערך: `https://localhost:5000/outlook_addin/manifest.xml`
5. הפעל מחדש את Outlook

### שלב 3: בדיקת ההתקנה
1. פתח Outlook
2. בחר מייל כלשהו
3. בדוק אם יש כפתור **AI Email Manager** בסרט
4. לחץ על הכפתור לפתיחת ה-Task Pane

## 🔧 פתרון בעיות

### בעיה: Add-in לא מופיע
**פתרון:**
- ודא שה-manifest.xml נגיש דרך HTTPS
- בדוק שה-URL ב-manifest.xml נכון
- ודא שה-Outlook תומך ב-Add-ins

### בעיה: שגיאת CORS
**פתרון:**
- ודא שהשרת מחזיר headers נכונים
- בדוק שה-URL הוא HTTPS ולא HTTP

### בעיה: שגיאת SSL
**פתרון:**
- התקן SSL certificate תקף
- או השתמש ב-self-signed certificate עם אישור ב-Outlook

### בעיה: Add-in לא מגיב
**פתרון:**
- בדוק את ה-Console ב-Outlook (F12)
- ודא שה-API endpoints עובדים
- בדוק שה-AI זמין במערכת

## 📱 שימוש ב-Add-in

### ניתוח מייל
1. בחר מייל ב-Outlook
2. לחץ על **AI Email Manager** בסרט
3. לחץ על **📊 ציון מיידי**
4. צפה בתוצאות הניתוח

### ניהול פרופיל
1. לחץ על **⚙️ הגדרות**
2. פתח את המערכת הראשית
3. עדכן את ההגדרות שלך

### סטטיסטיקות
1. לחץ על **📈 סטטיסטיקות**
2. צפה בנתונים מתקדמים

## 🔒 אבטחה

- כל התקשורת נעשית דרך HTTPS
- הנתונים נשמרים רק במערכת המקומית
- אין שליחה של נתונים לשרתים חיצוניים

## 📞 תמיכה

אם יש בעיות:
1. בדוק את הלוגים בקונסול המערכת
2. בדוק את ה-Console ב-Outlook (F12)
3. ודא שכל הדרישות מתקיימות

## 🔄 עדכונים

כדי לעדכן את ה-Add-in:
1. החלף את הקבצים ב-`outlook_addin/`
2. הפעל מחדש את Outlook
3. ה-Add-in יתעדכן אוטומטית






