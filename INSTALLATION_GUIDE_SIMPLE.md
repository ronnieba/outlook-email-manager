# 🚀 AI Email Manager - מדריך התקנה פשוט

מדריך שלב אחר שלב להתקנת תוסף AI Email Manager ב-Microsoft Outlook.

## 🎯 מה זה AI Email Manager?

תוסף מתקדם ל-Outlook שמנתח מיילים ופגישות עם בינה מלאכותית ומציג:
- **ציון חשיבות** לכל מייל (0-100%)
- **קטגוריזציה אוטומטית** (דחוף, חשוב, רגיל)
- **סיכום חכם** של תוכן המייל
- **דגלים אוטומטיים** לפי חשיבות

## 📋 דרישות מערכת

### חומרה
- Windows 10/11
- 4GB RAM (מומלץ 8GB)
- 500MB מקום פנוי

### תוכנה
- **Python 3.8+** - [הורדה](https://www.python.org/downloads/)
- **Microsoft Outlook** 2016 או חדש יותר
- **Google Gemini API Key** - [קבל כאן](https://makersuite.google.com/app/apikey)

## 🚀 התקנה מהירה (5 דקות)

### שלב 1: הכנת הסביבה
```bash
# ודא ש-Python מותקן
python --version

# התקן תלויות
pip install pywin32 requests flask flask-cors google-generativeai
```

### שלב 2: הגדרת API Key
1. עבור ל-[Google AI Studio](https://makersuite.google.com/app/apikey)
2. צור API Key חדש
3. פתח את `config.py` והוסף:
```python
GEMINI_API_KEY = "your-api-key-here"
```

### שלב 3: התקנת התוסף
```bash
# הפעל את סקריפט ההתקנה
install_final_simple.bat
```

### שלב 4: הפעלת המערכת
```bash
# הפעל את השרת הראשי
python app_with_ai.py

# פתח את Outlook - התוסף אמור להופיע ב-Ribbon
```

## ✅ בדיקת התקנה

### בדיקה ב-Outlook
1. פתח Microsoft Outlook
2. חפש Tab חדש: **"AI Email Manager"**
3. בחר מייל ולחץ על **"נתח מייל נוכחי"**

### בדיקה ידנית
```bash
# בדוק שהתוסף נרשם
python outlook_com_addin_final.py --register

# בדוק את השרת
curl http://localhost:5000/api/status
```

## 🎯 שימוש בתוסף

### דרך Ribbon (מומלץ)
1. פתח Outlook
2. לחץ על Tab **"AI Email Manager"**
3. בחר מייל ולחץ **"נתח מייל נוכחי"**
4. התוצאות יופיעו בחלון נפרד

### אפשרויות נוספות
- **"נתח מיילים נבחרים"** - ניתוח מספר מיילים בבת אחת
- **"פתח ממשק Web"** - פתיחת הממשק המקוון
- **"הצג סטטיסטיקות"** - סטטיסטיקות על הניתוחים

## 📊 איך המידע ייראה

### במייל עצמו
```
===== 🤖 ניתוח AI =====
ציון חשיבות: 85%
קטגוריה: urgent
סיכום: מייל דחוף בנושא פרויקט חשוב

פעולות נדרשות:
- להגיב עד מחר
- לשלוח מסמכים
- לתאם פגישה
====================
```

### ב-Custom Properties
- `AI_Score`: 85%
- `AI_Category`: urgent
- `AI_Summary`: סיכום המייל
- `AI_Analyzed`: 2024-01-15 14:30

### בדגלים
- **דגל אדום** למיילים חשובים (80%+)
- **דגל צהוב** למיילים בינוניים (60%+)

## 🔧 יצירת עמודה AI ב-Outlook

### שיטה 1: אוטומטית (אם עובד)
התוסף ינסה ליצור עמודה אוטומטית.

### שיטה 2: ידנית
1. פתח Outlook
2. לחץ על **'תצוגה'** (View)
3. לחץ על **'הגדרות תצוגה'** (View Settings)
4. לחץ על **'עמודות'** (Columns)
5. לחץ על **'חדש...'** (New...)
6. הזן שם: `AI_Score`
7. בחר סוג: **טקסט** (Text)
8. לחץ **'אישור'**
9. גרור את השדה החדש לתצוגה
10. לחץ **'אישור'**

## 🐛 פתרון בעיות נפוצות

### בעיה: התוסף לא מופיע ב-Outlook
**פתרונות:**
1. ודא שהתוסף הותקן בהצלחה
2. סגור את Outlook לחלוטין
3. הפעל מחדש את Outlook
4. בדוק ב-Outlook: File → Options → Add-ins

### בעיה: שגיאת COM
**פתרונות:**
1. ודא ש-pywin32 מותקן: `pip install pywin32`
2. הפעל מחדש את Outlook
3. בדוק שאין חסימות אנטי-וירוס
4. נסה להפעיל את Outlook כמנהל

### בעיה: השרת לא זמין
**פתרונות:**
1. ודא שהשרת רץ: `python app_with_ai.py`
2. בדוק את הפורט: `http://localhost:5000`
3. בדוק את הגדרות Firewall
4. ודא שה-API Key תקין ב-`config.py`

### בעיה: ניתוח AI לא עובד
**פתרונות:**
1. בדוק את ה-API Key ב-`config.py`
2. ודא שיש חיבור לאינטרנט
3. בדוק את הלוגים ב-`%TEMP%\ai_email_manager.log`
4. נסה לנתח מייל פשוט

## 📝 לוגים וניטור

### קבצי לוג
- `%TEMP%\ai_email_manager.log` - לוגים של התוסף
- `email_manager.db` - בסיס נתונים של הניתוחים
- `email_preferences.db` - העדפות משתמש

### בדיקת לוגים
```bash
# בדיקת לוגים אחרונים
type "%TEMP%\ai_email_manager.log" | findstr ERROR

# בדיקת סטטיסטיקות
python -c "
import sqlite3
conn = sqlite3.connect('email_manager.db')
cursor = conn.cursor()
cursor.execute('SELECT COUNT(*) FROM email_ai_analysis')
print(f'מיילים נותחים: {cursor.fetchone()[0]}')
conn.close()
"
```

## 🔄 הסרת התוסף

### הסרה מלאה
```bash
# ביטול רישום התוסף
python outlook_com_addin_final.py --unregister

# מחיקת רישומים
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /f

# הפעל מחדש את Outlook
```

## 📞 תמיכה טכנית

### דיווח באגים
1. בדוק את הלוגים ב-`%TEMP%\ai_email_manager.log`
2. צלם מסך של השגיאה
3. פתח Issue עם פרטי השגיאה

### שאלות נפוצות
- **איך להסיר את התוסף?** - הרץ `python outlook_com_addin_final.py --unregister`
- **איך לעדכן את התוסף?** - הרץ `install_final_simple.bat` שוב
- **איך לשנות הגדרות?** - ערוך את `config.py`

### קישורים שימושיים
- [מדריך התקנה מפורט](docs/INSTALLATION.md)
- [מדריך משתמש](docs/USER_GUIDE.md)
- [תיעוד API](docs/API_DOCUMENTATION.md)

---

**פותח עם ❤️ בישראל** 🇮🇱

**גרסה**: 2.0  
**תאריך**: ינואר 2024  
**תמיכה**: Windows 10/11, Outlook 2016+


