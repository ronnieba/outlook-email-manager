# 🚀 התקנה מהירה - תוסף Outlook

מדריך מהיר להתקנת תוסף AI Email Manager ב-Microsoft Outlook.

## ⚡ התקנה מהירה (5 דקות)

### שלב 1: דרישות מערכת
- Windows 10/11
- Python 3.8+
- Microsoft Outlook 2016+
- Google Gemini API Key

### שלב 2: הורדת הפרויקט
```bash
git clone https://github.com/your-repo/outlook-email-manager.git
cd outlook-email-manager
```

### שלב 3: התקנת תלויות
```bash
pip install -r requirements.txt
pip install pywin32
```

### שלב 4: הגדרת AI
1. עבור ל-[Google AI Studio](https://makersuite.google.com/app/apikey)
2. צור API Key חדש
3. פתח את `config.py` והוסף:
```python
GEMINI_API_KEY = "your-api-key-here"
```

### שלב 5: התקנת התוסף
```bash
# תוסף COM (מומלץ)
.\install_final_com_addin.bat

# או תוסף Office (Web Add-in)
.\install_office_addin.bat
```

### שלב 6: הפעלה
```bash
# הפעלת שרת AI
python app_with_ai.py
```

## ✅ בדיקת התקנה

### בדיקת התוסף
1. פתח Microsoft Outlook
2. File → Options → Add-ins
3. בחר "COM Add-ins" או "Web Add-ins"
4. לחץ "Go..."
5. ודא ש-"AI Email Manager" מופיע ומסומן

### בדיקת השרת
פתח דפדפן ב-`http://localhost:5000`

## 🎯 שימוש ראשוני

### תוסף COM
התוסף פועל אוטומטית ברקע ומנתח מיילים ופגישות.

### תוסף Office
1. לחץ על כפתור "AI Email Manager" ב-Ribbon
2. השתמש בכפתורים לניתוח מיילים ופגישות

## 🐛 פתרון בעיות מהיר

### התוסף לא מופיע
- ודא שהתוסף הותקן בהצלחה
- בדוק את הלוגים ב-`outlook_addin_success.log`
- נסה להפעיל את Outlook כמנהל

### התוסף לא עובד
- ודא שהשרת רץ (`python app_with_ai.py`)
- בדוק את החיבור לאינטרנט
- ודא שה-API Key תקין

### שגיאות נפוצות
- **"Add-in not loaded"** - התוסף לא נטען
- **"Runtime error"** - שגיאת זמן ריצה
- **"Connection failed"** - חיבור לשרת נכשל

## 📚 מדריכים מפורטים

- [📋 מדריך התקנה מפורט](docs/INSTALLATION.md)
- [🔌 מדריך התקנה תוסף Outlook](docs/OUTLOOK_ADDIN_INSTALLATION.md)
- [🔌 מדריך משתמש תוסף Outlook](docs/OUTLOOK_ADDIN_USER_GUIDE.md)
- [🔧 מדריך מפתח](docs/DEVELOPER_GUIDE.md)
- [🔌 מדריך מפתח תוסף Outlook](docs/OUTLOOK_ADDIN_DEVELOPER_GUIDE.md)

## 📞 תמיכה

אם נתקלת בבעיות:
1. בדוק את הלוגים
2. נסה לפתור את הבעיה
3. דווח על באג עם פרטי השגיאה
4. צור קשר דרך Issues

---

**בהצלחה בהתקנה! 🎉**