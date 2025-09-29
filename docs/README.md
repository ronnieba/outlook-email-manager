# 📧 Outlook Email Manager with AI

מערכת ניהול מיילים חכמה המשלבת Microsoft Outlook עם בינה מלאכותית לניתוח אוטומטי של חשיבות המיילים וניהול פגישות.

## 🌟 תכונות עיקריות

### 📧 ניהול מיילים חכם
- **ניתוח AI אוטומטי** - ניתוח חשיבות המיילים עם Gemini AI
- **סינון חכם** - מיילים קריטיים, חשובים, בינוניים ונמוכים
- **משוב משתמש** - מערכת למידה מהמשוב שלך
- **ניתוח קטגוריות** - זיהוי אוטומטי של סוגי מיילים

### 📅 ניהול פגישות
- **סינכרון Outlook** - טעינה אוטומטית של פגישות
- **כפתורי עדיפות** - סימון עדיפות פגישות עם LED חזותי
- **סטטיסטיקות** - ניתוח דפוסי פגישות
- **ניהול למידה** - מערכת למידה מתקדמת

### 🖥️ קונסול ניהול
- **מעקב בזמן אמת** - לוגים חיים של פעילות המערכת
- **ניהול שרת** - הפעלה מחדש וגיבויים
- **פרומפטים ל-Cursor** - יצירת קבצי עזר לפיתוח

## 🚀 התחלה מהירה

### דרישות מערכת
- Windows 10/11
- Python 3.8+
- Microsoft Outlook
- Google Gemini API Key

### התקנה מהירה
```powershell
# הפעלת הפרויקט
.\quick_start.ps1
```

### הפעלה ידנית
```powershell
# התקנת תלויות
pip install -r requirements.txt

# הפעלת השרת
python app_with_ai.py
```

## 📁 מבנה הפרויקט

```
outlook_email_manager/
├── 📧 app_with_ai.py          # אפליקציה ראשית
├── 🤖 ai_analyzer.py          # מנוע AI
├── 👤 user_profile_manager.py # ניהול פרופיל משתמש
├── 📄 config.py               # הגדרות
├── 📁 templates/              # תבניות HTML
│   ├── index.html            # דף ניהול מיילים
│   ├── meetings.html         # דף ניהול פגישות
│   └── consol.html           # דף קונסול
├── 📁 docs/                  # תיעוד
├── 📁 Cursor_Prompts/        # פרומפטים לפיתוח
└── 📁 Old/                   # קבצים ישנים
```

## 🔧 הגדרה

### 1. הגדרת Outlook
- התקן Microsoft Outlook
- התחבר לחשבון שלך
- הפעל את הפרויקט

### 2. הגדרת AI
- קבל API Key מ-Google Gemini
- הוסף את המפתח לקובץ `config.py`
- הפעל את המערכת

### 3. הגדרת בסיס נתונים
- המערכת יוצרת אוטומטית את בסיס הנתונים
- נתונים נשמרים ב-`email_manager.db`

## 📖 מדריכים

- [📋 מדריך התקנה מפורט](INSTALLATION.md)
- [👤 מדריך משתמש](USER_GUIDE.md)
- [🔧 מדריך מפתח](DEVELOPER_GUIDE.md)
- [🌐 תיעוד API](API_DOCUMENTATION.md)

## 🤝 תרומה לפרויקט

1. Fork את הפרויקט
2. צור branch חדש (`git checkout -b feature/amazing-feature`)
3. Commit את השינויים (`git commit -m 'Add amazing feature'`)
4. Push ל-branch (`git push origin feature/amazing-feature`)
5. פתח Pull Request

## 📝 רישיון

פרויקט זה מופץ תחת רישיון MIT. ראה קובץ `LICENSE` לפרטים נוספים.

## 📞 תמיכה

- 🐛 דיווח באגים: פתח Issue חדש
- 💡 הצעות תכונות: פתח Issue עם תווית "enhancement"
- 📧 שאלות: צור קשר דרך Issues

## 🏆 הישגים

- ✅ אינטגרציה מלאה עם Microsoft Outlook
- ✅ ניתוח AI מתקדם עם Gemini
- ✅ ממשק משתמש אינטואיטיבי
- ✅ מערכת למידה אדפטיבית
- ✅ ניהול פגישות חכם
- ✅ קונסול ניהול מתקדם

---

**פותח עם ❤️ בישראל** 🇮🇱
