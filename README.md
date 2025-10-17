# 📧 Outlook Email Manager with AI

מערכת ניהול מיילים חכמה המשלבת Microsoft Outlook עם בינה מלאכותית לניתוח אוטומטי של חשיבות המיילים.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Flask](https://img.shields.io/badge/Flask-2.3.3-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

## 🌟 תכונות עיקריות

### 📧 ניתוח AI חכם
- **ניתוח אוטומטי** - ניתוח חשיבות מיילים עם Gemini AI
- **ציון חשיבות** - מתן ציון 0-100 לכל מייל
- **קטגוריזציה** - זיהוי אוטומטי של סוגי מיילים (דחוף, חשוב, רגיל)
- **סיכום חכם** - יצירת סיכום קצר לכל מייל
- **פעולות מומלצות** - המלצות לטיפול במייל

### 🎨 תצוגה ויזואלית
- **עמודת AI Score** - הצגת ציון החשיבות ב-Outlook
- **קטגוריות צבעוניות** - סימון ויזואלי של רמת החשיבות
- **דגלים אוטומטיים** - סימון מיילים חשובים עם דגלים
- **ממשק אינטואיטיבי** - ממשק משתמש ידידותי בעברית

### 🖥️ ממשק ניהול Web
- **דף ניהול מיילים** - צפייה וניהול מיילים מותאמים
- **דף ניהול פגישות** - ניהול פגישות מ-Outlook
- **קונסול ניהול** - מעקב בזמן אמת על פעילות המערכת
- **סטטיסטיקות** - ניתוח דפוסי עבודה

## 🚀 התחלה מהירה

### דרישות מערכת
- Windows 10/11
- Python 3.8 ומעלה
- Microsoft Outlook 2016 ומעלה
- Google Gemini API Key (חינמי)

### התקנה מהירה (5 דקות)

1. **שכפול הפרויקט**
```bash
git clone https://github.com/your-username/outlook_email_manager.git
cd outlook_email_manager
```

2. **התקנת תלויות**
```bash
pip install -r requirements.txt
```

3. **הגדרת API Key** (⚠️ חשוב!)
```bash
# העתק את קובץ הדוגמה
copy env.example .env

# ערוך את .env והוסף את ה-API Key שלך
notepad .env
```

קבל API Key חינמי מ-[Google AI Studio](https://makersuite.google.com/app/apikey)

4. **בדיקת התקנה** (מומלץ)
```bash
python verify_installation.py
```

5. **הפעלת השרת**
```bash
python app_with_ai.py
```

6. **פתיחת הממשק**
פתח דפדפן וגש ל-`http://localhost:5000`

---

📖 **למדריך התקנה מפורט**: ראה [INSTALLATION_GUIDE_SIMPLE.md](INSTALLATION_GUIDE_SIMPLE.md)

## 📖 שימוש במערכת

### דרך 1: ניתוח מייל בודד

1. פתח את Outlook ובחר מייל
2. הפעל:
```bash
python working_email_analyzer.py
```
3. המערכת תנתח את המייל ותוסיף לו:
   - ציון חשיבות (AI_Score)
   - קטגוריה (AI_Category)
   - סיכום (AI_Summary)
   - דגל אוטומטי (לפי החשיבות)

### דרך 2: ממשק Web

1. פתח את הממשק ב-`http://localhost:5000`
2. עבור לדף "ניהול מיילים"
3. בחר מיילים לניתוח
4. צפה בתוצאות והסטטיסטיקות

### יצירת עמודת AI Score ב-Outlook

1. פתח Outlook והיכנס לתיקייה (למשל Inbox)
2. לחץ על **View → View Settings → Columns**
3. בחר **User-defined fields in folder**
4. לחץ על **New Field**
5. הזן:
   - **Name**: `AI_Score`
   - **Type**: `Number`
   - **Format**: `123`
6. לחץ **OK** והוסף את העמודה לתצוגה
7. גרור את העמודה למיקום הרצוי

## 📁 מבנה הפרויקט

```
outlook_email_manager/
│
├── 📄 Core Application Files
│   ├── app_with_ai.py              # 🖥️  אפליקציית Flask הראשית + API
│   ├── ai_analyzer.py              # 🤖 מנוע ניתוח AI (Gemini)
│   ├── user_profile_manager.py     # 👤 ניהול פרופיל + למידה
│   ├── working_email_analyzer.py   # 📧 מנתח standalone למייל בודד
│   ├── outlook_com_addin_final.py  # 🔌 COM Add-in ל-Outlook
│   ├── collapsible_logger.py       # 📝 מערכת לוגים מתקדמת
│   └── config.py                   # ⚙️  קובץ הגדרות (טוען .env)
│
├── 🔧 Configuration & Setup
│   ├── env.example                 # 🔑 דוגמה להגדרת משתני סביבה
│   ├── .env                        # 🔐 משתני סביבה (לא מגובה ל-Git)
│   ├── requirements.txt            # 📦 תלויות Python
│   ├── verify_installation.py      # ✅ בדיקת התקנה אוטומטית
│   ├── install_final_simple.bat    # 🚀 התקנת COM Add-in
│   └── .gitignore                  # 🚫 קבצים שלא מגובים
│
├── 📁 templates/                   # 🎨 תבניות HTML
│   ├── index.html                  #    דף ניהול מיילים
│   ├── meetings.html               #    דף ניהול פגישות
│   ├── consol.html                 #    קונסול ניהול בזמן אמת
│   └── profile.html                #    דף ניהול פרופיל
│
├── 📁 Cursor_Prompts/              # 💡 פרומפטים לפיתוח עם Cursor
│   ├── הסברים.txt                 #    הוראות מפורטות
│   ├── 01_Main_Project_Prompt.txt  #    פרומפט ראשי
│   ├── 02_Flask_Application.txt    #    Flask Backend
│   ├── 03_Frontend_Development.txt #    HTML/CSS/JS
│   ├── 04_Outlook_Integration.txt  #    COM Integration
│   ├── 05_AI_Integration.txt       #    Gemini AI
│   ├── 06_Deployment.txt           #    Deployment
│   └── README.md                   #    תיאור התיקייה
│
├── 📁 docs/                        # 📚 תיעוד מפורט
│   ├── README.md                   #    סקירה כללית
│   ├── INSTALLATION.md             #    מדריך התקנה מפורט
│   ├── USER_GUIDE.md               #    מדריך משתמש
│   ├── DEVELOPER_GUIDE.md          #    מדריך מפתח
│   ├── API_DOCUMENTATION.md        #    תיעוד API
│   ├── OUTLOOK_ADDIN_*.md          #    תיעוד Add-in
│   └── CHANGELOG.md                #    היסטוריית שינויים
│
├── 📁 Database Files (נוצרים אוטומטית)
│   ├── email_manager.db            # 🗄️  מסד נתונים ראשי
│   └── email_preferences.db        # 💾 העדפות משתמש
│
├── 📄 Documentation (Root Level)
│   ├── README.md                   # 📖 תיאור הפרויקט הראשי
│   ├── INSTALLATION_GUIDE_SIMPLE.md # 🚀 התקנה פשוטה
│   ├── SYSTEM_ARCHITECTURE.md      # 🏗️  ארכיטקטורת המערכת
│   ├── AISCORE_COLUMN_SETUP.md     # 📊 הגדרת עמודת AI ב-Outlook
│   ├── COM_ADDIN_REGISTRATION_GUIDE.md # 🔌 רישום COM Add-in
│   ├── GITHUB_BACKUP_GUIDE.md      # 💾 גיבוי GitHub
│   ├── VERIFICATION_REPORT.md      # ✅ דוח אימות
│   ├── TESTING_GUIDE.md            # 🧪 מדריך בדיקות
│   └── VISUAL_GUIDE.md             # 🎨 מדריך ויזואלי
│
├── 📁 outlook_addin/               # 🔧 Office Add-in (Web-based)
│   ├── manifest.xml                #    מניפסט Add-in
│   ├── taskpane.html               #    Task Pane UI
│   ├── taskpane.js                 #    לוגיקת Add-in
│   └── taskpane.css                #    עיצוב Add-in
│
└── 📁 Utility Scripts (לבדיקות)
    ├── create_test_emails_and_meetings.py
    ├── check_outlook_items.py
    └── create_full_backup.py
```

### 📝 הערות חשובות:
- **קבצי .env ו-.db**: לא מגובים ל-Git (נמצאים ב-.gitignore)
- **env.example**: דוגמה בלבד - העתק ל-.env והוסף API Key
- **בסיסי נתונים**: נוצרים אוטומטית בהרצה ראשונה

## 🔧 קבצים עיקריים

### Backend
- **`app_with_ai.py`** - שרת Flask עם API endpoints
- **`ai_analyzer.py`** - מודול ניתוח AI עם Gemini
- **`user_profile_manager.py`** - מערכת למידה והתאמה אישית
- **`collapsible_logger.py`** - מערכת לוגים מתקדמת

### Frontend
- **`templates/index.html`** - ממשק ניהול מיילים
- **`templates/meetings.html`** - ממשק ניהול פגישות
- **`templates/consol.html`** - קונסול ניהול

### Tools
- **`working_email_analyzer.py`** - כלי עצמאי לניתוח מיילים
- **`outlook_com_addin_final.py`** - תוסף COM ל-Outlook (מתקדם)

### Configuration
- **`config.py`** - הגדרות כלליות (API Keys, thresholds)
- **`requirements.txt`** - רשימת תלויות Python

## 📚 מדריכים מפורטים

- [📋 מדריך התקנה מפורט](INSTALLATION_GUIDE_SIMPLE.md)
- [🚀 התחלה מהירה](QUICK_START_OUTLOOK_ADDIN.md)
- [📊 הגדרת עמודת AI Score](AISCORE_COLUMN_SETUP.md)
- [✅ פתרון שעובד](FINAL_WORKING_SOLUTION.md)
- [🔄 מדריך Sync אוטומטי](AUTO_SYNC_GUIDE.md)
- [🎨 מדריך ויזואלי](VISUAL_GUIDE.md)
- [🧪 מדריך בדיקות](TESTING_GUIDE.md)
- [📖 תיעוד מפורט](docs/)

## 🔌 API Endpoints

### Emails
- `GET /api/emails` - קבלת רשימת מיילים
- `POST /api/analyze_email` - ניתוח מייל
- `POST /api/feedback` - משוב על ניתוח

### Meetings
- `GET /api/meetings` - קבלת רשימת פגישות
- `POST /api/meetings` - יצירת פגישה חדשה

### System
- `GET /api/status` - סטטוס המערכת
- `GET /api/console_logs` - לוגים בזמן אמת

לתיעוד מלא ראה [API Documentation](docs/API_DOCUMENTATION.md)

## 🗄️ בסיסי נתונים

המערכת משתמשת ב-SQLite עם שני קבצים:

- **`email_manager.db`** - מיילים, פגישות, ניתוחי AI
- **`email_preferences.db`** - העדפות משתמש, למידה

## 🤖 ניתוח AI

המערכת משתמשת ב-**Google Gemini API** לניתוח מיילים:

### מה המערכת מנתחת?
- נושא המייל
- שולח המייל
- תוכן המייל
- קיום קבצים מצורפים
- זמן קבלת המייל

### מה המערכת מחזירה?
- **ציון חשיבות** (0-100)
- **קטגוריה** (urgent, work, personal, marketing, etc.)
- **סיכום** קצר של המייל
- **פעולות מומלצות** (אופציונלי)

## 🎯 רמות חשיבות

| ציון | קטגוריה | סימון |
|------|---------|-------|
| 80-100 | דחוף | דגל אדום + סימן קריאה |
| 60-79 | חשוב | דגל צהוב |
| 40-59 | בינוני | ללא סימון |
| 0-39 | נמוך | ללא סימון |

## 🐛 פתרון בעיות

### השרת לא עובד
```bash
# בדוק שהפורט פנוי
netstat -ano | findstr :5000

# נסה פורט אחר
# ערוך את app_with_ai.py: app.run(port=5001)
```

### ניתוח AI לא עובד
1. בדוק את ה-API Key ב-`config.py`
2. ודא שיש חיבור לאינטרנט
3. בדוק את הלוגים: `%TEMP%\ai_email_manager.log`

### Outlook לא מתחבר
1. ודא ש-Outlook פתוח
2. נסה להפעיל Outlook כמנהל
3. בדוק שה-pywin32 מותקן: `pip install pywin32`

### עמודת AI Score לא מתמלאת
1. ודא שניתחת מייל אחרי יצירת העמודה
2. רענן את Outlook (F5)
3. בדוק שהשדה נוצר כ-Number ולא Text

## 🔒 אבטחה ופרטיות

- כל הנתונים נשמרים **מקומית** על המחשב שלך
- **אין העלאה לענן** של תוכן המיילים (מלבד לGemini API לניתוח)
- API Key נשמר **מקומית** בלבד
- ניתן להפעיל את המערכת **ללא חיבור לאינטרנט** (ללא AI)

## 🚧 פיתוח והרחבה

### הוספת מודל AI אחר
ערוך את `ai_analyzer.py` והחלף את `analyze_with_gemini()` במודל אחר.

### הוספת שדות מותאמים אישית
ערוך את `working_email_analyzer.py` והוסף UserProperties נוספים.

### שינוי עיצוב
ערוך את קבצי ה-HTML ב-`templates/` עם CSS מותאם אישית.

## 📝 רישיון

פרויקט זה מופץ תחת רישיון MIT. ראה קובץ `LICENSE` לפרטים.

## 🤝 תרומה לפרויקט

1. Fork את הפרויקט
2. צור branch חדש (`git checkout -b feature/amazing-feature`)
3. Commit את השינויים (`git commit -m 'Add amazing feature'`)
4. Push ל-branch (`git push origin feature/amazing-feature`)
5. פתח Pull Request

## 📞 תמיכה וקשר

- 🐛 **דיווח באגים**: פתח Issue ב-GitHub
- 💡 **הצעות תכונות**: פתח Issue עם תווית "enhancement"
- 📧 **שאלות**: השתמש ב-Discussions

## 🏆 הישגים

- ✅ אינטגרציה מלאה עם Microsoft Outlook
- ✅ ניתוח AI מתקדם עם Google Gemini
- ✅ ממשק משתמש אינטואיטיבי בעברית
- ✅ מערכת לוגים מתקדמת
- ✅ תמיכה מלאה ב-Custom Properties
- ✅ ממשק Web ניהול מתקדם

## 🗺️ Roadmap

### בקרוב
- [ ] ניתוח אוטומטי של מיילים חדשים
- [ ] תמיכה במודלי AI נוספים (ChatGPT, Claude)
- [ ] אפליקציית Mobile
- [ ] דוחות וגרפים מתקדמים

### עתידי
- [ ] תמיכה במספר חשבונות Email
- [ ] אינטגרציה עם Google Calendar
- [ ] תבניות תגובות אוטומטיות
- [ ] מערכת התראות חכמה

---

**פותח עם ❤️ בישראל** 🇮🇱

**גרסה**: 2.0  
**תאריך עדכון אחרון**: אוקטובר 2024  
**Python**: 3.8+  
**Outlook**: 2016+
