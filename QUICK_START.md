# ⚡ Quick Start - התחלה מהירה

## 🎯 הקמת הפרויקט ב-5 דקות

### ✅ דרישות מוקדמות
- Windows 10/11
- Python 3.8+
- Microsoft Outlook 2016+
- Git (אופציונלי)

---

## 📦 שלב 1: שכפול הפרויקט

```bash
# עם Git
git clone https://github.com/your-username/outlook_email_manager.git
cd outlook_email_manager

# או הורד ZIP ופרוס
```

---

## 🔧 שלב 2: התקנת תלויות

```bash
pip install -r requirements.txt
```

---

## 🔑 שלב 3: הגדרת API Key

```bash
# העתק קובץ דוגמה
copy env.example .env

# ערוך והוסף API Key
notepad .env
```

קבל API Key חינמי: https://makersuite.google.com/app/apikey

---

## ✅ שלב 4: בדיקה (אופציונלי)

```bash
python verify_installation.py
```

---

## 🚀 שלב 5: הפעלה

```bash
# הפעל שרת
python app_with_ai.py

# פתח דפדפן
start http://localhost:5000
```

---

## 🎉 זהו! המערכת פועלת!

### שימוש מהיר:

#### ניתוח מייל בודד:
1. פתח Outlook ובחר מייל
2. הרץ: `python working_email_analyzer.py`

#### ממשק Web:
1. גש ל-`http://localhost:5000`
2. בחר מיילים לניתוח
3. צפה בתוצאות

---

## 📚 מדריכים מפורטים

- 📖 **התקנה מפורטת**: [INSTALLATION_GUIDE_SIMPLE.md](INSTALLATION_GUIDE_SIMPLE.md)
- 🏗️ **ארכיטקטורה**: [SYSTEM_ARCHITECTURE.md](SYSTEM_ARCHITECTURE.md)
- 🔌 **COM Add-in**: [COM_ADDIN_REGISTRATION_GUIDE.md](COM_ADDIN_REGISTRATION_GUIDE.md)
- 📊 **עמודת AI**: [AISCORE_COLUMN_SETUP.md](AISCORE_COLUMN_SETUP.md)

---

## ❓ בעיות?

```bash
# הרץ בדיקה
python verify_installation.py

# בדוק לוגים
type "%TEMP%\ai_email_manager.log"
```

---

**פותח עם ❤️ בישראל** 🇮🇱



