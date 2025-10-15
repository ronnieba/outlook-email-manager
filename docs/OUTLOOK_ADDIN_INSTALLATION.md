# 🔌 מדריך התקנה מפורט - תוסף Outlook

מדריך שלב-אחר-שלב להתקנת תוסף AI Email Manager ב-Microsoft Outlook.

## 🔧 דרישות מערכת

### חומרה
- **מעבד**: Intel Core i3 או AMD Ryzen 3 ומעלה
- **זיכרון**: 4GB RAM (מומלץ 8GB)
- **אחסון**: 100MB מקום פנוי לתוסף
- **מערכת הפעלה**: Windows 10/11

### תוכנה
- **Python 3.8+** - [הורדה](https://www.python.org/downloads/)
- **Microsoft Outlook** - גרסה 2016 ומעלה
- **pywin32** - להתקנת תוסף COM
- **Google Gemini API Key** - לניתוח AI

## 🚀 התקנה מהירה

### שלב 1: הורדת הפרויקט
```bash
# דרך Git
git clone https://github.com/your-repo/outlook-email-manager.git
cd outlook-email-manager

# או הורדה ישירה
# הורד את הקובץ ZIP ופתח אותו
```

### שלב 2: התקנת Python
1. הורד Python מ-[python.org](https://www.python.org/downloads/)
2. התקן עם אפשרות "Add to PATH"
3. בדוק התקנה:
```bash
python --version
pip --version
```

### שלב 3: התקנת תלויות
```bash
pip install -r requirements.txt
pip install pywin32
```

### שלב 4: הגדרת Gemini AI
1. עבור ל-[Google AI Studio](https://makersuite.google.com/app/apikey)
2. צור API Key חדש
3. העתק את המפתח
4. פתח את `config.py` והוסף:
```python
GEMINI_API_KEY = "your-api-key-here"
```

### שלב 5: התקנת תוסף Outlook

#### אפשרות A: תוסף COM (מומלץ)
```bash
# התקנה אוטומטית
.\install_final_com_addin.bat

# או התקנה ידנית
python outlook_com_addin_final.py --register
```

#### אפשרות B: תוסף Office (Web Add-in)
```bash
# התקנה אוטומטית
.\install_office_addin.bat

# או התקנה ידנית
# 1. Start web server
python -m http.server 3000 --directory outlook_addin
# 2. Install add-in in Outlook
# File → Options → Add-ins → Web Add-ins → Add → Select manifest.xml
```

### שלב 6: הפעלה
```bash
# הפעלת שרת AI
python app_with_ai.py
```

## 🔧 התקנה ידנית מפורטת

### שלב 1: הכנת הסביבה

#### בדיקת Python
```bash
python --version
# צריך להציג Python 3.8.0 או גרסה חדשה יותר
```

#### יצירת סביבה וירטואלית (מומלץ)
```bash
python -m venv outlook_manager_env
outlook_manager_env\Scripts\activate
```

### שלב 2: התקנת חבילות

#### חבילות בסיסיות
```bash
pip install flask==2.3.3
pip install flask-cors==4.0.0
pip install pywin32>=307
pip install google-generativeai==0.3.2
```

#### או התקנה מקובץ requirements
```bash
pip install -r requirements.txt
```

### שלב 3: הגדרת Outlook

#### בדיקת Outlook
1. פתח Microsoft Outlook
2. התחבר לחשבון שלך
3. ודא שיש לך גישה למיילים ופגישות

#### הרשאות COM
- Outlook צריך להיות פתוח בעת הפעלת הפרויקט
- ודא שאין חסימות אנטי-וירוס ל-COM objects

### שלב 4: הגדרת AI

#### קבלת API Key
1. עבור ל-[Google AI Studio](https://makersuite.google.com/app/apikey)
2. התחבר עם חשבון Google
3. לחץ "Create API Key"
4. העתק את המפתח

#### הגדרת המפתח
```python
# בקובץ config.py
GEMINI_API_KEY = "AIzaSyBOUWyZ-Dq2yPopzSZ6oopN7V6oeoB2iNY"  # המפתח שלך
```

### שלב 5: התקנת תוסף Outlook

#### תוסף COM (מומלץ)
```bash
# התקנה אוטומטית
.\install_final_com_addin.bat

# או התקנה ידנית
python outlook_com_addin_final.py --register
```

#### תוסף Office (Web Add-in)
```bash
# התקנה אוטומטית
.\install_office_addin.bat

# או התקנה ידנית
# 1. פתח Outlook
# 2. File → Options → Add-ins
# 3. בחר "Web Add-ins" ולחץ "Go..."
# 4. לחץ "Add..." ובחר את manifest.xml
```

#### בדיקת התוסף
1. פתח Microsoft Outlook
2. File → Options → Add-ins
3. בחר "COM Add-ins" או "Web Add-ins"
4. לחץ "Go..."
5. ודא ש-"AI Email Manager" מופיע ומסומן

### שלב 6: בדיקת התקנה

#### בדיקת חיבורים
```bash
python -c "import win32com.client; print('Outlook COM: OK')"
python -c "import google.generativeai; print('Gemini AI: OK')"
```

#### הפעלת השרת
```bash
python app_with_ai.py
```

#### בדיקת דפדפן
פתח דפדפן ב-`http://localhost:5000`

## 🐛 פתרון בעיות נפוצות

### בעיה: Python לא נמצא
```bash
# פתרון: הוסף Python ל-PATH
# או השתמש בנתיב המלא
C:\Python39\python.exe app_with_ai.py
```

### בעיה: Outlook לא נפתח
- ודא ש-Outlook מותקן ופתוח
- בדוק שאין חסימות אנטי-וירוס
- נסה להפעיל את Outlook כמנהל

### בעיה: API Key לא עובד
- בדוק שהמפתח תקין ב-Google AI Studio
- ודא שיש לך quota זמין
- בדוק את החיבור לאינטרנט

### בעיה: Port תפוס
```bash
# שנה את הפורט בקובץ app_with_ai.py
app.run(host='0.0.0.0', port=5001)  # במקום 5000
```

### בעיה: מודולים חסרים
```bash
pip install --upgrade pip
pip install -r requirements.txt --force-reinstall
```

### בעיה: תוסף לא נטען
- ודא שהתוסף נרשם ב-COM
- בדוק את הלוגים ב-`outlook_addin_success.log`
- נסה להפעיל את Outlook כמנהל
- בדוק שאין חסימות אנטי-וירוס

### בעיה: תוסף לא מופיע ב-Outlook
- בדוק שהתוסף נרשם ב-Registry
- ודא ש-LoadBehavior = 3
- נסה להסיר ולהוסיף מחדש את התוסף
- בדוק שהתוסף תואם לגרסת Outlook שלך

## 🔄 עדכון הפרויקט

### עדכון דרך Git
```bash
git pull origin main
pip install -r requirements.txt --upgrade
```

### עדכון ידני
1. הורד את הגרסה החדשה
2. החלף את הקבצים הישנים
3. התקן תלויות חדשות:
```bash
pip install -r requirements.txt --upgrade
```

## 📞 תמיכה טכנית

אם נתקלת בבעיות:

1. **בדוק את הלוגים** - פתח את הקונסול ב-`http://localhost:5000/consol`
2. **בדוק דרישות** - ודא שכל הדרישות מותקנות
3. **נסה פתרון אחד** - פתור בעיה אחת בכל פעם
4. **דווח על באג** - פתח Issue עם פרטי השגיאה

## 🎯 שלבים הבאים

לאחר התקנה מוצלחת:

1. 📖 קרא את [מדריך המשתמש](OUTLOOK_ADDIN_USER_GUIDE.md)
2. 🔧 עיין ב-[מדריך המפתח](OUTLOOK_ADDIN_DEVELOPER_GUIDE.md)
3. 🌐 בדוק את [תיעוד ה-API](API_DOCUMENTATION.md)
4. 🚀 התחל להשתמש במערכת!

---

**בהצלחה בהתקנה! 🎉**










