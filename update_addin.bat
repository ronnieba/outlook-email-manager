@echo off
chcp 65001 > nul
echo ================================
echo עדכון תוסף AI Email Manager
echo ================================
echo.

echo שלב 1: סגירת Outlook...
taskkill /F /IM OUTLOOK.EXE 2>nul
timeout /t 2 /nobreak >nul
echo ✓ Outlook נסגר

echo.
echo שלב 2: ביטול רישום התוסף הישן...
python outlook_com_addin_final.py --unregister
echo ✓ התוסף הישן בוטל

echo.
echo שלב 3: רישום התוסף החדש...
python outlook_com_addin_final.py --register
echo ✓ התוסף החדש נרשם

echo.
echo שלב 4: פתיחת Outlook...
start "" "OUTLOOK.EXE"
timeout /t 3 /nobreak >nul

echo.
echo ================================
echo ✓ העדכון הושלם בהצלחה!
echo ================================
echo.
echo התוסף המעודכן כולל:
echo - שדה AISCORE מספרי לתצוגה בעמודה
echo - קטגוריה "AI: XX%%" לתצוגה מיידית
echo - חשיבות אוטומטית (אדום/רגיל/נמוך)
echo - דגלים צבעוניים
echo.
echo קרא את הקובץ AISCORE_COLUMN_SETUP.md להוראות מפורטות
echo.
pause
