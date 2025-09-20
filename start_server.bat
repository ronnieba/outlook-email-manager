@echo off
echo 🚀 מפעיל את Outlook Email Manager...
echo.

cd /d "%~dp0"

echo 📦 בודק תלויות...
pip install -r requirements.txt

echo.
echo 🐍 מפעיל את השרת...
python app_outlook_fixed.py

pause


