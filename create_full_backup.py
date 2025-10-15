# -*- coding: utf-8 -*-
"""
יצירת גיבוי מלא של הפרויקט
"""

import os
import shutil
import zipfile
from datetime import datetime
import sqlite3

def create_backup():
    """יצירת גיבוי מלא"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = f"backup_{timestamp}"
    
    print("="*60)
    print("🔄 יצירת גיבוי מלא של הפרויקט")
    print("="*60)
    
    # יצירת תיקיית גיבוי
    os.makedirs(backup_dir, exist_ok=True)
    print(f"\n📁 נוצרה תיקייה: {backup_dir}")
    
    # 1. גיבוי בסיסי נתונים
    print("\n1️⃣ מגבה בסיסי נתונים...")
    if os.path.exists("email_manager.db"):
        shutil.copy2("email_manager.db", f"{backup_dir}/email_manager.db")
        print("   ✅ email_manager.db")
    if os.path.exists("email_preferences.db"):
        shutil.copy2("email_preferences.db", f"{backup_dir}/email_preferences.db")
        print("   ✅ email_preferences.db")
    
    # 2. גיבוי קבצי Python עיקריים
    print("\n2️⃣ מגבה קבצי Python...")
    python_files = [
        "app_with_ai.py",
        "ai_analyzer.py",
        "user_profile_manager.py",
        "working_email_analyzer.py",
        "outlook_com_addin_final.py",
        "collapsible_logger.py",
        "config.py"
    ]
    for file in python_files:
        if os.path.exists(file):
            shutil.copy2(file, f"{backup_dir}/{file}")
            print(f"   ✅ {file}")
    
    # 3. גיבוי תבניות HTML
    print("\n3️⃣ מגבה תבניות HTML...")
    if os.path.exists("templates"):
        shutil.copytree("templates", f"{backup_dir}/templates", dirs_exist_ok=True)
        print("   ✅ templates/")
    
    # 4. גיבוי תיעוד
    print("\n4️⃣ מגבה תיעוד...")
    doc_files = [
        "README.md",
        "INSTALLATION_GUIDE_SIMPLE.md",
        "AISCORE_COLUMN_SETUP.md",
        "FINAL_WORKING_SOLUTION.md",
        "QUICK_START_OUTLOOK_ADDIN.md",
        "AUTO_SYNC_GUIDE.md",
        "VISUAL_GUIDE.md",
        "TESTING_GUIDE.md",
        "VERIFICATION_REPORT.md",
        "SYSTEM_ARCHITECTURE.md"
    ]
    for file in doc_files:
        if os.path.exists(file):
            shutil.copy2(file, f"{backup_dir}/{file}")
            print(f"   ✅ {file}")
    
    # גיבוי תיקיית docs
    if os.path.exists("docs"):
        shutil.copytree("docs", f"{backup_dir}/docs", dirs_exist_ok=True)
        print("   ✅ docs/")
    
    # 5. גיבוי פרומפטים
    print("\n5️⃣ מגבה פרומפטים...")
    if os.path.exists("Cursor_Prompts"):
        shutil.copytree("Cursor_Prompts", f"{backup_dir}/Cursor_Prompts", dirs_exist_ok=True)
        print("   ✅ Cursor_Prompts/")
    
    # 6. גיבוי תוסף Outlook
    print("\n6️⃣ מגבה תוסף Outlook...")
    if os.path.exists("outlook_addin"):
        shutil.copytree("outlook_addin", f"{backup_dir}/outlook_addin", dirs_exist_ok=True)
        print("   ✅ outlook_addin/")
    
    # 7. גיבוי קבצי התקנה
    print("\n7️⃣ מגבה קבצי התקנה...")
    install_files = [
        "install.bat",
        "install_final_com_addin.bat",
        "install_final_simple.bat",
        "install_office_addin.bat",
        "requirements.txt"
    ]
    for file in install_files:
        if os.path.exists(file):
            shutil.copy2(file, f"{backup_dir}/{file}")
            print(f"   ✅ {file}")
    
    # 8. יצירת קובץ ZIP
    print("\n8️⃣ יוצר קובץ ZIP...")
    zip_filename = f"backup_{timestamp}.zip"
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(backup_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, backup_dir)
                zipf.write(file_path, arcname)
    print(f"   ✅ {zip_filename}")
    
    # 9. סטטיסטיקות גיבוי
    print("\n9️⃣ סטטיסטיקות:")
    
    # גודל הגיבוי
    zip_size = os.path.getsize(zip_filename) / (1024 * 1024)  # MB
    print(f"   📦 גודל הגיבוי: {zip_size:.2f} MB")
    
    # ספירת קבצים
    total_files = sum([len(files) for _, _, files in os.walk(backup_dir)])
    print(f"   📄 מספר קבצים: {total_files}")
    
    # סטטיסטיקות בסיס נתונים
    if os.path.exists("email_manager.db"):
        conn = sqlite3.connect("email_manager.db")
        cursor = conn.cursor()
        
        try:
            cursor.execute("SELECT COUNT(*) FROM email_ai_analysis")
            emails_analyzed = cursor.fetchone()[0]
            print(f"   📧 מיילים מנותחים: {emails_analyzed}")
        except:
            pass
        
        try:
            cursor.execute("SELECT COUNT(*) FROM meeting_ai_analysis")
            meetings_analyzed = cursor.fetchone()[0]
            print(f"   📅 פגישות מנותחות: {meetings_analyzed}")
        except:
            pass
        
        conn.close()
    
    # 10. סיכום
    print("\n" + "="*60)
    print("✅ הגיבוי הושלם בהצלחה!")
    print("="*60)
    print(f"📁 תיקייה: {backup_dir}/")
    print(f"📦 ZIP: {zip_filename}")
    print(f"💾 מיקום: {os.path.abspath(zip_filename)}")
    print("="*60)
    
    return zip_filename, backup_dir

if __name__ == "__main__":
    create_backup()

