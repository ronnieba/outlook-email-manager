"""
🔒 Complete Backup Script with New Files
יצירת גיבוי מלא כולל כל הקבצים החדשים שנוצרו
"""
import os
import shutil
import zipfile
from datetime import datetime

def create_backup():
    """יצירת גיבוי מלא של הפרויקט"""
    
    # שם הגיבוי עם תאריך ושעה
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_name = f'complete_backup_{timestamp}'
    backup_dir = backup_name
    backup_zip = f'{backup_name}.zip'
    
    print("="*70)
    print(f"  🔒 Creating Complete Backup")
    print(f"  📦 Backup Name: {backup_name}")
    print("="*70)
    print()
    
    # יצירת תיקיית גיבוי
    if os.path.exists(backup_dir):
        shutil.rmtree(backup_dir)
    os.makedirs(backup_dir)
    
    # רשימת קבצים לגיבוי
    files_to_backup = [
        # Core Application Files
        'app_with_ai.py',
        'ai_analyzer.py',
        'user_profile_manager.py',
        'working_email_analyzer.py',
        'outlook_com_addin_final.py',
        'collapsible_logger.py',
        'config.py',
        
        # Configuration & Setup (NEW!)
        'env.example',
        'verify_installation.py',
        'requirements.txt',
        '.gitignore',
        
        # Installation Scripts
        'install_final_simple.bat',
        'install_final_com_addin.bat',
        'install_com_addin.bat',
        
        # Documentation - Root Level
        'README.md',
        'QUICK_START.md',  # NEW!
        'INSTALLATION_GUIDE_SIMPLE.md',
        'SYSTEM_ARCHITECTURE.md',
        'AISCORE_COLUMN_SETUP.md',
        'COM_ADDIN_REGISTRATION_GUIDE.md',  # NEW!
        'PROJECT_COMPLETION_SUMMARY.md',  # NEW!
        'GITHUB_BACKUP_GUIDE.md',
        'VERIFICATION_REPORT.md',
        'TESTING_GUIDE.md',
        'VISUAL_GUIDE.md',
        'AUTO_SYNC_GUIDE.md',
        'FINAL_WORKING_SOLUTION.md',
        'QUICK_START_OUTLOOK_ADDIN.md',
    ]
    
    # תיקיות לגיבוי
    dirs_to_backup = [
        'templates',
        'Cursor_Prompts',
        'docs',
        'outlook_addin',
    ]
    
    # בסיסי נתונים (אופציונלי)
    db_files = [
        'email_manager.db',
        'email_preferences.db',
    ]
    
    # העתקת קבצים בודדים
    print("📄 Copying files...")
    copied_files = 0
    for file in files_to_backup:
        if os.path.exists(file):
            dest = os.path.join(backup_dir, file)
            shutil.copy2(file, dest)
            print(f"  ✅ {file}")
            copied_files += 1
        else:
            print(f"  ⚠️  {file} (not found)")
    
    print()
    
    # העתקת תיקיות
    print("📁 Copying directories...")
    copied_dirs = 0
    for dir_name in dirs_to_backup:
        if os.path.exists(dir_name):
            dest = os.path.join(backup_dir, dir_name)
            shutil.copytree(dir_name, dest)
            file_count = sum([len(files) for _, _, files in os.walk(dir_name)])
            print(f"  ✅ {dir_name}/ ({file_count} files)")
            copied_dirs += 1
        else:
            print(f"  ⚠️  {dir_name}/ (not found)")
    
    print()
    
    # בסיסי נתונים
    print("🗄️  Copying databases...")
    copied_dbs = 0
    for db in db_files:
        if os.path.exists(db):
            dest = os.path.join(backup_dir, db)
            shutil.copy2(db, dest)
            size_mb = os.path.getsize(db) / (1024 * 1024)
            print(f"  ✅ {db} ({size_mb:.2f} MB)")
            copied_dbs += 1
        else:
            print(f"  ⚠️  {db} (not found)")
    
    print()
    
    # יצירת קובץ מידע על הגיבוי
    info_file = os.path.join(backup_dir, 'BACKUP_INFO.txt')
    with open(info_file, 'w', encoding='utf-8') as f:
        f.write(f"# Backup Information\n\n")
        f.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Files copied: {copied_files}\n")
        f.write(f"Directories copied: {copied_dirs}\n")
        f.write(f"Databases copied: {copied_dbs}\n")
        f.write(f"\n## New Files Added:\n")
        f.write(f"- env.example (API Key template)\n")
        f.write(f"- verify_installation.py (Installation checker)\n")
        f.write(f"- COM_ADDIN_REGISTRATION_GUIDE.md (COM Add-in guide)\n")
        f.write(f"- PROJECT_COMPLETION_SUMMARY.md (Completion summary)\n")
        f.write(f"- QUICK_START.md (Quick start guide)\n")
        f.write(f"\n## Updated Files:\n")
        f.write(f"- config.py (Now loads .env)\n")
        f.write(f"- README.md (Enhanced structure)\n")
        f.write(f"- INSTALLATION_GUIDE_SIMPLE.md (Added verify steps)\n")
        f.write(f"- .gitignore (Added .env)\n")
    
    print("📝 Created BACKUP_INFO.txt")
    print()
    
    # דחיסה ל-ZIP
    print("🗜️  Creating ZIP archive...")
    with zipfile.ZipFile(backup_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(backup_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, backup_dir)
                zipf.write(file_path, arcname)
    
    zip_size_mb = os.path.getsize(backup_zip) / (1024 * 1024)
    print(f"  ✅ {backup_zip} ({zip_size_mb:.2f} MB)")
    print()
    
    # ניקוי תיקייה זמנית
    print("🧹 Cleaning up temporary directory...")
    shutil.rmtree(backup_dir)
    print("  ✅ Done")
    print()
    
    # סיכום
    print("="*70)
    print("  🎉 Backup Completed Successfully!")
    print("="*70)
    print()
    print(f"📦 Backup file: {backup_zip}")
    print(f"💾 Size: {zip_size_mb:.2f} MB")
    print(f"📊 Statistics:")
    print(f"   - Files: {copied_files}")
    print(f"   - Directories: {copied_dirs}")
    print(f"   - Databases: {copied_dbs}")
    print()
    print("✅ All new files included:")
    print("   - env.example")
    print("   - verify_installation.py")
    print("   - COM_ADDIN_REGISTRATION_GUIDE.md")
    print("   - PROJECT_COMPLETION_SUMMARY.md")
    print("   - QUICK_START.md")
    print()
    print("✅ Updated files included:")
    print("   - config.py")
    print("   - README.md")
    print("   - INSTALLATION_GUIDE_SIMPLE.md")
    print("   - .gitignore")
    print()
    print("🔒 Your project is safely backed up!")
    print()

if __name__ == '__main__':
    try:
        create_backup()
    except Exception as e:
        print(f"\n❌ Error creating backup: {str(e)}")
        import traceback
        traceback.print_exc()







