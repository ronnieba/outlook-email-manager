"""
🔍 Verify Installation Script
סקריפט בדיקה לוודא שכל דרישות המערכת מותקנות
"""
import sys
import os
import subprocess
import platform

def print_header(text):
    """הדפסת כותרת מעוצבת"""
    print(f"\n{'='*60}")
    print(f"  {text}")
    print(f"{'='*60}\n")

def print_status(check_name, status, message=""):
    """הדפסת סטטוס בדיקה"""
    status_icon = "✅" if status else "❌"
    print(f"{status_icon} {check_name:.<50} {'OK' if status else 'FAILED'}")
    if message:
        print(f"   ℹ️  {message}")

def check_python_version():
    """בדיקת גרסת Python"""
    print_header("🐍 בדיקת Python")
    version = sys.version_info
    is_valid = version.major == 3 and version.minor >= 8
    print_status(
        "Python Version",
        is_valid,
        f"Found: Python {version.major}.{version.minor}.{version.micro}"
    )
    if not is_valid:
        print("   ⚠️  נדרש Python 3.8 ומעלה!")
    return is_valid

def check_windows():
    """בדיקה שזה Windows"""
    print_header("💻 בדיקת מערכת הפעלה")
    is_windows = platform.system() == "Windows"
    print_status(
        "Operating System",
        is_windows,
        f"Found: {platform.system()} {platform.release()}"
    )
    if not is_windows:
        print("   ⚠️  המערכת תומכת רק ב-Windows!")
    return is_windows

def check_outlook():
    """בדיקה ש-Outlook מותקן"""
    print_header("📧 בדיקת Microsoft Outlook")
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        version = outlook.Version
        print_status(
            "Microsoft Outlook",
            True,
            f"Found: Outlook {version}"
        )
        return True
    except Exception as e:
        print_status(
            "Microsoft Outlook",
            False,
            f"Error: {str(e)}"
        )
        print("   ⚠️  וודא ש-Outlook מותקן ופתוח!")
        return False

def check_requirements():
    """בדיקת חבילות Python נדרשות"""
    print_header("📦 בדיקת תלויות Python")
    
    required_packages = {
        'flask': 'Flask',
        'flask_cors': 'flask-cors',
        'win32com.client': 'pywin32',
        'google.generativeai': 'google-generativeai'
    }
    
    all_installed = True
    
    for module_name, package_name in required_packages.items():
        try:
            __import__(module_name)
            print_status(f"{package_name}", True)
        except ImportError:
            print_status(
                f"{package_name}",
                False,
                f"Run: pip install {package_name}"
            )
            all_installed = False
    
    return all_installed

def check_config_file():
    """בדיקת קובץ config.py"""
    print_header("⚙️ בדיקת קונפיגורציה")
    
    if not os.path.exists('config.py'):
        print_status(
            "config.py",
            False,
            "File not found!"
        )
        return False
    
    try:
        import config
        
        # בדיקת API Key
        has_api_key = hasattr(config, 'GEMINI_API_KEY') and \
                      config.GEMINI_API_KEY and \
                      config.GEMINI_API_KEY != 'your-api-key-here'
        
        print_status(
            "config.py exists",
            True
        )
        print_status(
            "GEMINI_API_KEY configured",
            has_api_key,
            "Set your API key in config.py or .env" if not has_api_key else ""
        )
        
        return has_api_key
        
    except Exception as e:
        print_status(
            "config.py",
            False,
            f"Error loading: {str(e)}"
        )
        return False

def check_database_files():
    """בדיקת קבצי מסד נתונים"""
    print_header("🗄️ בדיקת מסדי נתונים")
    
    db_files = [
        ('email_manager.db', 'נוצר אוטומטית בהרצה ראשונה'),
        ('email_preferences.db', 'נוצר אוטומטית בהרצה ראשונה')
    ]
    
    for db_file, note in db_files:
        exists = os.path.exists(db_file)
        print_status(
            db_file,
            True,  # תמיד True כי הם נוצרים אוטומטית
            f"{'Found' if exists else note}"
        )
    
    return True

def check_main_files():
    """בדיקת קבצים עיקריים"""
    print_header("📄 בדיקת קבצים עיקריים")
    
    required_files = [
        'app_with_ai.py',
        'working_email_analyzer.py',
        'outlook_com_addin_final.py',
        'requirements.txt',
        'config.py'
    ]
    
    all_exist = True
    
    for file in required_files:
        exists = os.path.exists(file)
        print_status(file, exists)
        if not exists:
            all_exist = False
    
    return all_exist

def check_templates():
    """בדיקת תיקיית templates"""
    print_header("🎨 בדיקת Templates")
    
    if not os.path.exists('templates'):
        print_status(
            "templates/ directory",
            False,
            "Directory not found!"
        )
        return False
    
    print_status("templates/ directory", True)
    
    # ספירת קבצי HTML
    html_files = [f for f in os.listdir('templates') if f.endswith('.html')]
    print(f"   ℹ️  Found {len(html_files)} HTML files")
    
    return True

def check_server_port():
    """בדיקה שפורט 5000 פנוי"""
    print_header("🌐 בדיקת פורט השרת")
    
    import socket
    
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        result = sock.connect_ex(('localhost', 5000))
        sock.close()
        
        if result == 0:
            print_status(
                "Port 5000",
                False,
                "Port is already in use. Server might be running or change port."
            )
            return False
        else:
            print_status(
                "Port 5000",
                True,
                "Port is available"
            )
            return True
    except Exception as e:
        print_status(
            "Port 5000",
            False,
            f"Error checking port: {str(e)}"
        )
        return False

def main():
    """הרצת כל הבדיקות"""
    print("\n" + "="*60)
    print("  🔍 Outlook Email Manager - Installation Verification")
    print("  📧 בדיקת התקנה מערכת ניהול מיילים חכמה")
    print("="*60)
    
    checks = [
        ("Python Version", check_python_version),
        ("Operating System", check_windows),
        ("Main Files", check_main_files),
        ("Python Packages", check_requirements),
        ("Microsoft Outlook", check_outlook),
        ("Configuration", check_config_file),
        ("Database Files", check_database_files),
        ("Templates", check_templates),
        ("Server Port", check_server_port)
    ]
    
    results = []
    
    for check_name, check_func in checks:
        try:
            results.append(check_func())
        except Exception as e:
            print(f"\n❌ Error in {check_name}: {str(e)}")
            results.append(False)
    
    # סיכום
    print_header("📊 סיכום בדיקות")
    
    passed = sum(results)
    total = len(results)
    percentage = (passed / total) * 100
    
    print(f"✅ עברו: {passed}/{total} ({percentage:.0f}%)")
    
    if passed == total:
        print("\n🎉 מצוין! כל הבדיקות עברו בהצלחה!")
        print("✅ אתה מוכן להפעיל את המערכת:")
        print("   python app_with_ai.py")
        print("   ואז פתח דפדפן ב-http://localhost:5000")
    elif passed >= total * 0.7:
        print("\n⚠️  רוב הבדיקות עברו, אבל יש כמה בעיות.")
        print("   תקן את הבעיות שלמעלה והרץ שוב.")
    else:
        print("\n❌ נמצאו בעיות רבות!")
        print("   עקוב אחר הוראות ההתקנה ב-INSTALLATION_GUIDE_SIMPLE.md")
    
    print("\n" + "="*60 + "\n")
    
    return passed == total

if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⚠️  בדיקה בוטלה על ידי המשתמש.")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ שגיאה כללית: {str(e)}")
        sys.exit(1)



