#!/usr/bin/env python3
"""
Test script for Outlook Email Manager
סקריפט בדיקה למערכת ניהול מיילים חכמה
"""

import sys
import os
import sqlite3
from datetime import datetime

def test_imports():
    """בדיקת ייבוא מודולים"""
    print("🔍 Testing imports...")
    
    try:
        import flask
        print("✅ Flask imported successfully")
    except ImportError as e:
        print(f"❌ Flask import failed: {e}")
        return False
    
    try:
        import win32com.client
        print("✅ pywin32 imported successfully")
    except ImportError as e:
        print(f"❌ pywin32 import failed: {e}")
        return False
    
    try:
        import google.generativeai
        print("✅ google-generativeai imported successfully")
    except ImportError as e:
        print(f"❌ google-generativeai import failed: {e}")
        return False
    
    return True

def test_database():
    """בדיקת מסד נתונים"""
    print("\n🔍 Testing database...")
    
    try:
        conn = sqlite3.connect("test_email_preferences.db")
        cursor = conn.cursor()
        
        # בדיקת יצירת טבלאות
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS test_table (
                id INTEGER PRIMARY KEY,
                test_value TEXT
            )
        ''')
        
        cursor.execute("INSERT INTO test_table (test_value) VALUES (?)", ("test",))
        cursor.execute("SELECT * FROM test_table")
        result = cursor.fetchone()
        
        if result and result[1] == "test":
            print("✅ Database operations working")
        else:
            print("❌ Database operations failed")
            return False
        
        conn.close()
        os.remove("test_email_preferences.db")
        return True
        
    except Exception as e:
        print(f"❌ Database test failed: {e}")
        return False

def test_ai_analyzer():
    """בדיקת מודול ניתוח AI"""
    print("\n🔍 Testing AI Analyzer...")
    
    try:
        from ai_analyzer import EmailAnalyzer
        analyzer = EmailAnalyzer()
        
        # בדיקת ניתוח בסיסי
        test_email = {
            'subject': 'Test email',
            'sender': 'test@example.com',
            'body_preview': 'This is a test email'
        }
        
        importance = analyzer.calculate_basic_importance(test_email)
        if 0 <= importance <= 1:
            print("✅ AI Analyzer basic functions working")
            return True
        else:
            print("❌ AI Analyzer importance score out of range")
            return False
            
    except Exception as e:
        print(f"❌ AI Analyzer test failed: {e}")
        return False

def test_user_profile_manager():
    """בדיקת מודול ניהול פרופיל"""
    print("\n🔍 Testing User Profile Manager...")
    
    try:
        from user_profile_manager import UserProfileManager
        profile_manager = UserProfileManager("test_profile.db")
        
        # בדיקת פונקציות בסיסיות
        test_email = {
            'id': 1,
            'subject': 'Test email',
            'sender': 'test@example.com'
        }
        
        importance = profile_manager.get_personalized_importance_score(test_email)
        if 0 <= importance <= 1:
            print("✅ User Profile Manager working")
            os.remove("test_profile.db")
            return True
        else:
            print("❌ User Profile Manager importance score out of range")
            return False
            
    except Exception as e:
        print(f"❌ User Profile Manager test failed: {e}")
        return False

def test_config():
    """בדיקת קובץ הגדרות"""
    print("\n🔍 Testing configuration...")
    
    try:
        from config import GEMINI_API_KEY
        if GEMINI_API_KEY and GEMINI_API_KEY != 'your_api_key_here':
            print("✅ Gemini API Key configured")
        else:
            print("⚠️ Gemini API Key not configured (AI features limited)")
        return True
        
    except Exception as e:
        print(f"❌ Configuration test failed: {e}")
        return False

def test_outlook_connection():
    """בדיקת חיבור ל-Outlook"""
    print("\n🔍 Testing Outlook connection...")
    
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        print("✅ Outlook connection successful")
        return True
        
    except Exception as e:
        print(f"⚠️ Outlook connection failed: {e}")
        print("   This is normal if Outlook is not running")
        return False

def main():
    """פונקציה ראשית"""
    print("🚀 Outlook Email Manager - Test Script")
    print("=====================================")
    print(f"📅 Test started at: {datetime.now()}")
    print()
    
    tests = [
        ("Imports", test_imports),
        ("Database", test_database),
        ("AI Analyzer", test_ai_analyzer),
        ("User Profile Manager", test_user_profile_manager),
        ("Configuration", test_config),
        ("Outlook Connection", test_outlook_connection)
    ]
    
    passed = 0
    total = len(tests)
    
    for test_name, test_func in tests:
        try:
            if test_func():
                passed += 1
        except Exception as e:
            print(f"❌ {test_name} test crashed: {e}")
    
    print("\n" + "="*50)
    print(f"📊 Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All tests passed! Project is ready to run.")
        return True
    elif passed >= total - 1:  # Allow Outlook to fail
        print("✅ Project is ready to run (Outlook optional).")
        return True
    else:
        print("❌ Some critical tests failed. Please fix issues before running.")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
