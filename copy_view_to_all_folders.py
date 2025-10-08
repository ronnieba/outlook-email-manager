# -*- coding: utf-8 -*-
"""
סקריפט פשוט להעתקת התצוגה עם AISCORE לכל התיקיות
"""

import win32com.client
import pythoncom

def main():
    print("=" * 70)
    print("העתקת תצוגה עם עמודת AISCORE לכל תיקיות המייל")
    print("=" * 70)
    print()
    
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # קבלת תיקיית INBOX
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        print(f"✓ נמצאה תיקיית INBOX: {inbox.Name}")
        
        # קבלת התצוגה הנוכחית של INBOX
        inbox_view = inbox.CurrentView
        print(f"✓ התצוגה הנוכחית: {inbox_view.Name}")
        print()
        
        # בדיקה אם יש עמודת AISCORE
        has_aiscore = False
        if inbox_view.ViewType == 0:  # Table view
            for field in inbox_view.ViewFields:
                field_name = str(field.ViewXMLSchemaName).upper()
                if "AISCORE" in field_name or "AI_SCORE" in field_name:
                    has_aiscore = True
                    print(f"✓ נמצאה עמודת AISCORE בתצוגה של INBOX")
                    break
        
        if not has_aiscore:
            print("⚠ לא נמצאה עמודת AISCORE ב-INBOX")
            print("אנא הוסף את העמודה ל-INBOX קודם")
            return
        
        print()
        print("מעתיק את התצוגה לתיקיות אחרות...")
        print("-" * 70)
        
        # פונקציה רקורסיבית לעיבוד תיקיות
        def process_folder(folder, level=0):
            indent = "  " * level
            try:
                # רק תיקיות מייל
                if folder.DefaultItemType == 0:  # Mail items
                    try:
                        # החלפת התצוגה הנוכחית
                        current_view = folder.CurrentView
                        
                        # רק אם זו תצוגת טבלה
                        if current_view.ViewType == 0:
                            # העתקת השדות מ-INBOX
                            # מחיקת כל השדות הקיימים
                            while current_view.ViewFields.Count > 0:
                                current_view.ViewFields.Remove(1)
                            
                            # העתקת השדות מ-INBOX
                            for field in inbox_view.ViewFields:
                                try:
                                    current_view.ViewFields.Add(field.ViewXMLSchemaName)
                                except:
                                    pass
                            
                            current_view.Save()
                            print(f"{indent}✓ {folder.Name}")
                    except Exception as e:
                        print(f"{indent}⚠ {folder.Name}: {e}")
                
                # עיבוד תיקיות משנה
                for subfolder in folder.Folders:
                    process_folder(subfolder, level + 1)
                    
            except Exception as e:
                print(f"{indent}⚠ שגיאה: {e}")
        
        # עיבוד כל החשבונות
        for store in namespace.Stores:
            try:
                root_folder = store.GetRootFolder()
                print(f"\n📁 חשבון: {store.DisplayName}")
                process_folder(root_folder, 1)
            except Exception as e:
                print(f"⚠ שגיאה בחשבון: {e}")
        
        print()
        print("=" * 70)
        print("✓ הסתיים בהצלחה!")
        print("=" * 70)
        print()
        print("כעת:")
        print("1. עבור לתיקייה TEST")
        print("2. לחץ F5 לרענון")
        print("3. העמודה AISCORE אמורה להופיע!")
        
    except Exception as e:
        print(f"❌ שגיאה: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
    print()
    input("לחץ Enter לסגירה...")
