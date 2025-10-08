# -*- coding: utf-8 -*-
"""
סקריפט להוספת עמודת AISCORE לכל תיקיות Outlook
"""

import win32com.client
import pythoncom

def add_aiscore_column_to_folder(folder, indent=0):
    """הוספת עמודת AISCORE לתיקייה"""
    try:
        folder_name = folder.Name
        print("  " * indent + f"מעבד תיקייה: {folder_name}")
        
        # רק תיקיות מייל
        if folder.DefaultItemType == 0:  # olMailItem
            try:
                # קבלת התצוגה הנוכחית
                current_view = folder.CurrentView
                
                # בדיקה אם זו תצוגת טבלה
                if current_view.ViewType == 0:  # olTableView
                    # בדיקה אם העמודה כבר קיימת
                    column_exists = False
                    for field in current_view.ViewFields:
                        if "AISCORE" in field.ViewXMLSchemaName.upper():
                            column_exists = True
                            break
                    
                    if not column_exists:
                        # הוספת העמודה
                        try:
                            # חיפוש השדה AISCORE
                            user_properties = folder.Items[1].UserProperties
                            aiscore_field = None
                            
                            for prop in user_properties:
                                if prop.Name == "AISCORE":
                                    aiscore_field = prop
                                    break
                            
                            if aiscore_field:
                                # הוספת העמודה לתצוגה
                                current_view.ViewFields.Add("AISCORE")
                                current_view.Save()
                                print("  " * indent + f"  ✓ העמודה AISCORE נוספה ל-{folder_name}")
                            else:
                                print("  " * indent + f"  ⚠ לא נמצא שדה AISCORE במיילים בתיקייה {folder_name}")
                        except Exception as e:
                            print("  " * indent + f"  ⚠ לא ניתן להוסיף עמודה ל-{folder_name}: {e}")
                    else:
                        print("  " * indent + f"  ℹ העמודה כבר קיימת ב-{folder_name}")
                        
            except Exception as e:
                print("  " * indent + f"  ⚠ שגיאה בעיבוד תצוגה של {folder_name}: {e}")
        
        # עיבוד תיקיות משנה
        if folder.Folders.Count > 0:
            for subfolder in folder.Folders:
                add_aiscore_column_to_folder(subfolder, indent + 1)
                
    except Exception as e:
        print("  " * indent + f"שגיאה בעיבוד תיקייה: {e}")

def main():
    """פונקציה ראשית"""
    print("=" * 60)
    print("הוספת עמודת AISCORE לכל תיקיות Outlook")
    print("=" * 60)
    print()
    
    try:
        # חיבור ל-Outlook
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("✓ מחובר ל-Outlook")
        print()
        
        # עיבוד כל התיקיות הראשיות
        for i in range(1, namespace.Folders.Count + 1):
            try:
                folder = namespace.Folders.Item(i)
                print(f"\nמעבד חשבון: {folder.Name}")
                print("-" * 60)
                add_aiscore_column_to_folder(folder)
            except Exception as e:
                print(f"שגיאה בעיבוד חשבון: {e}")
        
        print()
        print("=" * 60)
        print("✓ הסתיים!")
        print("=" * 60)
        print()
        print("עכשיו רענן את Outlook (F5) כדי לראות את השינויים")
        
    except Exception as e:
        print(f"שגיאה: {e}")
    
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
    input("\nלחץ Enter לסגירה...")
