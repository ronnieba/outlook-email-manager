# -*- coding: utf-8 -*-
"""
בדיקת כמות מיילים ופגישות ב-Outlook
"""

import win32com.client

def count_outlook_items():
    """ספירת מיילים ופגישות"""
    try:
        print("מתחבר ל-Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # ספירת מיילים ב-Inbox
        inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
        emails_count = inbox.Items.Count
        
        # ספירת פגישות ביומן
        calendar = namespace.GetDefaultFolder(9)  # 9 = Calendar
        meetings_count = calendar.Items.Count
        
        print("\n" + "="*60)
        print("📊 סטטיסטיקות Outlook")
        print("="*60)
        print(f"📧 מיילים ב-Inbox: {emails_count}")
        print(f"📅 פגישות ביומן: {meetings_count}")
        print(f"📊 סה״כ פריטים: {emails_count + meetings_count}")
        print("="*60)
        
        return emails_count, meetings_count
        
    except Exception as e:
        print(f"❌ שגיאה: {e}")
        return 0, 0

if __name__ == "__main__":
    count_outlook_items()

