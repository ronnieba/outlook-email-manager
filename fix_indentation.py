#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Script to fix indentation issues in app_with_ai.py

def fix_file():
    with open('app_with_ai.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Fix the specific indentation issues
    lines = content.split('\n')
    
    # Fix lines 147-151
    if len(lines) > 146:
        if lines[146].strip() == 'try:':
            # Fix the indentation for the try block
            lines[147] = '                messages = self.inbox.Items'
            lines[148] = '                emails = []'
            lines[149] = '                '
            lines[150] = '                print(f"📧 נמצאו {messages.Count} מיילים ב-Outlook")'
            lines[151] = '                log_to_console(f"📧 נמצאו {messages.Count} מיילים ב-Outlook", "INFO")'
    
    # Fix lines 187-188
    if len(lines) > 186:
        if 'email_data[\'summary\']' in lines[187]:
            lines[187] = '                    email_data[\'summary\'] = f"מייל מ-{email_data[\'sender\']}: {email_data[\'subject\']}"'
            lines[188] = '                    email_data[\'action_items\'] = []'
    
    # Write back to file
    with open('app_with_ai.py', 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    
    print("Fixed indentation issues!")

if __name__ == "__main__":
    fix_file()


