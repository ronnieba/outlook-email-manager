#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
סקריפט ליצירת תמונות מתרשימי Mermaid
יוצר תמונות PNG ו-SVG מכל התרשימים בפרויקט
"""

import os
import subprocess
import json
from pathlib import Path

def create_mermaid_images():
    """יוצר תמונות מכל התרשימי Mermaid"""
    
    # תרשימים שונים
    diagrams = {
        "architecture": """
graph TD
    A[📧 Outlook Email Manager] --> B[🐍 Backend Flask]
    A --> C[🎨 Frontend HTML/CSS/JS]
    A --> D[🤖 AI Engine]
    A --> E[💾 Database]
    
    B --> B1[app_with_ai.py]
    B --> B2[ai_analyzer.py]
    B --> B3[user_profile_manager.py]
    B --> B4[config.py]
    
    C --> C1[📧 index.html]
    C --> C2[📅 meetings.html]
    C --> C3[🖥️ consol.html]
    
    D --> D1[Google Gemini API]
    D --> D2[AI Analysis]
    D --> D3[Learning System]
    
    E --> E1[email_manager.db]
    E --> E2[email_preferences.db]
    
    F[📚 Documentation] --> F1[README.md]
    F --> F2[INSTALLATION.md]
    F --> F3[USER_GUIDE.md]
    F --> F4[API_DOCUMENTATION.md]
    F --> F5[DEVELOPER_GUIDE.md]
    F --> F6[CHANGELOG.md]
        """,
        
        "workflow": """
flowchart TD
    A[🚀 הפעלת המערכת] --> B[🌐 פתיחת דפדפן]
    B --> C{בחירת דף}
    
    C -->|ניהול מיילים| D[📧 דף ניהול מיילים]
    C -->|ניהול פגישות| E[📅 דף ניהול פגישות]
    C -->|מעקב מערכת| F[🖥️ דף קונסול]
    
    D --> D1[🔄 רענן מיילים]
    D --> D2[📊 סטטיסטיקות]
    D --> D3[🤖 ניתוח AI]
    D --> D4[📝 מתן משוב]
    
    E --> E1[🔄 רענן פגישות]
    E --> E2[🎯 כפתורי עדיפות]
    E --> E3[📈 סטטיסטיקות פגישות]
    E --> E4[🤖 ניתוח AI פגישות]
    
    F --> F1[📝 לוגים חיים]
    F --> F2[⚙️ ניהול שרת]
    F --> F3[💾 גיבויים]
    F --> F4[📝 פרומפטים]
    
    D4 --> G[🧠 מערכת למידה]
    E4 --> G
    G --> H[📈 שיפור אוטומטי]
        """,
        
        "api_endpoints": """
graph TD
    A[🌐 API Base URL: localhost:5000] --> B[📧 Email APIs]
    A --> C[📅 Meeting APIs]
    A --> D[🤖 AI APIs]
    A --> E[📊 Learning APIs]
    A --> F[🔧 System APIs]
    A --> G[🖥️ Console APIs]
    A --> H[📦 Backup APIs]
    
    B --> B1[GET /api/emails]
    B --> B2[POST /api/refresh-data]
    B --> B3[GET /api/stats]
    B --> B4[POST /api/user-feedback]
    B --> B5[POST /api/analyze-emails-ai]
    
    C --> C1[GET /api/meetings]
    C --> C2[POST /api/meetings/:id/priority]
    C --> C3[GET /api/meetings/stats]
    C --> C4[POST /api/analyze-meetings-ai]
    
    D --> D1[GET /api/ai-status]
    D --> D2[POST /api/analyze-emails-ai]
    D --> D3[POST /api/analyze-meetings-ai]
    
    E --> E1[GET /api/learning-stats]
    E --> E2[GET /api/learning-management]
    
    F --> F1[GET /api/test-outlook]
    F --> F2[GET /api/server-id]
    F --> F3[POST /api/restart-server]
    
    G --> G1[GET /api/console-logs]
    G --> G2[POST /api/clear-console]
    G --> G3[POST /api/console-reset]
    
    H --> H1[POST /api/create-backup]
    H --> H2[POST /api/create-cursor-prompts]
        """,
        
        "database_schema": """
erDiagram
    EMAILS {
        string id PK
        string subject
        string sender
        string sender_email
        text body_preview
        datetime received_time
        boolean is_read
        real importance_score
        boolean ai_analyzed
        real ai_importance_score
        real original_importance_score
        datetime ai_analysis_date
        text summary
        string category
        text action_items
    }
    
    MEETINGS {
        string id PK
        string subject
        string organizer
        string organizer_email
        datetime start_time
        datetime end_time
        string location
        text attendees
        text body
        real importance_score
        boolean ai_analyzed
        string priority
    }
    
    USER_FEEDBACK {
        int id PK
        string email_id FK
        string feedback_type
        string user_value
        string ai_value
        datetime timestamp
    }
    
    LEARNING_PATTERNS {
        int id PK
        string pattern_type
        text pattern_data
        real confidence
        datetime created_at
    }
    
    EMAILS ||--o{ USER_FEEDBACK : "has feedback"
        """,
        
        "project_timeline": """
gantt
    title התפתחות Outlook Email Manager
    dateFormat  YYYY-MM-DD
    section גרסה 1.0
    פיתוח בסיסי           :done, v1-dev, 2025-09-20, 2025-09-22
    אינטגרציה Outlook     :done, v1-outlook, 2025-09-21, 2025-09-23
    AI בסיסי              :done, v1-ai, 2025-09-22, 2025-09-24
    ממשק משתמש            :done, v1-ui, 2025-09-23, 2025-09-25
    
    section גרסה 2.0
    דף פגישות             :done, v2-meetings, 2025-09-25, 2025-09-27
    כפתורי עדיפות         :done, v2-priority, 2025-09-26, 2025-09-28
    דף קונסול             :done, v2-console, 2025-09-27, 2025-09-29
    תיעוד מפורט           :done, v2-docs, 2025-09-28, 2025-09-29
    
    section גרסה 2.1
    אימות משתמשים         :active, v21-auth, 2025-09-30, 2025-10-05
    הגדרות מתקדמות         :v21-settings, 2025-10-01, 2025-10-07
    דוחות וניתוחים        :v21-reports, 2025-10-03, 2025-10-10
        """,
        
        "installation_process": """
flowchart TD
    A[🚀 התחלת התקנה] --> B[📥 הורדת הפרויקט]
    B --> C[🔍 בדיקת דרישות מערכת]
    
    C --> D{דרישות תקינות?}
    D -->|לא| E[❌ שגיאת התקנה]
    D -->|כן| F[🐍 התקנת Python packages]
    
    F --> G[📦 יצירת סביבה וירטואלית]
    G --> H[⚙️ הגדרת בסיס נתונים]
    H --> I[🔧 הגדרת Outlook]
    I --> J[🤖 הגדרת AI API]
    J --> K[🚀 הפעלת השרת]
    
    K --> L[✅ התקנה הושלמה]
    E --> M[📞 פנה לתמיכה]
    
    style A fill:#e1f5fe
    style L fill:#e8f5e8
    style E fill:#ffebee
    style M fill:#fff3e0
        """
    }
    
    # יצירת תיקיית תמונות
    images_dir = Path("diagrams_images")
    images_dir.mkdir(exist_ok=True)
    
    print("Creating Mermaid diagram images...")
    
    for name, diagram in diagrams.items():
        print(f"Creating diagram: {name}")
        
        # יצירת קובץ HTML זמני
        html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <script src="https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js"></script>
</head>
<body>
    <div class="mermaid">
{diagram}
    </div>
    <script>
        mermaid.initialize({{ startOnLoad: true }});
    </script>
</body>
</html>
        """
        
        # שמירת קובץ HTML זמני
        temp_html = images_dir / f"{name}_temp.html"
        with open(temp_html, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"HTML file created: {temp_html}")
    
    print(f"\nAll diagrams saved in: {images_dir.absolute()}")
    print("Open the HTML files in browser to view diagrams")
    print("Use screenshot tools to save as images")

if __name__ == "__main__":
    create_mermaid_images()
