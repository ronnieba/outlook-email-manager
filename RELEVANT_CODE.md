# קוד רלוונטי לבעיית ההודעות החוזרות

## 1. API console-logs (app_with_ai.py שורות 1753-1757)
```python
@app.route('/api/console-logs')
def get_console_logs():
    """API לקבלת לוגים מהקונסול"""
    # מחזיר את כל הלוגים (עד 50)
    return jsonify(all_console_logs)
```

## 2. פונקציה refreshLogs (templates/consol.html שורות 1636-1686)
```javascript
// מערך לשמירת הלוגים שכבר נוספו (hash של התוכן)
window.addedLogsSet = new Set();

function refreshLogs() {
    // הגבלה על תדירות הרענון כדי למנוע תקיעות
    const now = Date.now();
    if (window.lastRefreshTime && (now - window.lastRefreshTime) < 5000) {
        return; // לא לרענן יותר מפעם ב-5 שניות
    }
    window.lastRefreshTime = now;
    
    // לא להציג הודעה בקונסול - רק לטעון לוגים בשקט
    fetch(`/api/console-logs?t=${now}`)
        .then(response => {
            if (!response.ok) {
                throw new Error('השרת לא זמין');
            }
            return response.json();
        })
        .then(logs => {
            // הוספת רק לוגים חדשים
            let addedCount = 0;
            
            logs.forEach(log => {
                let logText = '';
                if (typeof log === 'object' && log.message) {
                    logText = log.message;
                } else {
                    logText = log;
                }
                
                // יצירת hash של הלוג (ללא זמן)
                const cleanLogText = logText.replace(/\[\d+:\d+:\d+\]\s*\d+:\s*/, '').trim();
                
                // בדיקה אם הלוג כבר נוסף
                if (!window.addedLogsSet.has(cleanLogText)) {
                    addLogEntry(logText);
                    window.addedLogsSet.add(cleanLogText);
                    addedCount++;
                }
            });

            // הצגת הודעה שקופצת רק אם נוספו לוגים חדשים
            if (addedCount > 0) {
                showRefreshStatus(`נטענו ${addedCount} לוגים חדשים`);
            }
        })
        .catch(error => {
            showStatus('שגיאה בטעינת לוגים: ' + error.message, 'error');
        });
}
```

## 3. API transfer-scores-to-outlook (app_with_ai.py שורות 3178-3302)
```python
@app.route('/api/transfer-scores-to-outlook', methods=['POST'])
def transfer_scores_to_outlook():
    """API להעברת ציונים ל-Outlook"""
    try:
        log_to_console("🚀 מתחיל העברת ציונים ל-Outlook...", "INFO")
        
        # בדיקה שיש נתונים זמינים
        if not cached_data['emails']:
            log_to_console("❌ אין מיילים זמינים להעברה", "ERROR")
            return jsonify({
                'success': False,
                'message': 'אין מיילים זמינים להעברה. נא לטעון את המיילים קודם.'
            }), 400
        
        emails_processed = 0
        emails_success = 0
        emails_failed = 0
        
        log_to_console(f"📧 נמצאו {len(cached_data['emails'])} מיילים עם ציונים מוכנים", "INFO")
        
        # עיבוד המיילים (כל המיילים)
        max_emails = len(cached_data['emails'])
        
        log_to_console(f"⚡ מעבד {max_emails} מיילים (כל המיילים)", "INFO")
        
        # בדיקת חיבור ל-Outlook
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            log_to_console("✅ חיבור ל-Outlook הצליח!", "SUCCESS")
        except Exception as e:
            log_to_console(f"❌ שגיאה בחיבור ל-Outlook: {e}", "ERROR")
            return jsonify({'success': False, 'error': str(e)})
        if not outlook:
            log_to_console("❌ לא ניתן להתחבר ל-Outlook", "ERROR")
            return jsonify({
                'success': False,
                'message': 'לא ניתן להתחבר ל-Outlook'
            }), 500
        
        log_to_console("✅ חיבור ל-Outlook הצליח!", "SUCCESS")
        
        # קבלת כל המיילים מ-Outlook
        try:
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)  # מיון לפי זמן קבלה
            
            log_to_console(f"📧 נמצאו {messages.Count} מיילים ב-Outlook", "INFO")
            
            for i in range(max_emails):
                # ... קוד עיבוד המיילים ...
                log_to_console(f"✅ מייל {i+1}: {email_subject} - ציון: {score}%", "SUCCESS")
```

## 4. קריאה ל-transfer-scores-to-outlook (templates/consol.html שורות 1900-1920)
```javascript
if (confirm('האם אתה בטוח שברצונך להעביר את כל הציונים ל-Outlook?\nזה יעביר את כל הציונים מהאפליקציה למיילים ב-Outlook.')) {
    showStatus('מעביר ציונים ל-Outlook...', 'info');
    addLogEntry('📊 מתחיל העברת ציונים ל-Outlook...');
    
    fetch('/api/transfer-scores-to-outlook', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        }
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('השרת לא זמין');
        }
        return response.json();
    })
    .then(data => {
        if (data.success) {
            addLogEntry('✅ ציונים הועברו ל-Outlook בהצלחה!');
            addLogEntry(`📧 מיילים שעובדו: ${data.emails_processed}`);
```

## 5. פונקציה log_to_console (app_with_ai.py שורות 58-87)
```python
def log_to_console(message, level="INFO"):
    """הוספת הודעה לקונסול"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    
    # ניקוי המילים באנגלית מההודעה לפני שמירה
    clean_message = message
    if level == "INFO" and message.startswith("INFO: "):
        clean_message = message[6:]  # הסרת "INFO: "
    elif level == "SUCCESS" and message.startswith("SUCCESS: "):
        clean_message = message[9:]  # הסרת "SUCCESS: "
    elif level == "ERROR" and message.startswith("ERROR: "):
        clean_message = message[7:]  # הסרת "ERROR: "
    elif level == "WARNING" and message.startswith("WARNING: "):
        clean_message = message[9:]  # הסרת "WARNING: "
    
    # ניקוי תווים בעייתיים לפני הדפסה
    safe_message = clean_message.encode('ascii', errors='ignore').decode('ascii')
    
    log_entry = {
        'message': clean_message,  # שמירת ההודעה הנקייה לרשימה
        'level': level,
        'timestamp': timestamp
    }
    all_console_logs.append(log_entry)
    
    # הדפסה עם הודעה נקייה
    print(f"[{timestamp}] {safe_message}")
```

## 6. הגדרת all_console_logs
```python
# רשימה גלובלית לשמירת כל הלוגים
all_console_logs = []
```
