// בדיקה אם Office זמין
if (typeof Office === 'undefined') {
    console.log('Office not available, setting up browser mode directly');
    setupBrowserMode();
} else {
    Office.onReady((info) => {
        console.log('Office.onReady called with info:', info);
        if (info.host === Office.HostType.Outlook) {
        console.log('AI Email Manager Add-in loaded successfully');
        
        // הגדרת event listeners
        document.getElementById('analyzeBtn').onclick = analyzeCurrentEmail;
        document.getElementById('profileBtn').onclick = openProfileSettings;
        document.getElementById('statsBtn').onclick = openStatistics;
        document.getElementById('refreshBtn').onclick = refreshEmailInfo;
        
        console.log('Event listeners attached');
        
        // עדכון מידע על המייל הנוכחי
        updateCurrentEmailInfo();
        
        // עדכון אוטומטי כשמשנים מייל
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, updateCurrentEmailInfo);
    } else {
        console.log('Not in Outlook, setting up browser mode');
        
        // הגדרת event listeners גם במצב דפדפן
        document.getElementById('analyzeBtn').onclick = analyzeCurrentEmail;
        document.getElementById('profileBtn').onclick = openProfileSettings;
        document.getElementById('statsBtn').onclick = openStatistics;
        document.getElementById('refreshBtn').onclick = refreshEmailInfo;
        
        console.log('Event listeners attached for browser mode');
        
        // עדכון מידע על המייל הנוכחי
        updateCurrentEmailInfo();
    }
    });
}

function setupBrowserMode() {
    console.log('Setting up browser mode');
    
    // הגדרת event listeners
    document.getElementById('analyzeBtn').onclick = analyzeCurrentEmail;
    document.getElementById('profileBtn').onclick = openProfileSettings;
    document.getElementById('statsBtn').onclick = openStatistics;
    document.getElementById('refreshBtn').onclick = refreshEmailInfo;
    
    console.log('Event listeners attached for browser mode');
    
    // עדכון מידע על המייל הנוכחי
    updateCurrentEmailInfo();
}

function updateCurrentEmailInfo() {
    try {
        // בדיקה אם אנחנו ב-Outlook או בדפדפן
        if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox) {
            // מצב Outlook - טעינת מייל אמיתי
            Office.context.mailbox.item.load(['subject', 'from', 'dateTimeCreated'], (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const item = Office.context.mailbox.item;
                    document.getElementById('emailSubject').textContent = item.subject || 'ללא נושא';
                    document.getElementById('emailSender').textContent = item.from?.displayName || 'שולח לא ידוע';
                    document.getElementById('emailTime').textContent = formatDate(item.dateTimeCreated);
                    
                    // הסתרת תוצאות ניתוח קודמות
                    document.getElementById('analysisResults').style.display = 'none';
                    hideError();
                } else {
                    showError('שגיאה בטעינת פרטי המייל');
                }
            });
        } else {
            // מצב בדיקה בדפדפן - נתונים דמה
            document.getElementById('emailSubject').textContent = 'מייל לדוגמה - בדיקת המערכת';
            document.getElementById('emailSender').textContent = 'משתמש לדוגמה <demo@example.com>';
            document.getElementById('emailTime').textContent = formatDate(new Date());
            
            // הסתרת תוצאות ניתוח קודמות
            document.getElementById('analysisResults').style.display = 'none';
            hideError();
        }
    } catch (error) {
        console.error('Error updating email info:', error);
        showError('שגיאה בחיבור ל-Outlook');
    }
}

function analyzeCurrentEmail() {
    console.log('=== analyzeCurrentEmail called ===');
    console.log('Button clicked successfully!');
    showLoading(true);
    hideError();
    
    try {
        // בדיקה אם אנחנו ב-Outlook או בדפדפן
        if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox) {
            // מצב Outlook - ניתוח מייל אמיתי
            Office.context.mailbox.item.load(['subject', 'body', 'from', 'dateTimeCreated'], (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const item = Office.context.mailbox.item;
                    
                    // קבלת תוכן המייל
                    item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
                        if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
                            const emailData = {
                                subject: item.subject || '',
                                body: bodyResult.value || '',
                                sender: item.from?.emailAddress || '',
                                sender_name: item.from?.displayName || '',
                                date: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : new Date().toISOString()
                            };
                            
                            // שליחה למערכת שלנו
                            sendAnalysisRequest(emailData);
                        } else {
                            showLoading(false);
                            showError('שגיאה בקריאת תוכן המייל');
                        }
                    });
                } else {
                    showLoading(false);
                    showError('שגיאה בטעינת נתוני המייל');
                }
            });
        } else {
            // מצב בדיקה בדפדפן - ניתוח דמה
            const demoEmailData = {
                subject: 'מייל לדוגמה - בדיקת המערכת',
                body: 'זהו מייל לדוגמה לבדיקת המערכת. המייל מכיל מידע חשוב על פרויקט חדש שדורש תשומת לב מיידית.',
                sender: 'demo@example.com',
                sender_name: 'משתמש לדוגמה',
                date: new Date().toISOString()
            };
            
            // שליחה למערכת שלנו עם נתונים דמה
            sendAnalysisRequest(demoEmailData);
        }
    } catch (error) {
        console.error('Error analyzing email:', error);
        showLoading(false);
        showError('שגיאה בניתוח המייל');
    }
}

function sendAnalysisRequest(emailData) {
    console.log('Sending analysis request:', emailData);
    console.log('URL: http://localhost:5000/api/outlook-addin/analyze-email');
    
    fetch('http://localhost:5000/api/outlook-addin/analyze-email', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(emailData)
    })
    .then(response => {
        console.log('Response received:', response);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        console.log('Analysis data received:', data);
        if (data.success) {
            displayAnalysisResults(data);
        } else {
            showError(data.error || 'שגיאה בניתוח המייל');
        }
        showLoading(false);
    })
    .catch(error => {
        console.error('Error:', error);
        showError('שגיאה בחיבור לשרת: ' + error.message);
        showLoading(false);
    });
}

function displayAnalysisResults(data) {
    // הצגת ציון חשיבות
    const scorePercent = Math.round(data.importance_score * 100);
    document.getElementById('importanceScore').textContent = scorePercent + '%';
    
    // הצגת כוכבים
    const stars = generateStars(data.importance_score);
    document.getElementById('importanceStars').innerHTML = stars;
    
    // הצגת קטגוריה
    document.getElementById('categoryBadge').textContent = data.category || 'לא סווג';
    
    // הצגת סיכום
    document.getElementById('summaryText').textContent = data.summary || 'אין סיכום זמין';
    
    // הצגת פעולות נדרשות
    if (data.action_items && data.action_items.length > 0) {
        const actionItemsHtml = data.action_items.map(item => `<li>${item}</li>`).join('');
        document.getElementById('actionItemsList').innerHTML = actionItemsHtml;
    } else {
        document.getElementById('actionItemsList').innerHTML = 'אין פעולות נדרשות';
    }
    
    // הצגת התוצאות
    document.getElementById('analysisResults').style.display = 'block';
}

function generateStars(score) {
    const numStars = Math.round(score * 5);
    return '⭐'.repeat(numStars) + '☆'.repeat(5 - numStars);
}

function showLoading(show) {
    document.getElementById('loadingIndicator').style.display = show ? 'block' : 'none';
    document.getElementById('analyzeBtn').disabled = show;
}

function showError(message) {
    document.getElementById('errorText').textContent = message;
    document.getElementById('errorMessage').style.display = 'block';
}

function hideError() {
    document.getElementById('errorMessage').style.display = 'none';
}

function refreshEmailInfo() {
    updateCurrentEmailInfo();
}

function formatDate(date) {
    if (!date) return 'תאריך לא זמין';
    
    try {
        const dateObj = new Date(date);
        return dateObj.toLocaleString('he-IL', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch (error) {
        return 'תאריך לא תקין';
    }
}

function openProfileSettings() {
    try {
        // פתיחת חלון הגדרות במערכת הראשית
        window.open('https://localhost:5000/learning-management', '_blank');
    } catch (error) {
        console.error('Error opening profile settings:', error);
        showError('שגיאה בפתיחת הגדרות');
    }
}

function openStatistics() {
    try {
        // פתיחת חלון סטטיסטיקות במערכת הראשית
        window.open('https://localhost:5000/learning-management', '_blank');
    } catch (error) {
        console.error('Error opening statistics:', error);
        showError('שגיאה בפתיחת סטטיסטיקות');
    }
}

// פונקציות עזר נוספות
function getCurrentEmailId() {
    try {
        return Office.context.mailbox.item.itemId;
    } catch (error) {
        console.error('Error getting email ID:', error);
        return null;
    }
}

function isEmailSelected() {
    try {
        return Office.context.mailbox.item !== null;
    } catch (error) {
        console.error('Error checking email selection:', error);
        return false;
    }
}

// עדכון סטטוס החיבור
function updateConnectionStatus() {
    const statusIndicator = document.getElementById('statusIndicator');
    try {
        // בדיקה פשוטה אם Outlook זמין
        if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox) {
            statusIndicator.textContent = '🟢 מחובר ל-Outlook';
            statusIndicator.style.background = '#e8f5e8';
            statusIndicator.style.color = '#2e7d32';
        } else {
            statusIndicator.textContent = '🟡 מצב בדיקה בדפדפן';
            statusIndicator.style.background = '#fff3e0';
            statusIndicator.style.color = '#f57c00';
        }
    } catch (error) {
        statusIndicator.textContent = '🔴 שגיאה';
        statusIndicator.style.background = '#ffebee';
        statusIndicator.style.color = '#c62828';
    }
}

// עדכון סטטוס החיבור בהתחלה
updateConnectionStatus();
