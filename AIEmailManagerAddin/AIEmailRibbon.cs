using Microsoft.Office.Tools.Ribbon;
using System;
using System.Net.Http;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;

namespace AIEmailManagerAddin
{
    public partial class AIEmailRibbon
    {
        private const string API_BASE_URL = "http://localhost:5000";
        private static readonly HttpClient client = new HttpClient();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // אתחול ה-Ribbon
        }

        private void btnAnalyzeCurrent_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer.Selection.Count > 0)
                {
                    var mailItem = explorer.Selection[1] as Outlook.MailItem;
                    if (mailItem != null)
                    {
                        AnalyzeEmail(mailItem);
                    }
                    else
                    {
                        MessageBox.Show("אנא בחר מייל לניתוח", "AI Email Manager",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("אנא בחר מייל לניתוח", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void AnalyzeEmail(Outlook.MailItem mailItem)
        {
            try
            {
                // הצגת הודעה שהניתוח מתחיל
                MessageBox.Show("מתחיל ניתוח המייל...", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                // הכנת הנתונים
                var emailData = new
                {
                    subject = mailItem.Subject,
                    body = mailItem.Body,
                    sender = mailItem.SenderEmailAddress,
                    sender_name = mailItem.SenderName,
                    received_time = mailItem.ReceivedTime.ToString(),
                    date = mailItem.ReceivedTime.ToShortDateString(),
                    itemId = mailItem.EntryID,  // שליחת EntryID כדי שהשרת יוכל למצוא את המייל
                    entryID = mailItem.EntryID  // גם בשם המקורי
                };

                // שליחה ל-API
                var json = JsonConvert.SerializeObject(emailData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync($"{API_BASE_URL}/api/outlook-addin/analyze-email", content);

                if (response.IsSuccessStatusCode)
                {
                    var resultJson = await response.Content.ReadAsStringAsync();
                    
                    // DEBUG: הצג את ה-JSON
                    System.Diagnostics.Debug.WriteLine("API Response: " + resultJson);
                    
                    dynamic analysis = JsonConvert.DeserializeObject(resultJson);

                    // עדכון המייל
                    UpdateEmailWithAnalysis(mailItem, analysis);

                    MessageBox.Show("המייל נותח בהצלחה!", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"שגיאה בניתוח: {response.StatusCode}", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateEmailWithAnalysis(Outlook.MailItem mailItem, dynamic analysis)
        {
            try
            {
                // הוספת קטגוריה
                if (analysis.category != null)
                {
                    mailItem.Categories = analysis.category.ToString();
                }

                // הגדרת דחיפות
                if (analysis.priority != null)
                {
                    string priority = analysis.priority.ToString();
                    if (priority == "גבוהה")
                        mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                    else if (priority == "נמוכה")
                        mailItem.Importance = Outlook.OlImportance.olImportanceLow;
                }

                // הוספת דגל
                if (analysis.requires_action != null && (bool)analysis.requires_action)
                {
                    mailItem.FlagRequest = "למעקב";
                }

                // שמירת ציון AI ב-AISCORE ו-PRIORITYNUM
                try
                {
                    string aiScore = null;
                    int scorePercent = 0;
                    
                    // נסה למצוא את הציון בכמה שמות אפשריים
                    double scoreValue = 0;
                    
                    // ניסיון 1: importance_score
                    try 
                    { 
                        if (analysis.importance_score != null) 
                        {
                            scoreValue = Convert.ToDouble(analysis.importance_score);
                            // אם הציון הוא בין 0 ל-1, הכפל ב-100
                            if (scoreValue > 0 && scoreValue <= 1)
                                scoreValue *= 100;
                            scorePercent = (int)Math.Round(scoreValue);
                            aiScore = scorePercent.ToString();
                            System.Diagnostics.Debug.WriteLine($"DEBUG: מצא importance_score = {scoreValue} -> {scorePercent}%");
                        }
                    } 
                    catch (Exception ex) 
                    { 
                        System.Diagnostics.Debug.WriteLine($"DEBUG: שגיאה ב-importance_score: {ex.Message}");
                    }
                    
                    // ניסיון 2: ai_score
                    try 
                    { 
                        if (string.IsNullOrEmpty(aiScore) && analysis.ai_score != null) 
                        {
                            scoreValue = Convert.ToDouble(analysis.ai_score);
                            if (scoreValue > 0 && scoreValue <= 1)
                                scoreValue *= 100;
                            scorePercent = (int)Math.Round(scoreValue);
                            aiScore = scorePercent.ToString();
                            System.Diagnostics.Debug.WriteLine($"DEBUG: מצא ai_score = {scoreValue} -> {scorePercent}%");
                        }
                    } 
                    catch (Exception ex) 
                    { 
                        System.Diagnostics.Debug.WriteLine($"DEBUG: שגיאה ב-ai_score: {ex.Message}");
                    }
                    
                    // ניסיון 3: score
                    try 
                    { 
                        if (string.IsNullOrEmpty(aiScore) && analysis.score != null) 
                        {
                            scoreValue = Convert.ToDouble(analysis.score);
                            if (scoreValue > 0 && scoreValue <= 1)
                                scoreValue *= 100;
                            scorePercent = (int)Math.Round(scoreValue);
                            aiScore = scorePercent.ToString();
                            System.Diagnostics.Debug.WriteLine($"DEBUG: מצא score = {scoreValue} -> {scorePercent}%");
                        }
                    } 
                    catch (Exception ex) 
                    { 
                        System.Diagnostics.Debug.WriteLine($"DEBUG: שגיאה ב-score: {ex.Message}");
                    }
                    
                    // אם לא מצאנו, נסה לחלץ מה-JSON
                    if (string.IsNullOrEmpty(aiScore))
                    {
                        string jsonStr = JsonConvert.SerializeObject(analysis);
                        if (jsonStr.Contains("\"ai_score\":"))
                        {
                            var match = System.Text.RegularExpressions.Regex.Match(jsonStr, @"""ai_score"":\s*(\d+)");
                            if (match.Success)
                                aiScore = match.Groups[1].Value;
                        }
                        else if (jsonStr.Contains("\"score\":"))
                        {
                            var match = System.Text.RegularExpressions.Regex.Match(jsonStr, @"""score"":\s*(\d+)");
                            if (match.Success)
                                aiScore = match.Groups[1].Value;
                        }
                    }
                    
                    if (!string.IsNullOrEmpty(aiScore) && scorePercent > 0)
                    {
                        // עדכון PRIORITYNUM (מספר שלם)
                        var priorityNumProperty = mailItem.UserProperties.Find("PRIORITYNUM");
                        if (priorityNumProperty == null)
                        {
                            priorityNumProperty = mailItem.UserProperties.Add(
                                "PRIORITYNUM",
                                Outlook.OlUserPropertyType.olNumber);
                        }
                        priorityNumProperty.Value = scorePercent;
                        
                        System.Diagnostics.Debug.WriteLine($"DEBUG: PRIORITYNUM עודכן ל-{scorePercent}");
                        
                        // עדכון AISCORE (טקסט עם %)
                        var aiScoreProperty = mailItem.UserProperties.Find("AISCORE");
                        if (aiScoreProperty == null)
                        {
                            aiScoreProperty = mailItem.UserProperties.Add(
                                "AISCORE",
                                Outlook.OlUserPropertyType.olText);
                        }
                        
                        // אם הציון לא מסתיים ב-%, הוסף אותו
                        string aiScoreText = scorePercent + "%";
                        aiScoreProperty.Value = aiScoreText;
                        
                        System.Diagnostics.Debug.WriteLine($"DEBUG: AISCORE עודכן ל-{aiScoreText}");
                        
                        // שמור את המייל כדי שהשינויים ישמרו
                        mailItem.Save();
                        
                        System.Diagnostics.Debug.WriteLine($"DEBUG: המייל נשמר בהצלחה עם ציון {scorePercent}");
                        
                        // DEBUG: הצג הודעה
                        MessageBox.Show($"✅ עודכן בהצלחה!\n\nPRIORITYNUM: {scorePercent}\nAISCORE: {aiScoreText}", "עדכון הצליח");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("DEBUG: לא נמצא ציון AI בתגובה");
                        MessageBox.Show("⚠️ לא נמצא ציון AI בתגובה מהשרת", "שגיאה בניתוח");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"DEBUG: שגיאה בעדכון PRIORITYNUM/AISCORE: {ex.Message}");
                    MessageBox.Show($"⚠️ שגיאה בעדכון ציון:\n{ex.Message}", "שגיאה");
                }

                // שמירת ניתוח מפורט
                var analysisProperty = mailItem.UserProperties.Find("AI Analysis");
                if (analysisProperty == null)
                {
                    analysisProperty = mailItem.UserProperties.Add(
                        "AI Analysis",
                        Outlook.OlUserPropertyType.olText);
                }
                analysisProperty.Value = JsonConvert.SerializeObject(analysis);

                mailItem.Save();
                
                // נסה לרענן את התצוגה
                try
                {
                    var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                    if (explorer != null && explorer.CurrentFolder != null)
                    {
                        // שמור את התיקייה הנוכחית
                        var currentFolder = explorer.CurrentFolder;
                        // רענן את התצוגה על ידי מעבר לתיקייה אחרת וחזרה
                        var inbox = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                        if (currentFolder.EntryID != inbox.EntryID)
                        {
                            explorer.CurrentFolder = inbox;
                            explorer.CurrentFolder = currentFolder;
                        }
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                throw new Exception($"שגיאה בעדכון המייל: {ex.Message}");
            }
        }

        private async void btnAnalyzeFolder_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null || explorer.CurrentFolder == null)
                {
                    MessageBox.Show("אנא בחר תיקייה", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var folder = explorer.CurrentFolder;
                var items = folder.Items;
                
                // ספירת מיילים
                int totalEmails = 0;
                foreach (var item in items)
                {
                    if (item is Outlook.MailItem)
                        totalEmails++;
                }

                if (totalEmails == 0)
                {
                    MessageBox.Show("אין מיילים בתיקייה זו", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // אישור מהמשתמש
                var result = MessageBox.Show(
                    $"נמצאו {totalEmails} מיילים בתיקייה '{folder.Name}'.\n\nהאם לנתח את כולם? זה עלול לקחת זמן...",
                    "AI Email Manager",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes)
                    return;

                // ניתוח כל המיילים
                int analyzed = 0;
                int errors = 0;

                foreach (var item in items)
                {
                    if (item is Outlook.MailItem mailItem)
                    {
                        try
                        {
                            await AnalyzeEmailSilent(mailItem);
                            analyzed++;
                        }
                        catch
                        {
                            errors++;
                        }
                    }
                }

                MessageBox.Show(
                    $"ניתוח הושלם!\n\n" +
                    $"✓ נותחו: {analyzed} מיילים\n" +
                    $"✗ שגיאות: {errors}",
                    "AI Email Manager",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async System.Threading.Tasks.Task AnalyzeEmailSilent(Outlook.MailItem mailItem)
        {
            try
            {
                // הכנת הנתונים
                var emailData = new
                {
                    subject = mailItem.Subject,
                    body = mailItem.Body,
                    sender = mailItem.SenderEmailAddress,
                    sender_name = mailItem.SenderName,
                    received_time = mailItem.ReceivedTime.ToString(),
                    date = mailItem.ReceivedTime.ToShortDateString(),
                    itemId = mailItem.EntryID,
                    entryID = mailItem.EntryID
                };

                // שליחה ל-API
                var json = JsonConvert.SerializeObject(emailData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync($"{API_BASE_URL}/api/outlook-addin/analyze-email", content);

                if (response.IsSuccessStatusCode)
                {
                    var resultJson = await response.Content.ReadAsStringAsync();
                    dynamic analysis = JsonConvert.DeserializeObject(resultJson);

                    // עדכון המייל
                    UpdateEmailWithAnalysis(mailItem, analysis);
                }
            }
            catch
            {
                // המשך לניתוח הבא גם אם יש שגיאה
                throw;
            }
        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // פתיחת קונסול
                System.Diagnostics.Process.Start($"{API_BASE_URL}/consol");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה בפתיחת קונסול: {ex.Message}", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void ShowSystemInfo()
        {
            try
            {
                // בדיקת חיבור לשרת
                var response = await client.GetAsync($"{API_BASE_URL}/api/ai-status");
                
                string message = "📊 מידע על המערכת\n\n";
                
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    dynamic status = JsonConvert.DeserializeObject(json);
                    
                    message += $"🌐 שרת: מחובר ✓\n";
                    message += $"🤖 AI: {(status.ai_available == true ? "זמין ✓" : "לא זמין ✗")}\n";
                    message += $"🧠 למידה: {(status.use_ai == true ? "פעיל ✓" : "כבוי ✗")}\n\n";
                    message += $"📍 כתובת: {API_BASE_URL}\n";
                    message += $"📅 גרסה: 2.0\n";
                }
                else
                {
                    message += "⚠️ לא ניתן להתחבר לשרת\n\n";
                    message += "אנא ודא שהשרת פועל:\n";
                    message += "python app_with_ai.py";
                }

                var moreOptions = MessageBox.Show(
                    message + "\n\nלפתוח עוד אפשרויות?",
                    "מידע מערכת",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (moreOptions == DialogResult.Yes)
                {
                    ShowAdvancedOptions();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"⚠️ שרת לא זמין\n\n{ex.Message}\n\n" +
                    "אנא הפעל את השרת:\npython app_with_ai.py",
                    "שגיאת חיבור",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private async void ShowAdvancedOptions()
        {
            var result = MessageBox.Show(
                "🔧 אפשרויות מתקדמות\n\n" +
                "✅ Yes - צור גיבוי מלא\n" +
                "❌ No - רענן את כל הנתונים מ-Outlook\n" +
                "⚠️ Cancel - חזור",
                "אפשרויות מתקדמות",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            switch (result)
            {
                case DialogResult.Yes:
                    await CreateBackup();
                    break;

                case DialogResult.No:
                    await RefreshAllData();
                    break;
            }
        }

        private async System.Threading.Tasks.Task CreateBackup()
        {
            try
            {
                MessageBox.Show("יוצר גיבוי... אנא המתן", "גיבוי",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                var response = await client.PostAsync($"{API_BASE_URL}/api/create-backup", null);

                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    dynamic result = JsonConvert.DeserializeObject(json);

                    MessageBox.Show(
                        $"✓ גיבוי נוצר בהצלחה!\n\n" +
                        $"📁 מיקום: {result.backup_path}\n" +
                        $"📦 גודל: {result.file_size}",
                        "גיבוי הושלם",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("שגיאה ביצירת גיבוי", "שגיאה",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "שגיאה",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async System.Threading.Tasks.Task RefreshAllData()
        {
            try
            {
                MessageBox.Show("מרענן נתונים... אנא המתן", "רענון",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                var content = new StringContent("{\"type\":null}", Encoding.UTF8, "application/json");
                var response = await client.PostAsync($"{API_BASE_URL}/api/refresh-data", content);

                if (response.IsSuccessStatusCode)
                {
                    MessageBox.Show(
                        "✓ הנתונים עודכנו בהצלחה!\n\n" +
                        "כל המיילים והפגישות סונכרנו מ-Outlook.",
                        "רענון הושלם",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("שגיאה ברענון הנתונים", "שגיאה",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "שגיאה",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOpenWeb_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // פתיחת דף ניהול פגישות
                System.Diagnostics.Process.Start($"{API_BASE_URL}/meetings");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה בפתיחת ניהול פגישות: {ex.Message}", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ==================== פגישות ====================

        private void btnAnalyzeMeeting_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer.Selection.Count > 0)
                {
                    var appointmentItem = explorer.Selection[1] as Outlook.AppointmentItem;
                    if (appointmentItem != null)
                    {
                        AnalyzeMeeting(appointmentItem);
                    }
                    else
                    {
                        MessageBox.Show("אנא בחר פגישה לניתוח", "AI Email Manager",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("אנא בחר פגישה לניתוח", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void AnalyzeMeeting(Outlook.AppointmentItem appointmentItem)
        {
            try
            {
                MessageBox.Show("מתחיל ניתוח הפגישה...", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                // הכנת הנתונים
                var meetingData = new
                {
                    subject = appointmentItem.Subject,
                    body = appointmentItem.Body,
                    organizer = appointmentItem.Organizer,
                    start_time = appointmentItem.Start.ToString(),
                    end_time = appointmentItem.End.ToString(),
                    location = appointmentItem.Location,
                    required_attendees = appointmentItem.RequiredAttendees,
                    optional_attendees = appointmentItem.OptionalAttendees
                };

                // שליחה ל-API
                var json = JsonConvert.SerializeObject(meetingData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync($"{API_BASE_URL}/api/analyze-meetings-ai", content);

                if (response.IsSuccessStatusCode)
                {
                    var resultJson = await response.Content.ReadAsStringAsync();
                    dynamic analysis = JsonConvert.DeserializeObject(resultJson);

                    // הצגת תוצאות
                    string message = "📅 ניתוח הפגישה:\n\n";
                    message += $"נושא: {appointmentItem.Subject}\n";
                    message += $"מארגן: {appointmentItem.Organizer}\n\n";
                    
                    if (analysis.importance_score != null)
                    {
                        double score = Convert.ToDouble(analysis.importance_score);
                        if (score > 0 && score < 1) score *= 100;
                        message += $"📊 ציון חשיבות: {Math.Round(score)}%\n";
                    }
                    
                    if (analysis.category != null)
                        message += $"🏷️ קטגוריה: {analysis.category}\n";
                    
                    if (analysis.summary != null)
                        message += $"\n📝 סיכום:\n{analysis.summary}";

                    MessageBox.Show(message, "תוצאות ניתוח פגישה",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"שגיאה בניתוח: {response.StatusCode}", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnAnalyzeMeetings_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null || explorer.CurrentFolder == null)
                {
                    MessageBox.Show("אנא בחר תיקייה", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var folder = explorer.CurrentFolder;
                var items = folder.Items;
                
                // ספירת פגישות
                int totalMeetings = 0;
                foreach (var item in items)
                {
                    if (item is Outlook.AppointmentItem)
                        totalMeetings++;
                }

                if (totalMeetings == 0)
                {
                    MessageBox.Show("אין פגישות בתיקייה זו", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // אישור מהמשתמש
                var result = MessageBox.Show(
                    $"נמצאו {totalMeetings} פגישות.\n\nהאם לנתח את כולן?",
                    "AI Email Manager",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes)
                    return;

                // ניתוח כל הפגישות
                int analyzed = 0;
                int errors = 0;

                foreach (var item in items)
                {
                    if (item is Outlook.AppointmentItem appointmentItem)
                    {
                        try
                        {
                            await AnalyzeMeetingSilent(appointmentItem);
                            analyzed++;
                        }
                        catch
                        {
                            errors++;
                        }
                    }
                }

                MessageBox.Show(
                    $"ניתוח הושלם!\n\n" +
                    $"✓ נותחו: {analyzed} פגישות\n" +
                    $"✗ שגיאות: {errors}",
                    "AI Email Manager",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async System.Threading.Tasks.Task AnalyzeMeetingSilent(Outlook.AppointmentItem appointmentItem)
        {
            try
            {
                var meetingData = new
                {
                    subject = appointmentItem.Subject,
                    body = appointmentItem.Body,
                    organizer = appointmentItem.Organizer,
                    start_time = appointmentItem.Start.ToString(),
                    end_time = appointmentItem.End.ToString()
                };

                var json = JsonConvert.SerializeObject(meetingData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                await client.PostAsync($"{API_BASE_URL}/api/analyze-meetings-ai", content);
            }
            catch
            {
                throw;
            }
        }

        private async void btnRefreshMeetings_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                MessageBox.Show("מרענן פגישות... אנא המתן", "רענון פגישות",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                var content = new StringContent("{\"type\":\"meetings\"}", Encoding.UTF8, "application/json");
                var response = await client.PostAsync($"{API_BASE_URL}/api/refresh-data", content);

                if (response.IsSuccessStatusCode)
                {
                    MessageBox.Show(
                        "✓ הפגישות עודכנו בהצלחה!\n\n" +
                        "כל הפגישות סונכרנו מ-Outlook.",
                        "רענון הושלם",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("שגיאה ברענון הפגישות", "שגיאה",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "שגיאה",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ==================== ניהול מערכת ====================

        private void btnRefreshEmails_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                MessageBox.Show("מרענן מיילים... תכונה זו תתווסף בקרוב", "רענון מיילים",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "שגיאה",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLearningManagement_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start($"{API_BASE_URL}/learning-management");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה בפתיחת ניהול למידה: {ex.Message}", "שגיאה",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnStats_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // פתיחת דף ניהול מיילים
                System.Diagnostics.Process.Start($"{API_BASE_URL}/");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה בפתיחת ניהול מיילים: {ex.Message}", "שגיאה",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}