using Microsoft.Office.Tools.Ribbon;
using System;
using System.Net.Http;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;

namespace AIEmailManagerAddin
{
    // מחלקות לתמיכה במשימות
    public class TaskItem
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string Priority { get; set; }
        public string Category { get; set; }
    }

    public class TaskGenerationResponse
    {
        public bool Success { get; set; }
        public List<TaskItem> Tasks { get; set; }
        public string Error { get; set; }
    }

    public class JiraSettings
    {
        public string JiraUrl { get; set; }
        public string Username { get; set; }
        public string ApiToken { get; set; }
        public string ProjectKey { get; set; }
        public string IssueType { get; set; }
    }

    public class CheckedIndexCollection
    {
        private int[] indices;
        
        public CheckedIndexCollection(int[] indices)
        {
            this.indices = indices;
        }
        
        public int Count => indices.Length;
        
        public int this[int index] => indices[index];
        
        public IEnumerator<int> GetEnumerator()
        {
            return ((IEnumerable<int>)indices).GetEnumerator();
        }
    }

    public partial class AIEmailRibbon
    {
        private const string API_BASE_URL = "http://localhost:5000";
        private static readonly HttpClient client = new HttpClient();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // אתחול ה-Ribbon
            // וידוא שכל הכפתורים נראים
            try
            {
                btnManageTasks.Visible = true;
                btnManageTasks.Enabled = true;
                btnExportToJira.Visible = true;
                btnExportToJira.Enabled = true;
                groupTasks.Visible = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"שגיאה באתחול Ribbon: {ex.Message}");
            }
        }

        // משתנה גלובלי לשמירת מידע על המייל הנוכחי
        private string currentMailItemId = null;
        private string currentMailSubject = null;
        private string currentMailSenderEmail = null;

        private void btnSummarizeEmail_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer != null && explorer.Selection.Count > 0)
                {
                    var selectedItem = explorer.Selection[1];
                    if (selectedItem is Outlook.MailItem mailItem)
                    {
                        // שמירת כל הנתונים מ-COM object על ה-UI thread
                        string itemId = mailItem.EntryID;
                        string subject = mailItem.Subject ?? "";
                        string body = mailItem.Body ?? "";
                        string senderEmail = mailItem.SenderEmailAddress ?? "";
                        string senderName = mailItem.SenderName ?? "";
                        
                        // שמירת מידע למשתנים גלובליים לשימוש מאוחר יותר
                        currentMailItemId = itemId;
                        currentMailSubject = subject;
                        currentMailSenderEmail = senderEmail;
                        
                        // שחרור ה-COM object
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem);
                        mailItem = null;
                        
                        // בדיקה אם יש סיכום קיים במאגר - synchronously
                        var checkData = new { item_id = itemId };
                        var checkJson = JsonConvert.SerializeObject(checkData);
                        var checkContent = new StringContent(checkJson, Encoding.UTF8, "application/json");
                        
                        var checkResponse = client.PostAsync($"{API_BASE_URL}/api/get-summary", checkContent).Result;
                        
                        if (checkResponse.IsSuccessStatusCode)
                        {
                            var checkResultJson = checkResponse.Content.ReadAsStringAsync().Result;
                            dynamic checkResult = JsonConvert.DeserializeObject(checkResultJson);
                            
                            // אם יש סיכום קיים - להציג אותו מיד
                            if (checkResult.has_summary == true)
                            {
                                System.Diagnostics.Debug.WriteLine("✅ נמצא סיכום קיים - מציג מיד");
                                ShowSummaryFormSync(checkResult, "סיכום המייל - AI (שמור)");
                                return;
                            }
                        }
                        
                        // אם אין סיכום קיים - מבצעים סיכום חדש
                        System.Diagnostics.Debug.WriteLine("🤖 אין סיכום קיים - מבצע סיכום חדש");
                        
                        // הצגת הודעת המתנה
                        var loadingForm = new Form
                        {
                            Text = "מעבד סיכום...",
                            Width = 400,
                            Height = 150,
                            StartPosition = FormStartPosition.CenterScreen,
                            FormBorderStyle = FormBorderStyle.FixedDialog,
                            MaximizeBox = false,
                            MinimizeBox = false,
                            RightToLeft = RightToLeft.Yes,
                            RightToLeftLayout = true,
                            TopMost = true
                        };

                        var loadingLabel = new Label
                        {
                            Text = "🤖 שולח מייל לסיכום AI...\n\nאנא המתן, התהליך עשוי לקחת מספר שניות.",
                            Dock = DockStyle.Fill,
                            TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                            Font = new System.Drawing.Font("Segoe UI", 11F),
                            Padding = new System.Windows.Forms.Padding(20)
                        };

                        loadingForm.Controls.Add(loadingLabel);
                        loadingForm.Show();
                        Application.DoEvents();
                        
                        try
                        {
                            // HTTP Request synchronously
                            var emailData = new
                            {
                                subject = subject,
                                body = body,
                                sender = senderEmail,
                                sender_name = senderName
                            };

                            var json = JsonConvert.SerializeObject(emailData);
                            var content = new StringContent(json, Encoding.UTF8, "application/json");

                            var response = client.PostAsync($"{API_BASE_URL}/api/summarize-email", content).Result;
                            
                            loadingForm.Close();

                            if (response.IsSuccessStatusCode)
                            {
                                var resultJson = response.Content.ReadAsStringAsync().Result;
                                dynamic analysis = JsonConvert.DeserializeObject(resultJson);

                                // שמירה למאגר נתונים - synchronously
                                var saveData = new
                                {
                                    item_id = itemId,
                                    summary = resultJson
                                };
                                var saveJson = JsonConvert.SerializeObject(saveData);
                                var saveContent = new StringContent(saveJson, Encoding.UTF8, "application/json");
                                var saveResponse = client.PostAsync($"{API_BASE_URL}/api/save-summary", saveContent).Result;
                                
                                if (saveResponse.IsSuccessStatusCode)
                                {
                                    System.Diagnostics.Debug.WriteLine("✅ הסיכום נשמר במאגר נתונים");
                                }

                                // הצגה ב-HTML
                                ShowSummaryFormSync(analysis, "סיכום המייל - AI");
                            }
                            else
                            {
                                MessageBox.Show($"שגיאה בסיכום המייל: {response.StatusCode}", "AI Email Manager",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        catch (Exception summaryEx)
                        {
                            loadingForm.Close();
                            MessageBox.Show($"שגיאה בעיבוד הסיכום: {summaryEx.Message}", "AI Email Manager",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("אנא בחר מייל לסיכום", "AI Email Manager",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("אנא בחר מייל לסיכום", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה: {ex.Message}", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowMeetingAnalysisForm(dynamic analysis, Outlook.AppointmentItem meeting, double scoreValue)
        {
            string summary = analysis.summary?.ToString() ?? "אין סיכום זמין";
            string category = analysis.category?.ToString() ?? "לא זוהה";
            string reason = analysis.reason?.ToString() ?? "";
            
            string htmlContent = $@"
<!DOCTYPE html>
<html dir='rtl' lang='he'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>ניתוח פגישה - AI</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            direction: rtl;
            padding: 20px;
        }}
        .container {{
            max-width: 800px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        .header h1 {{
            font-size: 28px;
            margin-bottom: 10px;
        }}
        .meeting-info {{
            background: rgba(255,255,255,0.1);
            padding: 15px;
            border-radius: 10px;
            margin-top: 15px;
        }}
        .meeting-info p {{
            margin: 5px 0;
            font-size: 14px;
        }}
        .content {{
            padding: 30px;
        }}
        .section {{
            margin-bottom: 25px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 12px;
            border-right: 4px solid #667eea;
        }}
        .section h2 {{
            color: #667eea;
            font-size: 20px;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        .section p {{
            color: #333;
            line-height: 1.8;
            font-size: 22px;
        }}
        .score-section {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            text-align: center;
        }}
        .score-section h2 {{
            color: white;
        }}
        .score-number {{
            font-size: 48px;
            font-weight: bold;
            margin: 10px 0;
        }}
        .category-badge {{
            display: inline-block;
            padding: 8px 20px;
            border-radius: 20px;
            background: rgba(255,255,255,0.2);
            font-weight: bold;
            font-size: 16px;
        }}
        .footer {{
            padding: 20px 30px;
            background: #f8f9fa;
            text-align: center;
            color: #666;
            font-size: 13px;
            border-top: 1px solid #dee2e6;
        }}
        .close-btn {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            padding: 12px 40px;
            border-radius: 25px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            margin-top: 15px;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
        }}
        .close-btn:hover {{
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4);
        }}
        .close-btn:active {{
            transform: translateY(0);
        }}
    </style>
</head>
<body>
    <div class='container'>
        <div class='header'>
            <h1>📅 ניתוח פגישה - AI</h1>
            <div class='meeting-info'>
                <p><strong>נושא:</strong> {meeting.Subject}</p>
                <p><strong>מארגן:</strong> {meeting.Organizer}</p>
                <p><strong>זמן:</strong> {meeting.Start:dd/MM/yyyy HH:mm} - {meeting.End:HH:mm}</p>
                {(string.IsNullOrEmpty(meeting.Location) ? "" : $"<p><strong>מיקום:</strong> {meeting.Location}</p>")}
            </div>
        </div>
        <div class='content'>
            <div class='section score-section'>
                <h2>📊 ציון חשיבות</h2>
                <div class='score-number'>{Math.Round(scoreValue)}%</div>
                <div class='category-badge'>{category}</div>
            </div>
            
            <div class='section'>
                <h2>📝 סיכום</h2>
                <p>{summary}</p>
            </div>
            
            {(string.IsNullOrEmpty(reason) ? "" : $@"
            <div class='section'>
                <h2>💡 הסבר לציון</h2>
                <p>{reason}</p>
            </div>
            ")}
        </div>
        <div class='footer'>
            <p>הניתוח נשמר אוטומטית במאגר הנתונים</p>
        </div>
    </div>
</body>
</html>";

            var form = new Form
            {
                Text = "ניתוח פגישה - AI",
                Width = 800,
                Height = 750,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.Sizable,
                MaximizeBox = true,
                MinimizeBox = true
            };

            var webBrowser = new WebBrowser
            {
                Dock = DockStyle.Fill,
                IsWebBrowserContextMenuEnabled = false,
                AllowNavigation = false,
                ScriptErrorsSuppressed = true
            };
            
            // כפתור סגירה (Windows Forms button)
            var closeButton = new Button
            {
                Text = "✓ סגור",
                Width = 120,
                Height = 45,
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                BackColor = ColorTranslator.FromHtml("#667eea"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Dock = DockStyle.Bottom
            };
            closeButton.FlatAppearance.BorderSize = 0;
            closeButton.Click += (s, e) => form.Close();
            
            webBrowser.DocumentText = htmlContent;
            form.Controls.Add(webBrowser);
            form.Controls.Add(closeButton);
            form.ShowDialog();
        }
        
        private void ShowSummaryFormSync(dynamic analysis, string title)
        {
            string summary = analysis.summary?.ToString() ?? "אין סיכום זמין";
            string sentiment = analysis.sentiment?.ToString() ?? "לא זוהה";
            
            string keyPointsHtml = "";
            if (analysis.key_points != null)
            {
                keyPointsHtml = "<ul style='margin: 10px 0; padding-right: 25px; line-height: 1.8;'>";
                if (analysis.key_points is string)
                {
                    keyPointsHtml += $"<li>{analysis.key_points}</li>";
                }
                else
                {
                    foreach (var point in analysis.key_points)
                    {
                        keyPointsHtml += $"<li>{point}</li>";
                    }
                }
                keyPointsHtml += "</ul>";
            }
            
            string actionItemsHtml = "";
            if (analysis.action_items != null)
            {
                actionItemsHtml = "<ul style='margin: 10px 0; padding-right: 25px; line-height: 1.8;'>";
                if (analysis.action_items is string)
                {
                    actionItemsHtml += $"<li>{analysis.action_items}</li>";
                }
                else
                {
                    foreach (var action in analysis.action_items)
                    {
                        actionItemsHtml += $"<li>{action}</li>";
                    }
                }
                actionItemsHtml += "</ul>";
            }
            
            string htmlContent = $@"
<!DOCTYPE html>
<html dir='rtl' lang='he'>
<head>
    <meta charset='UTF-8'>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            margin: 0;
            padding: 20px;
            direction: rtl;
        }}
        .container {{
            max-width: 900px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px 30px;
            text-align: center;
        }}
        .header h1 {{
            margin: 0;
            font-size: 28px;
            font-weight: 600;
        }}
        .content {{
            padding: 30px;
        }}
        .section {{
            margin-bottom: 25px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 10px;
            border-right: 4px solid #667eea;
        }}
        .section h2 {{
            color: #667eea;
            font-size: 20px;
            margin: 0 0 15px 0;
            font-weight: 600;
        }}
        .section p {{
            color: #333;
            line-height: 1.8;
            margin: 0;
            font-size: 22px;
        }}
        .section ul {{
            color: #333;
            margin: 10px 0;
            padding-right: 25px;
        }}
        .section li {{
            line-height: 1.8;
            margin-bottom: 8px;
            font-size: 22px;
        }}
        .sentiment {{
            display: inline-block;
            padding: 8px 20px;
            background: #28a745;
            color: white;
            border-radius: 20px;
            font-weight: 600;
            font-size: 14px;
        }}
        .footer {{
            padding: 20px 30px;
            background: #f8f9fa;
            text-align: center;
            color: #666;
            font-size: 13px;
            border-top: 1px solid #dee2e6;
        }}
    </style>
</head>
<body>
    <div class='container'>
        <div class='header'>
            <h1>סיכום המייל</h1>
        </div>
        <div class='content'>
            <div class='section'>
                <h2>סיכום</h2>
                <p>{summary}</p>
            </div>
            
            {(string.IsNullOrEmpty(keyPointsHtml) ? "" : $@"
            <div class='section'>
                <h2>נקודות מרכזיות</h2>
                {keyPointsHtml}
            </div>
            ")}
            
            {(string.IsNullOrEmpty(actionItemsHtml) ? "" : $@"
            <div class='section'>
                <h2>פעולות נדרשות</h2>
                {actionItemsHtml}
            </div>
            ")}
            
            <div class='section'>
                <h2>טון ההודעה</h2>
                <span class='sentiment'>{sentiment}</span>
            </div>
        </div>
        <div class='footer'>
            הסיכום נשמר אוטומטית במאגר הנתונים
        </div>
    </div>
</body>
</html>";

            var form = new Form
            {
                Text = title,
                Width = 1000,
                Height = 800,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.Sizable,
                MaximizeBox = true,
                MinimizeBox = true
            };

            var webBrowser = new WebBrowser
            {
                Dock = DockStyle.Fill,
                IsWebBrowserContextMenuEnabled = false,
                AllowNavigation = false,
                ScriptErrorsSuppressed = true
            };
            
            webBrowser.DocumentText = htmlContent;

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new System.Windows.Forms.Padding(10)
            };

            var btnReply = new Button
            {
                Text = "החזר תשובה",
                Width = 180,
                Height = 35,
                Margin = new System.Windows.Forms.Padding(5),
                BackColor = ColorTranslator.FromHtml("#667eea"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnReply.FlatAppearance.BorderSize = 0;
            btnReply.Click += (s, ev) => {
                // רק פתיחת חלון תשובה ללא סגירת החלון
                ShowReplyDialog();
            };

            // כפתור ייצור משימות
            var btnTasks = new Button
            {
                Text = "ייצר משימות",
                Width = 180,
                Height = 35,
                Margin = new System.Windows.Forms.Padding(5),
                BackColor = ColorTranslator.FromHtml("#28a745"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnTasks.FlatAppearance.BorderSize = 0;
            btnTasks.Click += (s, ev) => {
                // רק ייצור משימות ללא סגירת החלון
                GenerateTasksFromSummary(summary);
            };

            var btnClose = new Button
            {
                Text = "סגור",
                Width = 100,
                Height = 35,
                DialogResult = DialogResult.OK,
                Margin = new System.Windows.Forms.Padding(5),
                Anchor = AnchorStyles.None
            };

            buttonPanel.Controls.Add(btnReply);
            buttonPanel.Controls.Add(btnTasks);
            buttonPanel.Controls.Add(btnClose);

            form.Controls.Add(webBrowser);
            form.Controls.Add(buttonPanel);
            form.AcceptButton = btnClose;
            form.ShowDialog();
        }

        private async void GenerateTasksFromSummary(string summary)
        {
            try
            {
                // הצגת חלון המתנה
                var loadingForm = new Form
                {
                    Text = "מעבד משימות...",
                    Width = 400,
                    Height = 150,
                    StartPosition = FormStartPosition.CenterScreen,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    RightToLeft = RightToLeft.Yes,
                    RightToLeftLayout = true,
                    TopMost = true
                };

                var loadingLabel = new Label
                {
                    Text = "🤖 יוצר רשימת משימות מהסיכום...\nאנא המתן...",
                    AutoSize = false,
                    Width = 380,
                    Height = 100,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Segoe UI", 12F),
                    RightToLeft = RightToLeft.Yes
                };

                loadingForm.Controls.Add(loadingLabel);
                loadingForm.Show();

                // שליחה לשרת AI לייצור משימות
                var tasks = await GenerateTasksWithAI(summary);
                
                // סגירת חלון ההמתנה
                loadingForm.Close();
                loadingForm.Dispose();

                if (tasks != null && tasks.Count > 0)
                {
                    // הצגת חלון בחירת משימות
                    ShowTaskSelectionDialog(tasks);
                }
                else
                {
                    MessageBox.Show("לא ניתן לייצר משימות מהסיכום הזה.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ שגיאה בייצור משימות: {ex.Message}");
                MessageBox.Show($"שגיאה בייצור משימות: {ex.Message}", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async System.Threading.Tasks.Task<List<TaskItem>> GenerateTasksWithAI(string summary)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"🔧 מאתחל ייצור משימות...");
                System.Diagnostics.Debug.WriteLine($"📧 קיבלתי סיכום לייצור משימות: {summary.Substring(0, Math.Min(100, summary.Length))}...");
                
                Console.WriteLine($"🔧 מאתחל ייצור משימות...");
                Console.WriteLine($"📧 קיבלתי סיכום לייצור משימות: {summary.Substring(0, Math.Min(100, summary.Length))}...");
                
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(30);
                    
                    var requestData = new
                    {
                        summary = summary
                    };

                    var json = JsonConvert.SerializeObject(requestData);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    System.Diagnostics.Debug.WriteLine($"📡 שולח בקשה לשרת: http://localhost:5000/api/generate-tasks");
                    var response = await client.PostAsync("http://localhost:5000/api/generate-tasks", content);
                    
                    System.Diagnostics.Debug.WriteLine($"📡 תגובת שרת: {response.StatusCode}");
                    
                    Console.WriteLine($"📡 שולח בקשה לשרת: http://localhost:5000/api/generate-tasks");
                    Console.WriteLine($"📡 תגובת שרת: {response.StatusCode}");
                    
                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = await response.Content.ReadAsStringAsync();
                        System.Diagnostics.Debug.WriteLine($"📄 תוכן תגובה: {responseContent}");
                        
                        var result = JsonConvert.DeserializeObject<TaskGenerationResponse>(responseContent);
                        
                        if (result.Success && result.Tasks != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"✅ נוצרו {result.Tasks.Count} משימות");
                            return result.Tasks;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"❌ AI לא הצליח ליצור משימות: Success={result.Success}, Tasks={result.Tasks?.Count}");
                        }
                    }
                    else
                    {
                        var errorContent = await response.Content.ReadAsStringAsync();
                        System.Diagnostics.Debug.WriteLine($"❌ שגיאת שרת: {response.StatusCode} - {errorContent}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ שגיאה בשליחה לשרת AI: {ex.Message}");
            }
            
            return null;
        }

        private void ShowTaskSelectionDialog(List<TaskItem> tasks)
        {
            var selectionForm = new Form
            {
                Text = "בחר משימות לייצור",
                Width = 800,
                Height = 600,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.Sizable,
                MaximizeBox = true,
                MinimizeBox = true,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true
            };

            var checkedListBox = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                Font = new Font("Segoe UI", 12F),
                RightToLeft = RightToLeft.Yes
            };

            foreach (var task in tasks)
            {
                checkedListBox.Items.Add($"[{task.Priority}] {task.Title}", true);
            }

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new System.Windows.Forms.Padding(10)
            };

            var btnCreateTasks = new Button
            {
                Text = "צור משימות נבחרות",
                Width = 200,
                Height = 35,
                Margin = new System.Windows.Forms.Padding(5),
                BackColor = ColorTranslator.FromHtml("#28a745"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnCreateTasks.FlatAppearance.BorderSize = 0;
            btnCreateTasks.Click += (s, ev) => {
                CreateSelectedTasks(tasks, checkedListBox.CheckedIndices);
                // לא סוגר את החלון - רק יוצר משימות
            };

            // כפתור ייצוא ל-JIRA
            var btnExportToJira = new Button
            {
                Text = "ייצא ל-JIRA",
                Width = 150,
                Height = 35,
                Margin = new System.Windows.Forms.Padding(5),
                BackColor = ColorTranslator.FromHtml("#0052cc"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnExportToJira.FlatAppearance.BorderSize = 0;
            btnExportToJira.Click += (s, ev) => {
                var indices = new List<int>();
                foreach (int index in checkedListBox.CheckedIndices)
                {
                    indices.Add(index);
                }
                ExportSelectedTasksToJira(tasks, indices);
                // לא סוגר את החלון - רק מייצא ל-JIRA
            };

            var btnCancel = new Button
            {
                Text = "ביטול",
                Width = 100,
                Height = 35,
                Margin = new System.Windows.Forms.Padding(5),
                DialogResult = DialogResult.Cancel
            };

            var btnClose = new Button
            {
                Text = "סגור",
                Width = 100,
                Height = 35,
                Margin = new System.Windows.Forms.Padding(5),
                BackColor = ColorTranslator.FromHtml("#6c757d"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, ev) => {
                selectionForm.Close();
            };

            buttonPanel.Controls.Add(btnCreateTasks);
            buttonPanel.Controls.Add(btnExportToJira);
            buttonPanel.Controls.Add(btnCancel);
            buttonPanel.Controls.Add(btnClose);

            selectionForm.Controls.Add(checkedListBox);
            selectionForm.Controls.Add(buttonPanel);
            selectionForm.ShowDialog();
        }

        private void CreateSelectedTasks(List<TaskItem> tasks, CheckedListBox.CheckedIndexCollection selectedIndices)
        {
            try
            {
                var outlookApp = Globals.ThisAddIn.Application;
                var tasksFolder = outlookApp.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderTasks);

                foreach (int index in selectedIndices)
                {
                    var task = tasks[index];
                    var outlookTask = tasksFolder.Items.Add(Microsoft.Office.Interop.Outlook.OlItemType.olTaskItem) as Microsoft.Office.Interop.Outlook.TaskItem;
                    
                    if (outlookTask != null)
                    {
                        outlookTask.Subject = task.Title;
                        outlookTask.Body = task.Description;
                        outlookTask.Importance = GetOutlookImportance(task.Priority);
                        outlookTask.Categories = GetTaskCategory(task.Priority);
                        outlookTask.Save();
                        
                        System.Diagnostics.Debug.WriteLine($"✅ נוצרה משימה: {task.Title}");
                    }
                }

                MessageBox.Show($"נוצרו {selectedIndices.Count} משימות בהצלחה!", "הצלחה", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ שגיאה ביצירת משימות: {ex.Message}");
                MessageBox.Show($"שגיאה ביצירת משימות: {ex.Message}", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Microsoft.Office.Interop.Outlook.OlImportance GetOutlookImportance(string priority)
        {
            switch (priority.ToLower())
            {
                case "קריטי":
                case "critical":
                    return Microsoft.Office.Interop.Outlook.OlImportance.olImportanceHigh;
                case "חשוב":
                case "high":
                    return Microsoft.Office.Interop.Outlook.OlImportance.olImportanceHigh;
                case "בינוני":
                case "medium":
                    return Microsoft.Office.Interop.Outlook.OlImportance.olImportanceNormal;
                case "נמוך":
                case "low":
                    return Microsoft.Office.Interop.Outlook.OlImportance.olImportanceLow;
                default:
                    return Microsoft.Office.Interop.Outlook.OlImportance.olImportanceNormal;
            }
        }

        private string GetTaskCategory(string priority)
        {
            switch (priority.ToLower())
            {
                case "קריטי":
                case "critical":
                    return "AI קריטי";
                case "חשוב":
                case "high":
                    return "AI חשוב";
                case "בינוני":
                case "medium":
                    return "AI בינוני";
                case "נמוך":
                case "low":
                    return "AI נמוך";
                default:
                    return "AI בינוני";
            }
        }

        private async void ExportSelectedTasksToJira(List<TaskItem> tasks, List<int> selectedIndices)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"🔧 מאתחל ייצוא ל-JIRA...");
                System.Diagnostics.Debug.WriteLine($"📋 מספר משימות נבחרות: {selectedIndices.Count}");
                
                Console.WriteLine($"🔧 מאתחל ייצוא ל-JIRA...");
                Console.WriteLine($"📋 מספר משימות נבחרות: {selectedIndices.Count}");
                
                // הגדרות JIRA - קרא ממשתני סביבה או קובץ config
                var jiraSettings = new JiraSettings
                {
                    JiraUrl = Environment.GetEnvironmentVariable("JIRA_URL") ?? "YOUR_JIRA_URL",
                    Username = Environment.GetEnvironmentVariable("JIRA_USERNAME") ?? "YOUR_JIRA_USERNAME",
                    ApiToken = Environment.GetEnvironmentVariable("JIRA_API_TOKEN") ?? "YOUR_JIRA_API_TOKEN",
                    ProjectKey = Environment.GetEnvironmentVariable("JIRA_PROJECT_KEY") ?? "KAN",
                    IssueType = "Task" // יוחלף אוטומטית לפי התוכן
                };

                System.Diagnostics.Debug.WriteLine($"🔗 JIRA URL: {jiraSettings.JiraUrl}");
                System.Diagnostics.Debug.WriteLine($"👤 Username: {jiraSettings.Username}");
                System.Diagnostics.Debug.WriteLine($"🔑 Project Key: {jiraSettings.ProjectKey}");

                Console.WriteLine($"🔗 JIRA URL: {jiraSettings.JiraUrl}");
                Console.WriteLine($"👤 Username: {jiraSettings.Username}");
                Console.WriteLine($"🔑 Project Key: {jiraSettings.ProjectKey}");

                // ייצוא ישיר ללא חלון הגדרות
                await ExportTasksToJira(tasks, selectedIndices, jiraSettings);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ שגיאה בייצוא ל-JIRA: {ex.Message}");
                MessageBox.Show($"שגיאה בייצוא ל-JIRA: {ex.Message}", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async System.Threading.Tasks.Task ExportTasksToJira(List<TaskItem> tasks, List<int> selectedIndices, JiraSettings settings)
        {
            try
            {
                // הצגת חלון המתנה
                var loadingForm = new Form
                {
                    Text = "מייצא ל-JIRA...",
                    Width = 400,
                    Height = 150,
                    StartPosition = FormStartPosition.CenterScreen,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    RightToLeft = RightToLeft.Yes,
                    RightToLeftLayout = true,
                    TopMost = true
                };

                var loadingLabel = new Label
                {
                    Text = "🤖 מייצא משימות ל-JIRA...\nאנא המתן...",
                    AutoSize = false,
                    Width = 380,
                    Height = 100,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Segoe UI", 12F),
                    RightToLeft = RightToLeft.Yes
                };

                loadingForm.Controls.Add(loadingLabel);
                loadingForm.Show();

                int successCount = 0;
                int failCount = 0;

                foreach (int index in selectedIndices)
                {
                    var task = tasks[index];
                    try
                    {
                        var success = await CreateJiraIssue(task, settings);
                        if (success)
                            successCount++;
                        else
                            failCount++;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"❌ שגיאה ביצירת משימה {task.Title}: {ex.Message}");
                        failCount++;
                    }
                }

                // סגירת חלון ההמתנה
                loadingForm.Close();
                loadingForm.Dispose();

                MessageBox.Show($"ייצוא הושלם!\nנוצרו בהצלחה: {successCount}\nנכשלו: {failCount}", 
                    "תוצאות ייצוא", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ שגיאה כללית בייצוא ל-JIRA: {ex.Message}");
                MessageBox.Show($"שגיאה בייצוא ל-JIRA: {ex.Message}", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async System.Threading.Tasks.Task<bool> CreateJiraIssue(TaskItem task, JiraSettings settings)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"🔧 יוצר משימה ב-JIRA: {task.Title}");
                System.Diagnostics.Debug.WriteLine($"📝 תיאור: {task.Description}");
                
                Console.WriteLine($"🔧 יוצר משימה ב-JIRA: {task.Title}");
                Console.WriteLine($"📝 תיאור: {task.Description}");
                
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(30);
                    
                    // הגדרת אימות
                    var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{settings.Username}:{settings.ApiToken}"));
                    client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", credentials);
                    client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Add("Content-Type", "application/json");

                    // בחירת סוג משימה אוטומטית לפי התוכן
                    var issueType = GetIssueTypeByContent(task.Title, task.Description);
                    System.Diagnostics.Debug.WriteLine($"🏷️ סוג משימה שנבחר: {issueType}");

                    Console.WriteLine($"🏷️ סוג משימה שנבחר: {issueType}");

                    // יצירת ה-JSON ל-JIRA עם פורמט Atlassian Document
                    var jiraIssue = new
                    {
                        fields = new
                        {
                            project = new { key = settings.ProjectKey },
                            summary = task.Title,
                            description = new
                            {
                                type = "doc",
                                version = 1,
                                content = new[]
                                {
                                    new
                                    {
                                        type = "paragraph",
                                        content = new[]
                                        {
                                            new
                                            {
                                                type = "text",
                                                text = task.Description
                                            }
                                        }
                                    }
                                }
                            },
                            issuetype = new { name = issueType },
                            priority = new { name = GetJiraPriority(task.Priority) }
                            // labels = new[] { "AI-Generated", GetJiraCategory(task.Priority) } // הסרתי זמנית
                        }
                    };

                    var json = JsonConvert.SerializeObject(jiraIssue);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    System.Diagnostics.Debug.WriteLine($"🔗 שולח ל-JIRA: {settings.JiraUrl}/rest/api/3/issue");
                    System.Diagnostics.Debug.WriteLine($"📝 JSON: {json}");
                    
                    Console.WriteLine($"🔗 שולח ל-JIRA: {settings.JiraUrl}/rest/api/3/issue");
                    Console.WriteLine($"📝 JSON: {json}");
                    
                    var response = await client.PostAsync($"{settings.JiraUrl}/rest/api/3/issue", content);
                    
                    System.Diagnostics.Debug.WriteLine($"📡 תגובת JIRA: {response.StatusCode}");
                    
                    Console.WriteLine($"📡 תגובת JIRA: {response.StatusCode}");
                    
                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = await response.Content.ReadAsStringAsync();
                        var result = JsonConvert.DeserializeObject<dynamic>(responseContent);
                        var issueKey = result.key;
                        
                        System.Diagnostics.Debug.WriteLine($"✅ נוצרה משימה ב-JIRA: {issueKey} - {task.Title} ({issueType})");
                        return true;
                    }
                    else
                    {
                        var errorContent = await response.Content.ReadAsStringAsync();
                        System.Diagnostics.Debug.WriteLine($"❌ שגיאה ביצירת משימה ב-JIRA: {response.StatusCode} - {errorContent}");
                        MessageBox.Show($"שגיאת JIRA: {response.StatusCode}\n{errorContent}", "שגיאת JIRA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ שגיאה ביצירת משימה ב-JIRA: {ex.Message}");
                return false;
            }
        }

        private string GetIssueTypeByContent(string title, string description)
        {
            var content = (title + " " + description).ToLower();
            
            // זיהוי באגים
            if (content.Contains("bug") || content.Contains("error") || content.Contains("שגיאה") || 
                content.Contains("בעיה") || content.Contains("תקלה") || content.Contains("לא עובד"))
            {
                return "Bug";
            }
            
            // זיהוי סיפורי משתמש
            if (content.Contains("story") || content.Contains("feature") || content.Contains("תכונה") || 
                content.Contains("סיפור") || content.Contains("דרישה") || content.Contains("requirement"))
            {
                return "Story";
            }
            
            // זיהוי משימות טכניות
            if (content.Contains("task") || content.Contains("משימה") || content.Contains("עבודה") || 
                content.Contains("פיתוח") || content.Contains("development") || content.Contains("קוד"))
            {
                return "Task";
            }
            
            // זיהוי שיפורים
            if (content.Contains("improvement") || content.Contains("enhancement") || content.Contains("שיפור") || 
                content.Contains("שיפור") || content.Contains("אופטימיזציה"))
            {
                return "Improvement";
            }
            
            // ברירת מחדל
            return "Task";
        }

        private string GetJiraPriority(string priority)
        {
            switch (priority.ToLower())
            {
                case "קריטי":
                case "critical":
                    return "Highest";
                case "חשוב":
                case "high":
                    return "High";
                case "בינוני":
                case "medium":
                    return "Medium";
                case "נמוך":
                case "low":
                    return "Low";
                default:
                    return "Medium";
            }
        }

        private string GetJiraCategory(string priority)
        {
            switch (priority.ToLower())
            {
                case "קריטי":
                case "critical":
                    return "AI-Critical";
                case "חשוב":
                case "high":
                    return "AI-High";
                case "בינוני":
                case "medium":
                    return "AI-Medium";
                case "נמוך":
                case "low":
                    return "AI-Low";
                default:
                    return "AI-Medium";
            }
        }

        private void btnManageTasks_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // פתיחת חלון ניהול משימות
                ShowTaskManagementDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ שגיאה בניהול משימות: {ex.Message}");
                MessageBox.Show($"שגיאה בניהול משימות: {ex.Message}", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportToJira_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // פתיחת חלון ייצוא ל-JIRA
                ShowJiraExportDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ שגיאה בייצוא ל-JIRA: {ex.Message}");
                MessageBox.Show($"שגיאה בייצוא ל-JIRA: {ex.Message}", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowTaskManagementDialog()
        {
            try
            {
                var outlookApp = Globals.ThisAddIn.Application;
                var tasksFolder = outlookApp.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderTasks);
                
                var tasks = new List<TaskItem>();
                
                // איסוף כל המשימות עם קטגוריות AI
                foreach (Microsoft.Office.Interop.Outlook.TaskItem task in tasksFolder.Items)
                {
                    if (task.Categories != null && task.Categories.Contains("AI"))
                    {
                        tasks.Add(new TaskItem
                        {
                            Title = task.Subject ?? "",
                            Description = task.Body ?? "",
                            Priority = GetPriorityFromOutlook(task.Importance),
                            Category = task.Categories ?? ""
                        });
                    }
                }

                if (tasks.Count == 0)
                {
                    MessageBox.Show("לא נמצאו משימות שנוצרו על ידי AI.", "אין משימות", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // הצגת חלון ניהול משימות
                var managementForm = new Form
                {
                    Text = "ניהול משימות AI",
                    Width = 900,
                    Height = 600,
                    StartPosition = FormStartPosition.CenterScreen,
                    FormBorderStyle = FormBorderStyle.Sizable,
                    MaximizeBox = true,
                    MinimizeBox = true,
                    RightToLeft = RightToLeft.Yes,
                    RightToLeftLayout = true
                };

                var dataGridView = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    MultiSelect = true,
                    RightToLeft = RightToLeft.Yes,
                    Font = new Font("Segoe UI", 10F)
                };

                // הוספת עמודות
                dataGridView.Columns.Add("Title", "כותרת");
                dataGridView.Columns.Add("Priority", "חשיבות");
                dataGridView.Columns.Add("Category", "קטגוריה");
                dataGridView.Columns.Add("Description", "תיאור");

                // הוספת נתונים
                foreach (var task in tasks)
                {
                    dataGridView.Rows.Add(task.Title, task.Priority, task.Category, task.Description);
                }

                var buttonPanel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Height = 50,
                    FlowDirection = FlowDirection.RightToLeft,
                    Padding = new System.Windows.Forms.Padding(10)
                };

                var btnExportSelected = new Button
                {
                    Text = "ייצא נבחרות ל-JIRA",
                    Width = 200,
                    Height = 35,
                    Margin = new System.Windows.Forms.Padding(5),
                    BackColor = ColorTranslator.FromHtml("#0052cc"),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 10F, FontStyle.Bold)
                };
                btnExportSelected.FlatAppearance.BorderSize = 0;
                btnExportSelected.Click += (s, ev) => {
                    var selectedTasks = new List<TaskItem>();
                    foreach (DataGridViewRow row in dataGridView.SelectedRows)
                    {
                        selectedTasks.Add(tasks[row.Index]);
                    }
                    
                    if (selectedTasks.Count > 0)
                    {
                        // יצירת רשימת אינדקסים
                        var indices = new List<int>();
                        foreach (DataGridViewRow row in dataGridView.SelectedRows)
                        {
                            indices.Add(row.Index);
                        }
                        
                        // יצירת CheckedIndexCollection מותאם
                        var checkedIndices = new List<int>();
                        foreach (DataGridViewRow row in dataGridView.SelectedRows)
                        {
                            checkedIndices.Add(row.Index);
                        }
                        
                        ExportSelectedTasksToJira(selectedTasks, checkedIndices);
                    }
                    else
                    {
                        MessageBox.Show("אנא בחר משימות לייצוא.", "אין בחירה", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                };

                var btnClose = new Button
                {
                    Text = "סגור",
                    Width = 100,
                    Height = 35,
                    Margin = new System.Windows.Forms.Padding(5),
                    DialogResult = DialogResult.OK
                };

                buttonPanel.Controls.Add(btnExportSelected);
                buttonPanel.Controls.Add(btnClose);

                managementForm.Controls.Add(dataGridView);
                managementForm.Controls.Add(buttonPanel);
                managementForm.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ שגיאה בהצגת ניהול משימות: {ex.Message}");
                MessageBox.Show($"שגיאה בהצגת ניהול משימות: {ex.Message}", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowJiraExportDialog()
        {
            try
            {
                var outlookApp = Globals.ThisAddIn.Application;
                var tasksFolder = outlookApp.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderTasks);
                
                var tasks = new List<TaskItem>();
                
                // איסוף כל המשימות עם קטגוריות AI
                foreach (Microsoft.Office.Interop.Outlook.TaskItem task in tasksFolder.Items)
                {
                    if (task.Categories != null && task.Categories.Contains("AI"))
                    {
                        tasks.Add(new TaskItem
                        {
                            Title = task.Subject ?? "",
                            Description = task.Body ?? "",
                            Priority = GetPriorityFromOutlook(task.Importance),
                            Category = task.Categories ?? ""
                        });
                    }
                }

                if (tasks.Count == 0)
                {
                    MessageBox.Show("לא נמצאו משימות שנוצרו על ידי AI.", "אין משימות", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // הצגת חלון בחירת משימות לייצוא
                var selectionForm = new Form
                {
                    Text = "בחר משימות לייצוא ל-JIRA",
                    Width = 800,
                    Height = 600,
                    StartPosition = FormStartPosition.CenterScreen,
                    FormBorderStyle = FormBorderStyle.Sizable,
                    MaximizeBox = true,
                    MinimizeBox = true,
                    RightToLeft = RightToLeft.Yes,
                    RightToLeftLayout = true
                };

                var checkedListBox = new CheckedListBox
                {
                    Dock = DockStyle.Fill,
                    CheckOnClick = true,
                    Font = new Font("Segoe UI", 12F),
                    RightToLeft = RightToLeft.Yes
                };

                foreach (var task in tasks)
                {
                    checkedListBox.Items.Add($"[{task.Priority}] {task.Title}", true);
                }

                var buttonPanel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Height = 50,
                    FlowDirection = FlowDirection.RightToLeft,
                    Padding = new System.Windows.Forms.Padding(10)
                };

                var btnExport = new Button
                {
                    Text = "ייצא ל-JIRA",
                    Width = 150,
                    Height = 35,
                    Margin = new System.Windows.Forms.Padding(5),
                    BackColor = ColorTranslator.FromHtml("#0052cc"),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 10F, FontStyle.Bold)
                };
                btnExport.FlatAppearance.BorderSize = 0;
                btnExport.Click += (s, ev) => {
                    var indices = new List<int>();
                    foreach (int index in checkedListBox.CheckedIndices)
                    {
                        indices.Add(index);
                    }
                    ExportSelectedTasksToJira(tasks, indices);
                    selectionForm.Close();
                };

                var btnCancel = new Button
                {
                    Text = "ביטול",
                    Width = 100,
                    Height = 35,
                    Margin = new System.Windows.Forms.Padding(5),
                    DialogResult = DialogResult.Cancel
                };

                buttonPanel.Controls.Add(btnExport);
                buttonPanel.Controls.Add(btnCancel);

                selectionForm.Controls.Add(checkedListBox);
                selectionForm.Controls.Add(buttonPanel);
                selectionForm.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ שגיאה בהצגת ייצוא JIRA: {ex.Message}");
                MessageBox.Show($"שגיאה בהצגת ייצוא JIRA: {ex.Message}", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetPriorityFromOutlook(Microsoft.Office.Interop.Outlook.OlImportance importance)
        {
            switch (importance)
            {
                case Microsoft.Office.Interop.Outlook.OlImportance.olImportanceHigh:
                    return "חשוב";
                case Microsoft.Office.Interop.Outlook.OlImportance.olImportanceLow:
                    return "נמוך";
                default:
                    return "בינוני";
            }
        }

        private void ShowReplyDialog()
        {
            // יצירת חלון לקלט טקסט התשובה
            var inputForm = new Form
            {
                Text = "החזר תשובה - AI",
                Width = 600,
                Height = 400,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.Sizable,
                MaximizeBox = false,
                MinimizeBox = false,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true
            };

            var label = new Label
            {
                Text = "הקלד את התשובה הקצרה שלך (AI ירחיב אותה לתשובה פורמלית באנגלית):",
                Dock = DockStyle.Top,
                Height = 60,
                Font = new Font("Segoe UI", 11F),
                Padding = new Padding(15),
                TextAlign = ContentAlignment.MiddleRight
            };

            var textBox = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10F),
                Padding = new Padding(10),
                ScrollBars = ScrollBars.Vertical
            };

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(15)
            };

            var btnSend = new Button
            {
                Text = "שלח להרחבה ופתח תשובה",
                Width = 200,
                Height = 40,
                BackColor = ColorTranslator.FromHtml("#667eea"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnSend.FlatAppearance.BorderSize = 0;
            btnSend.Click += async (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(textBox.Text))
                {
                    MessageBox.Show("אנא הקלד טקסט לתשובה", "שגיאה", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                inputForm.Enabled = false;
                btnSend.Text = "מעבד...";
                
                string userText = textBox.Text;

                try
                {
                    // סגירה מלאה של חלון הקלט לפני התהליך
                    inputForm.Hide();
                    Application.DoEvents();
                    
                    await ExpandAndReply(userText);
                    
                    // סגירה סופית
                    if (!inputForm.IsDisposed)
                    {
                        inputForm.Close();
                        inputForm.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    if (!inputForm.IsDisposed)
                    {
                        inputForm.Show();
                        MessageBox.Show($"שגיאה: {ex.Message}", "שגיאה",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        inputForm.Enabled = true;
                        btnSend.Text = "שלח להרחבה ופתח תשובה";
                    }
                }
            };

            var btnCancel = new Button
            {
                Text = "ביטול",
                Width = 100,
                Height = 40,
                Margin = new Padding(5, 0, 0, 0)
            };
            btnCancel.Click += (s, e) => inputForm.Close();

            buttonPanel.Controls.Add(btnSend);
            buttonPanel.Controls.Add(btnCancel);

            inputForm.Controls.Add(textBox);
            inputForm.Controls.Add(label);
            inputForm.Controls.Add(buttonPanel);

            inputForm.ShowDialog();
        }

        private async System.Threading.Tasks.Task ExpandAndReply(string briefText)
        {
            // הצגת חלון המתנה
            var loadingForm = new Form
            {
                Text = "מעבד...",
                Width = 400,
                Height = 150,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true,
                TopMost = true
            };

            var loadingLabel = new Label
            {
                Text = "🤖 מרחיב את התשובה עם AI...\n\nאנא המתן, התהליך עשוי לקחת מספר שניות.",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 11F),
                Padding = new Padding(20)
            };

            loadingForm.Controls.Add(loadingLabel);
            loadingForm.Show();
            Application.DoEvents();

            try
            {
                // שליחת הטקסט לשרת להרחבה
                var requestData = new
                {
                    brief_text = briefText,
                    sender_email = currentMailSenderEmail,
                    original_subject = currentMailSubject
                };

                var json = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync($"{API_BASE_URL}/api/expand-reply", content);

                // סגירה מלאה של חלון ההמתנה
                if (loadingForm != null && !loadingForm.IsDisposed)
                {
                    loadingForm.Close();
                    loadingForm.Dispose();
                    loadingForm = null;
                }
                
                // וידוא שהחלון נסגר לגמרי לפני שממשיכים
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);

                if (response.IsSuccessStatusCode)
                {
                    var resultJson = await response.Content.ReadAsStringAsync();
                    dynamic result = JsonConvert.DeserializeObject(resultJson);

                    if (result.success == true && result.expanded_text != null)
                    {
                        string expandedText = result.expanded_text.ToString();
                        
                        // פתיחת חלון Reply ב-Outlook עם הטקסט המורחב
                        OpenReplyWithExpandedText(expandedText);
                    }
                    else
                    {
                        MessageBox.Show("שגיאה: לא התקבל טקסט מורחב מהשרת", "שגיאה",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show($"שגיאה בהרחבת הטקסט: {response.StatusCode}", "שגיאה",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                // סגירה בטוחה של חלון ההמתנה במקרה של שגיאה
                if (loadingForm != null && !loadingForm.IsDisposed)
                {
                    loadingForm.Close();
                    loadingForm.Dispose();
                }
                throw;
            }
        }

        private void OpenReplyWithExpandedText(string expandedText)
        {
            try
            {
                // מציאת המייל המקורי
                if (string.IsNullOrEmpty(currentMailItemId))
                {
                    MessageBox.Show("לא נמצא מייל מקורי", "שגיאה",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // קבלת הגישה ל-Outlook
                var outlookApp = Globals.ThisAddIn.Application;
                var ns = outlookApp.GetNamespace("MAPI");

                // מציאת המייל לפי EntryID
                Outlook.MailItem originalMail = null;
                try
                {
                    originalMail = ns.GetItemFromID(currentMailItemId) as Outlook.MailItem;
                }
                catch
                {
                    MessageBox.Show("לא ניתן למצוא את המייל המקורי", "שגיאה",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (originalMail != null)
                {
                    try
                    {
                        // יצירת Reply
                        var replyMail = originalMail.Reply() as Outlook.MailItem;
                        
                        if (replyMail != null)
                        {
                            try
                            {
                                // תיקון כתובת הנמען אם היא שגויה
                                if (!string.IsNullOrEmpty(currentMailSenderEmail) && 
                                    replyMail.To.Contains("reply-") && 
                                    replyMail.To.Contains("@email.microsoftemail.com"))
                                {
                                    // החלפת כתובת ה-reply בכתובת המקורית
                                    replyMail.To = currentMailSenderEmail;
                                    System.Diagnostics.Debug.WriteLine($"✅ תוקנה כתובת הנמען ל: {currentMailSenderEmail}");
                                }
                                
                                // הוספת הטקסט המורחב כ-HTML לגוף המייל
                                replyMail.HTMLBody = expandedText + "<br/><br/>" + replyMail.HTMLBody;
                                
                                // הצגת חלון ה-Reply (ללא modal - לא חוסם)
                                replyMail.Display(false);
                                
                                // אין להציג MessageBox כאן - זה יוצר קונפליקט עם חלון ה-Reply
                                System.Diagnostics.Debug.WriteLine("✅ חלון Reply נפתח בהצלחה");
                            }
                            finally
                            {
                                // שחרור COM object של Reply
                                if (replyMail != null)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(replyMail);
                                }
                            }
                        }
                    }
                    finally
                    {
                        // שחרור COM object של המייל המקורי
                        if (originalMail != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(originalMail);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה בפתיחת חלון Reply: {ex.Message}", "שגיאה",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

                    // הצגת ציון בהודעה
                    string message = "✅ המייל נותח בהצלחה!\n\n";
                    
                    // נסה לחלץ את הציון
                    double scoreValue = 0;
                    if (analysis.importance_score != null)
                    {
                        scoreValue = Convert.ToDouble(analysis.importance_score);
                        if (scoreValue > 0 && scoreValue <= 1) scoreValue *= 100;
                    }
                    else if (analysis.ai_score != null)
                    {
                        scoreValue = Convert.ToDouble(analysis.ai_score);
                        if (scoreValue > 0 && scoreValue <= 1) scoreValue *= 100;
                    }
                    
                    if (scoreValue > 0)
                    {
                        message += $"📊 ציון חשיבות: {Math.Round(scoreValue)}%\n";
                    }
                    
                    if (analysis.category != null)
                    {
                        message += $"🏷️ קטגוריה: {analysis.category}\n";
                    }
                    
                    MessageBox.Show(message, "תוצאות ניתוח",
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
                // נשמור את הציון לשימוש בקטגוריה
                int scorePercent = 0;
                try
                {
                    if (analysis.importance_score != null)
                    {
                        double scoreValue = Convert.ToDouble(analysis.importance_score);
                        if (scoreValue > 0 && scoreValue <= 1)
                            scoreValue *= 100;
                        scorePercent = (int)Math.Round(scoreValue);
                    }
                }
                catch { }
                
                // הוספת קטגוריה עם ציון
                try
                {
                    string categoryName = scorePercent > 0 ? $"AI: {scorePercent}%" : "AI";
                    
                    // שמור קטגוריות קיימות (אם יש) ומוסיף את החדשה
                    string existingCategories = mailItem.Categories;
                    if (!string.IsNullOrEmpty(existingCategories))
                    {
                        // מחק קטגוריות AI קודמות ושמור את השאר
                        var categories = existingCategories.Split(',')
                            .Select(c => c.Trim())
                            .Where(c => !c.StartsWith("AI:") && c != "AI")
                            .ToList();
                        categories.Add(categoryName);
                        mailItem.Categories = string.Join(", ", categories);
                    }
                    else
                    {
                        mailItem.Categories = categoryName;
                    }
                    System.Diagnostics.Debug.WriteLine($"DEBUG: קטגוריה עודכנה ל-{categoryName}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"DEBUG: שגיאה בעדכון קטגוריה: {ex.Message}");
                }
                
                // הוספת קטגוריה נוספת אם הוגדרה
                if (analysis.category != null)
                {
                    string additionalCategory = analysis.category.ToString();
                    if (!string.IsNullOrEmpty(additionalCategory))
                    {
                        string currentCategories = mailItem.Categories;
                        if (!currentCategories.Contains(additionalCategory))
                        {
                            mailItem.Categories = string.IsNullOrEmpty(currentCategories) 
                                ? additionalCategory 
                                : $"{currentCategories}, {additionalCategory}";
                        }
                    }
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
            Form loadingForm = null;
            try
            {
                // הודעת המתנה בזמן הניתוח
                loadingForm = new Form
                {
                    Text = "מנתח...",
                    Width = 300,
                    Height = 120,
                    StartPosition = FormStartPosition.CenterScreen,
                    FormBorderStyle = FormBorderStyle.None,
                    TopMost = true,
                    RightToLeft = RightToLeft.Yes
                };
                
                var loadingLabel = new Label
                {
                    Text = "🤖 מנתח את הפגישה עם AI...\nאנא המתן...",
                    AutoSize = false,
                    Width = 280,
                    Height = 80,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new System.Drawing.Font("Segoe UI", 12, System.Drawing.FontStyle.Bold),
                    Dock = DockStyle.Fill
                };
                
                loadingForm.Controls.Add(loadingLabel);
                loadingForm.Show();

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

                var response = await client.PostAsync($"{API_BASE_URL}/api/analyze-meeting", content);

                if (response.IsSuccessStatusCode)
                {
                    var resultJson = await response.Content.ReadAsStringAsync();
                    dynamic analysis = JsonConvert.DeserializeObject(resultJson);

                    // שמירת הציון ב-UserProperty
                    double scoreValue = 0;
                    if (analysis.importance_score != null)
                    {
                        scoreValue = Convert.ToDouble(analysis.importance_score);
                        if (scoreValue > 0 && scoreValue < 1) scoreValue *= 100;
                        
                        try
                        {
                            // עדכון PRIORITYNUM
                            var priorityNumProperty = appointmentItem.UserProperties.Find("PRIORITYNUM");
                            if (priorityNumProperty == null)
                            {
                                priorityNumProperty = appointmentItem.UserProperties.Add(
                                    "PRIORITYNUM",
                                    Outlook.OlUserPropertyType.olNumber);
                            }
                            priorityNumProperty.Value = (int)Math.Round(scoreValue);
                            
                            // עדכון AISCORE
                            var aiScoreProperty = appointmentItem.UserProperties.Find("AISCORE");
                            if (aiScoreProperty == null)
                            {
                                aiScoreProperty = appointmentItem.UserProperties.Add(
                                    "AISCORE",
                                    Outlook.OlUserPropertyType.olText);
                            }
                            aiScoreProperty.Value = $"{Math.Round(scoreValue)}%";
                            
                            // עדכון קטגוריה לפי הציון (כמו במיילים)
                            string categoryName = "";
                            if (scoreValue >= 80)
                                categoryName = "AI קריטי";
                            else if (scoreValue >= 60)
                                categoryName = "AI חשוב";
                            else if (scoreValue >= 40)
                                categoryName = "AI בינוני";
                            else
                                categoryName = "AI נמוך";
                            
                            appointmentItem.Categories = categoryName;
                            System.Diagnostics.Debug.WriteLine($"🏷️ קטגוריה עודכנה: {categoryName} (ציון: {scoreValue})");
                            
                            // עדכון דחיפות
                            if (scoreValue >= 80)
                                appointmentItem.Importance = Outlook.OlImportance.olImportanceHigh;
                            else if (scoreValue >= 50)
                                appointmentItem.Importance = Outlook.OlImportance.olImportanceNormal;
                            else
                                appointmentItem.Importance = Outlook.OlImportance.olImportanceLow;
                            
                            appointmentItem.Save();
                            System.Diagnostics.Debug.WriteLine($"✅ הפגישה נשמרה עם ציון {scoreValue}");
                        }
                        catch (Exception saveEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"⚠️ שגיאה בשמירת ציון: {saveEx.Message}");
                        }
                    }

                    // סגירת חלון ההמתנה
                    loadingForm.Close();
                    loadingForm.Dispose();
                    
                    // הצגת תוצאות בתצוגת HTML יפה
                    ShowMeetingAnalysisForm(analysis, appointmentItem, scoreValue);
                }
                else
                {
                    loadingForm.Close();
                    loadingForm.Dispose();
                    MessageBox.Show($"שגיאה בניתוח: {response.StatusCode}", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                if (loadingForm != null && !loadingForm.IsDisposed)
                {
                    loadingForm.Close();
                    loadingForm.Dispose();
                }
                MessageBox.Show($"שגיאה: {ex.Message}", "AI Email Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnAnalyzeMeetings_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null || explorer.Selection == null || explorer.Selection.Count == 0)
                {
                    MessageBox.Show("אנא בחר פגישות לניתוח", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // ספירת פגישות נבחרות
                int totalMeetings = 0;
                foreach (var item in explorer.Selection)
                {
                    if (item is Outlook.AppointmentItem)
                        totalMeetings++;
                }

                if (totalMeetings == 0)
                {
                    MessageBox.Show("לא נבחרו פגישות", "AI Email Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // אישור מהמשתמש
                var result = MessageBox.Show(
                    $"נמצאו {totalMeetings} פגישות נבחרות.\n\nהאם לנתח את כולן?",
                    "AI Email Manager",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes)
                    return;

                // ניתוח כל הפגישות הנבחרות
                int analyzed = 0;
                int errors = 0;

                foreach (var item in explorer.Selection)
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

                var response = await client.PostAsync($"{API_BASE_URL}/api/analyze-meeting", content);
                
                if (response.IsSuccessStatusCode)
                {
                    var resultJson = await response.Content.ReadAsStringAsync();
                    dynamic analysis = JsonConvert.DeserializeObject(resultJson);

                    // שמירת הציון ב-UserProperty
                    if (analysis.importance_score != null)
                    {
                        double scoreValue = Convert.ToDouble(analysis.importance_score);
                        if (scoreValue > 0 && scoreValue < 1) scoreValue *= 100;
                        
                        try
                        {
                            // עדכון PRIORITYNUM
                            var priorityNumProperty = appointmentItem.UserProperties.Find("PRIORITYNUM");
                            if (priorityNumProperty == null)
                            {
                                priorityNumProperty = appointmentItem.UserProperties.Add(
                                    "PRIORITYNUM",
                                    Outlook.OlUserPropertyType.olNumber);
                            }
                            priorityNumProperty.Value = (int)Math.Round(scoreValue);
                            
                            // עדכון AISCORE
                            var aiScoreProperty = appointmentItem.UserProperties.Find("AISCORE");
                            if (aiScoreProperty == null)
                            {
                                aiScoreProperty = appointmentItem.UserProperties.Add(
                                    "AISCORE",
                                    Outlook.OlUserPropertyType.olText);
                            }
                            aiScoreProperty.Value = $"{Math.Round(scoreValue)}%";
                            
                            // עדכון קטגוריה לפי הציון (כמו במיילים)
                            string categoryName = "";
                            if (scoreValue >= 80)
                                categoryName = "AI קריטי";
                            else if (scoreValue >= 60)
                                categoryName = "AI חשוב";
                            else if (scoreValue >= 40)
                                categoryName = "AI בינוני";
                            else
                                categoryName = "AI נמוך";
                            
                            appointmentItem.Categories = categoryName;
                            System.Diagnostics.Debug.WriteLine($"🏷️ קטגוריה עודכנה: {categoryName} (ציון: {scoreValue})");
                            
                            // עדכון דחיפות
                            if (scoreValue >= 80)
                                appointmentItem.Importance = Outlook.OlImportance.olImportanceHigh;
                            else if (scoreValue >= 50)
                                appointmentItem.Importance = Outlook.OlImportance.olImportanceNormal;
                            else
                                appointmentItem.Importance = Outlook.OlImportance.olImportanceLow;
                            
                            appointmentItem.Save();
                            System.Diagnostics.Debug.WriteLine($"✅ פגישה נשמרה: {appointmentItem.Subject} - ציון {scoreValue}");
                        }
                        catch (Exception saveEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"⚠️ שגיאה בשמירת פגישה {appointmentItem.Subject}: {saveEx.Message}");
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
        }

        private void btnRefreshMeetings_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // שאלת אישור
                var result = MessageBox.Show(
                    "האם לסנכרן את הפגישות מ-Outlook?\n\n" +
                    "⚠️ שים לב: פעולה זו רק מסנכרנת את רשימת הפגישות\n" +
                    "ללא ניתוח AI (חוסך כסף 💰)",
                    "סנכרון פגישות",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes)
                    return;

                // הודעת המתנה
                var loadingForm = new Form
                {
                    Text = "מסנכרן פגישות...",
                    Width = 400,
                    Height = 150,
                    StartPosition = FormStartPosition.CenterScreen,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    RightToLeft = RightToLeft.Yes,
                    RightToLeftLayout = true,
                    TopMost = true
                };

                var loadingLabel = new Label
                {
                    Text = "📅 מסנכרן פגישות מ-Outlook...\n\nאנא המתן...",
                    Dock = DockStyle.Fill,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                    Font = new System.Drawing.Font("Segoe UI", 11F),
                    Padding = new System.Windows.Forms.Padding(20)
                };

                loadingForm.Controls.Add(loadingLabel);
                loadingForm.Show();
                Application.DoEvents();

                try
                {
                    // שליחת בקשה לרענון פגישות בלבד (ללא AI)
                    var refreshData = new { type = "meetings" };
                    var json = JsonConvert.SerializeObject(refreshData);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = client.PostAsync($"{API_BASE_URL}/api/refresh-data", content).Result;
                    
                    loadingForm.Close();

                    if (response.IsSuccessStatusCode)
                    {
                        var resultJson = response.Content.ReadAsStringAsync().Result;
                        dynamic refreshResult = JsonConvert.DeserializeObject(resultJson);

                        string message = "✅ הפגישות עודכנו בהצלחה!\n\n";
                        
                        if (refreshResult.meetings_synced != null)
                            message += $"📅 פגישות שסונכרנו: {refreshResult.meetings_synced}\n";
                        
                        if (refreshResult.duration != null)
                            message += $"⏱️ משך: {refreshResult.duration}\n";
                        
                        message += "\n💡 לא בוצע ניתוח AI (חוסך כסף)";

                        MessageBox.Show(message, "סנכרון הושלם",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"שגיאה בסנכרון: {response.StatusCode}", "שגיאה",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception refreshEx)
                {
                    loadingForm.Close();
                    MessageBox.Show($"שגיאה בסנכרון הפגישות: {refreshEx.Message}", "שגיאה",
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
                // שאלת אישור
                var result = MessageBox.Show(
                    "האם לסנכרן את המיילים מ-Outlook?\n\n" +
                    "⚠️ שים לב: פעולה זו רק מסנכרנת את רשימת המיילים\n" +
                    "ללא ניתוח AI (חוסך כסף 💰)",
                    "סנכרון מיילים",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes)
                    return;

                // הודעת המתנה
                var loadingForm = new Form
                {
                    Text = "מסנכרן מיילים...",
                    Width = 400,
                    Height = 150,
                    StartPosition = FormStartPosition.CenterScreen,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    RightToLeft = RightToLeft.Yes,
                    RightToLeftLayout = true,
                    TopMost = true
                };

                var loadingLabel = new Label
                {
                    Text = "📧 מסנכרן מיילים מ-Outlook...\n\nאנא המתן...",
                    Dock = DockStyle.Fill,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                    Font = new System.Drawing.Font("Segoe UI", 11F),
                    Padding = new System.Windows.Forms.Padding(20)
                };

                loadingForm.Controls.Add(loadingLabel);
                loadingForm.Show();
                Application.DoEvents();

                try
                {
                    // שליחת בקשה לרענון מיילים בלבד (ללא AI)
                    var refreshData = new { type = "emails" };
                    var json = JsonConvert.SerializeObject(refreshData);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = client.PostAsync($"{API_BASE_URL}/api/refresh-data", content).Result;
                    
                    loadingForm.Close();

                    if (response.IsSuccessStatusCode)
                    {
                        var resultJson = response.Content.ReadAsStringAsync().Result;
                        dynamic refreshResult = JsonConvert.DeserializeObject(resultJson);

                        string message = "✅ המיילים עודכנו בהצלחה!\n\n";
                        
                        if (refreshResult.emails_synced != null)
                            message += $"📧 מיילים שסונכרנו: {refreshResult.emails_synced}\n";
                        
                        if (refreshResult.duration != null)
                            message += $"⏱️ משך: {refreshResult.duration}\n";
                        
                        message += "\n💡 לא בוצע ניתוח AI (חוסך כסף)";

                        MessageBox.Show(message, "סנכרון הושלם",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"שגיאה בסנכרון: {response.StatusCode}", "שגיאה",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception refreshEx)
                {
                    loadingForm.Close();
                    MessageBox.Show($"שגיאה בסנכרון המיילים: {refreshEx.Message}", "שגיאה",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
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