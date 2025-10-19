"""
AI Email Analyzer using Gemini API
מערכת ניתוח מיילים חכמה עם AI
"""
import google.generativeai as genai
import json
import os
import sys
from datetime import datetime
from config import GEMINI_API_KEY

# בלוע הודעות WARNING של Gemini
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
os.environ['GRPC_VERBOSITY'] = 'ERROR'
os.environ['GLOG_minloglevel'] = '3'
os.environ['GRPC_TRACE'] = ''
os.environ['ABSL_LOG_LEVEL'] = 'ERROR'

# השתקת הודעות שגיאה של gRPC
import warnings
warnings.filterwarnings("ignore")

# השתקת לוגים ברמה הגלובלית
import logging
logging.getLogger('google').setLevel(logging.ERROR)
logging.getLogger('grpc').setLevel(logging.ERROR)
logging.getLogger('absl').setLevel(logging.ERROR)
logging.getLogger('google.generativeai').setLevel(logging.ERROR)

class EmailAnalyzer:
    def __init__(self):
        self.model = None
        self.setup_gemini()
    
    def setup_gemini(self):
        """הגדרת Gemini API"""
        try:
            if GEMINI_API_KEY == 'your_api_key_here':
                # מפתח לא מוגדר – לא מדפיסים לקונסול/טרמינל
                return False
            
            # ההגדרות כבר מוגדרות ברמה הגלובלית
            
            genai.configure(api_key=GEMINI_API_KEY)
            # נסה מודלים שונים עד שנמצא אחד שעובד - התחלה עם המודלים החדשים
            models_to_try = ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.0-flash', 'gemini-2.0-pro-exp', 'gemini-flash-latest', 'gemini-pro-latest', 'gemini-1.5-pro', 'gemini-1.5-flash', 'gemini-pro', 'gemini-1.0-pro']
            self.model = None
            
            # בדיקה איזה מודלים זמינים
            try:
                available_models = genai.list_models()
                # print(f"Available models: {[m.name for m in available_models]}")
            except Exception as e:
                pass
                # print(f"Could not list models: {e}")
            
            # נסה עם גרסת API שונה
            try:
                import google.generativeai as genai_v1beta
                genai_v1beta.configure(api_key=GEMINI_API_KEY)
                available_models_v1beta = genai_v1beta.list_models()
                # print(f"Available models (v1beta): {[m.name for m in available_models_v1beta]}")
            except Exception as e:
                pass
                # print(f"Could not list models (v1beta): {e}")
            
            for model_name in models_to_try:
                try:
                    self.model = genai.GenerativeModel(model_name)
                    # בדיקה שהמודל עובד
                    test_response = self.model.generate_content("test")
                    # הצלחה – אין הדפסה לטרמינל
                    break
                except Exception as e:
                    # print(f"Model {model_name} failed: {e}")
                    continue
            
            if not self.model:
                # אין מודלים זמינים – שקט בטרמינל
                return False
        except Exception as e:
            # לא מדפיסים שגיאה לטרמינל
            return False
    
    def analyze_email_importance(self, email_data):
        """ניתוח חשיבות מייל עם AI"""
        # הפעלת AI אמיתי במקום fallback
        if not self.model:
            return self.calculate_basic_importance(email_data)
        
        try:
            prompt = f"""
            נתח את החשיבות של המייל הבא (ציון 0-1):
            
            נושא: {email_data.get('subject', '')}
            שולח: {email_data.get('sender', '')}
            תוכן: {email_data.get('body_preview', '')}
            
            קח בחשבון:
            - מילות מפתח חשובות (urgent, important, meeting, etc.)
            - שולח חשוב (manager, hr, it, etc.)
            - תוכן דחוף או קריטי
            - רלוונטיות לעבודה
            
            החזר רק ציון מספרי בין 0 ל-1 (לדוגמה: 0.8)
            """
            
            response = self.model.generate_content(prompt, generation_config={
                'max_output_tokens': 50,
                'temperature': 0.1
            })
            
            # קריאת התגובה מהמודל החדש
            try:
                # נסה דרך candidates
                if hasattr(response, 'candidates') and response.candidates:
                    candidate = response.candidates[0]
                    if hasattr(candidate, 'content') and hasattr(candidate.content, 'parts') and candidate.content.parts:
                        importance_score = float(candidate.content.parts[0].text.strip())
                    else:
                        importance_score = float(str(candidate).strip())
                elif hasattr(response, 'parts') and response.parts and len(response.parts) > 0:
                    importance_score = float(response.parts[0].text.strip())
                elif hasattr(response, 'text'):
                    importance_score = float(response.text.strip())
                else:
                    # נסה דרך אחרת
                    importance_score = float(str(response).strip())
            except Exception as parse_error:
                # אם יש שגיאה בפרסור, נשתמש בחישוב בסיסי
                return self.calculate_basic_importance(email_data)
            
            # הגבלת הציון לטווח 0-1
            importance_score = max(0.0, min(1.0, importance_score))
            
            # print(f"AI importance analysis: {importance_score}")
            return importance_score
            
        except Exception as e:
            # print(f"Error in AI analysis: {e}")
            return self.calculate_basic_importance(email_data)
    
    def calculate_basic_importance(self, email_data):
        """חישוב בסיסי של חשיבות (fallback)
        שמרני יותר כדי לא להגיע בקלות ל-100%.
        """
        score = 0.5

        # בדיקת מילות מפתח חשובות (משקלים מתונים)
        important_keywords = ['חשוב', 'דחוף', 'urgent', 'important', 'meeting', 'פגישה', 'azure', 'microsoft', 'security', 'alert']
        subject = str(email_data.get('subject', '')).lower()
        body = str(email_data.get('body_preview', '')).lower()

        for keyword in important_keywords:
            if keyword in subject:
                score += 0.09
            if keyword in body:
                score += 0.05

        # בדיקת שולח חשוב
        important_senders = ['manager', 'boss', 'מנהל', 'hr', 'it', 'microsoft', 'azure', 'security', 'admin']
        sender = str(email_data.get('sender', '')).lower()

        for important_sender in important_senders:
            if important_sender in sender:
                score += 0.12

        # cap ל-0.9 כדי להבחין מציון AI אמיתי שיכול להגיע ל-1.0
        return min(score, 0.85)
    
    def summarize_email(self, email_data):
        """סיכום מייל עם AI"""
        if not self.model:
            return self.basic_summary(email_data)
        
        try:
            prompt = f"""
            סכם בקצרה (משפט אחד): {email_data.get('subject', '')} מ-{email_data.get('sender', '')}
            """
            
            response = self.model.generate_content(prompt, generation_config={
                'max_output_tokens': 100,
                'temperature': 0.1
            })
            
            # קריאת התגובה מהמודל החדש
            summary = ""
            if hasattr(response, 'candidates') and response.candidates:
                candidate = response.candidates[0]
                if hasattr(candidate, 'content') and hasattr(candidate.content, 'parts') and candidate.content.parts:
                    summary = candidate.content.parts[0].text.strip()
                else:
                    summary = str(candidate).strip()
            elif hasattr(response, 'parts') and response.parts:
                summary = response.parts[0].text.strip()
            elif hasattr(response, 'text'):
                summary = response.text.strip()
            else:
                summary = str(response).strip()
            
            # print(f"AI summary: {summary[:30]}...")
            return summary
            
        except Exception as e:
            # print(f"Error in AI summary: {e}")
            return self.basic_summary(email_data)
    
    def generate_tasks_from_summary(self, summary):
        """יצירת רשימת משימות מהסיכום"""
        try:
            # ניתוח הסיכום לייצור משימות
            prompt = f"""
            על בסיס הסיכום הבא של מייל, צור רשימת משימות מעשיות:
            
            סיכום המייל:
            {summary}
            
            החזר רשימת משימות בפורמט JSON עם השדות הבאים:
            - title: כותרת המשימה (בעברית)
            - description: תיאור מפורט של המשימה (בעברית)
            - priority: רמת חשיבות (קריטי/חשוב/בינוני/נמוך)
            - category: קטגוריה (AI קריטי/AI חשוב/AI בינוני/AI נמוך)
            
            דוגמה:
            {{
                "tasks": [
                    {{
                        "title": "בדיקת הגדרות אבטחה",
                        "description": "בדוק את הגדרות האבטחה בחשבון Microsoft",
                        "priority": "חשוב",
                        "category": "AI חשוב"
                    }}
                ]
            }}
            
            חשוב: החזר רק משימות מעשיות שניתן לבצע. מקסימום 5 משימות.
            """
            
            response = self.gemini_model.generate_content(prompt)
            tasks_text = response.text
            
            # ניקוי התגובה
            tasks_text = self.clean_response_text(tasks_text)
            
            # חילוץ JSON
            import re
            json_match = re.search(r'\{.*\}', tasks_text, re.DOTALL)
            if json_match:
                json_str = json_match.group()
                import json
                data = json.loads(json_str)
                return data.get('tasks', [])
            
            # אם לא נמצא JSON, יצירת משימות בסיסיות
            return self.create_basic_tasks(summary)
            
        except Exception as e:
            print(f"❌ שגיאה בייצור משימות: {e}")
            return self.create_basic_tasks(summary)
    
    def create_basic_tasks(self, summary):
        """יצירת משימות בסיסיות אם AI נכשל"""
        tasks = []
        
        # זיהוי מילות מפתח ליצירת משימות
        summary_lower = summary.lower()
        print(f"🔍 מחפש מילות מפתח בסיכום: {summary_lower}")
        
        # זיהוי משימות טכניות
        if any(word in summary_lower for word in ["ג'וב", "job", "שרת", "server", "איפוס", "reset"]):
            print("✅ נמצאו מילות מפתח טכניות")
            tasks.append({
                "title": "יצירת ג'וב לאיפוס שרתים",
                "description": "צור ג'וב חדש לאיפוס השרתים כפי שנדרש",
                "priority": "חשוב",
                "category": "AI חשוב"
            })
        
        # זיהוי משימות בדיקה
        if any(word in summary_lower for word in ["בדיקה", "check", "בדוק", "היסטוריה", "history"]):
            tasks.append({
                "title": "בדיקת אפשרות למחיקת היסטוריה",
                "description": "בדוק איך ניתן למחוק את ההיסטוריה במערכת",
                "priority": "בינוני",
                "category": "AI בינוני"
            })
        
        # זיהוי משימות גיבוי
        if any(word in summary_lower for word in ["גיבוי", "backup", "גיבויים", "backups"]):
            tasks.append({
                "title": "בדיקת נושא גיבויים",
                "description": "בדוק את מצב הגיבויים של הג'ובים הקיימים",
                "priority": "חשוב",
                "category": "AI חשוב"
            })
        
        if any(word in summary_lower for word in ["אבטחה", "security", "הגנה"]):
            tasks.append({
                "title": "בדיקת הגדרות אבטחה",
                "description": "בדוק את הגדרות האבטחה בחשבון",
                "priority": "חשוב",
                "category": "AI חשוב"
            })
        
        if any(word in summary_lower for word in ["פגישה", "meeting", "ישיבה"]):
            tasks.append({
                "title": "הכנה לפגישה",
                "description": "הכן חומרים ומידע לפגישה הקרובה",
                "priority": "בינוני",
                "category": "AI בינוני"
            })
        
        if any(word in summary_lower for word in ["דוח", "report", "דיווח"]):
            tasks.append({
                "title": "הכנת דוח",
                "description": "הכן דוח על הנושא הנדון",
                "priority": "בינוני",
                "category": "AI בינוני"
            })
        
        if any(word in summary_lower for word in ["עדכון", "update", "שינוי"]):
            tasks.append({
                "title": "עדכון מידע",
                "description": "עדכן מידע רלוונטי",
                "priority": "נמוך",
                "category": "AI נמוך"
            })
        
        # אם לא נמצאו מילות מפתח, משימה כללית
        if not tasks:
            print("⚠️ לא נמצאו מילות מפתח ספציפיות, יוצר משימה כללית")
            tasks.append({
                "title": "פעולה נדרשת",
                "description": "בצע פעולה בהתאם לתוכן המייל",
                "priority": "בינוני",
                "category": "AI בינוני"
            })
        
        print(f"📋 נוצרו {len(tasks)} משימות בסיסיות")
        return tasks

    def expand_reply_text(self, brief_text, sender_email="", original_subject=""):
        """הרחבת טקסט תשובה קצר לתשובה פורמלית באנגלית ב-HTML"""
        
        # יצירת תשובה חכמה מבוססת על הטקסט שהמשתמש כתב
        expanded_text = self.create_smart_reply(brief_text, sender_email, original_subject)
        
        # יצירת HTML יפה
        return self.create_html_email(expanded_text, sender_email, original_subject)
    
    def create_smart_reply(self, brief_text, sender_email="", original_subject=""):
        """יצירת תשובה חכמה מבוססת על הטקסט הקצר"""
        # חילוץ שם מהכתובת
        sender_name = "Sir/Madam"
        if sender_email and "@" in sender_email:
            sender_name = sender_email.split("@")[0].replace(".", " ").replace("_", " ").title()
        
        # בדיקה אם הטקסט בעברית
        is_hebrew = any('\u0590' <= char <= '\u05FF' for char in brief_text)
        
        # ניתוח הטקסט הקצר ויצירת תשובה מתאימה
        brief_lower = brief_text.lower().strip()
        
        if is_hebrew:
            # תשובות בעברית
            if any(word in brief_lower for word in ["תודה", "תוד"]):
                if any(word in brief_lower for word in ["אישרתי", "אישור", "אוקיי", "בסדר"]):
                    return f"""שלום {sender_name},

תודה על המייל. אני מאשר שקיבלתי את הבקשה ואישרתי אותה.

אני מעריך את הפנייה ומצפה להמשך שיתוף הפעולה.

בברכה"""
                else:
                    return f"""שלום {sender_name},

תודה על המייל. אני מעריך את הפנייה.

אני אבדוק את ההודעה ואחזור אליך בהתאם.

בברכה"""
            
            elif any(word in brief_lower for word in ["אישרתי", "אישור", "אוקיי", "בסדר", "כן"]):
                return f"""שלום {sender_name},

תודה על המייל. אני מאשר שקיבלתי את הבקשה ואישרתי אותה.

הכל נראה טוב מצדי ואני אמשיך בהתאם.

בברכה"""
            
            elif any(word in brief_lower for word in ["לא", "לא רוצה", "דחה"]):
                return f"""שלום {sender_name},

תודה על המייל. לאחר שיקול דעת, אני נאלץ לדחות את הבקשה כרגע.

אני מעריך את ההבנה ומקווה שנוכל לעבוד יחד בעתיד.

בברכה"""
            
            elif any(word in brief_lower for word in ["אבדוק", "אני אבדוק", "אחזור"]):
                return f"""שלום {sender_name},

תודה על המייל. אני אבדוק את הבקשה ואחזור אליך בהקדם האפשרי.

אני מעריך את הסבלנות ואתן לך עדכון בקרוב.

בברכה"""
            
            elif any(word in brief_lower for word in ["פגישה", "מפגש", "ישיבה"]):
                return f"""שלום {sender_name},

תודה על המייל בנושא הפגישה. אני מעריך את הפנייה.

אני אבדוק את הפרטים ואאשר את הזמינות שלי.

בברכה"""
            
            else:
                # תשובה כללית בעברית מבוססת על הטקסט המקורי - עריכה חכמה
                return f"""שלום {sender_name},

תודה על המייל. {self.fix_hebrew_text(brief_text)}

אני מעריך את הפנייה ואחזור אליך בהתאם.

בברכה"""
        
        else:
            # תשובות באנגלית (הקוד הקיים)
            if any(word in brief_lower for word in ["תודה", "thanks", "thank you"]):
                if any(word in brief_lower for word in ["אישרתי", "confirmed", "approve", "ok", "okay"]):
                    return f"""Dear {sender_name},

Thank you for your email. I can confirm that I have reviewed and approved your request.

I appreciate you keeping me informed and look forward to our continued collaboration.

Best regards"""
                else:
                    return f"""Dear {sender_name},

Thank you for your email. I appreciate you taking the time to reach out to me.

I will review your message and respond accordingly.

Best regards"""
            
            elif any(word in brief_lower for word in ["אישרתי", "confirmed", "approve", "ok", "okay", "yes"]):
                return f"""Dear {sender_name},

Thank you for your email. I can confirm that I have approved your request.

Everything looks good on my end, and I will proceed accordingly.

Best regards"""
            
            elif any(word in brief_lower for word in ["לא", "no", "reject", "decline"]):
                return f"""Dear {sender_name},

Thank you for your email. After careful consideration, I must decline your request at this time.

I appreciate your understanding and hope we can work together in the future.

Best regards"""
            
            elif any(word in brief_lower for word in ["אני אבדוק", "i will check", "checking", "review"]):
                return f"""Dear {sender_name},

Thank you for your email. I will review your request and get back to you as soon as possible.

I appreciate your patience and will provide you with an update shortly.

Best regards"""
            
            elif any(word in brief_lower for word in ["פגישה", "meeting", "appointment"]):
                return f"""Dear {sender_name},

Thank you for your email regarding the meeting. I appreciate you reaching out.

I will review the details and confirm my availability.

Best regards"""
            
            else:
                # תשובה כללית מבוססת על הטקסט המקורי - עריכה חכמה
                return f"""Dear {sender_name},

Thank you for your email. {self.fix_english_text(brief_text)}

I appreciate you reaching out and will respond accordingly.

Best regards"""
    
    def create_html_email(self, content, sender_email="", subject=""):
        """יצירת מייל HTML פשוט ויפה"""
        # חילוץ שם מהכתובת
        sender_name = "Sir/Madam"
        if sender_email and "@" in sender_email:
            sender_name = sender_email.split("@")[0].replace(".", " ").replace("_", " ").title()
        
        # ניקוי התוכן
        content = content.replace("Dear Sender", f"Dear {sender_name}")
        content = content.replace("Dear [Name]", f"Dear {sender_name}")
        
        # בדיקה אם התוכן בעברית
        is_hebrew = any('\u0590' <= char <= '\u05FF' for char in content)
        direction = "rtl" if is_hebrew else "ltr"
        font_family = "'Segoe UI', 'David', 'Arial Hebrew', Tahoma, Geneva, Verdana, sans-serif" if is_hebrew else "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
        
        html_content = f"""
        <div style="font-family: {font_family}; max-width: 600px; margin: 0 auto; background-color: #ffffff; direction: {direction};">
            <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; border-radius: 8px 8px 0 0;">
                <h2 style="color: white; margin: 0; font-size: 18px; font-weight: 600;">Reply to: {subject}</h2>
            </div>
            
            <div style="padding: 30px; background-color: #ffffff; border: 1px solid #e1e5e9; border-top: none; border-radius: 0 0 8px 8px;">
                <div style="white-space: pre-line; line-height: 1.6; color: #333333; font-size: 14px;">
                    {content}
                </div>
            </div>
        </div>
        """
        
        return html_content
    
    def fix_hebrew_text(self, text):
        """תיקון שגיאות כתיב נפוצות בעברית"""
        # תיקון שגיאות כתיב נפוצות
        fixes = {
            'אחשר': 'אחזור',
            'אחשור': 'אחזור', 
            'מאור': 'מאוחר',
            'יותא': 'יותר',
            'יותר': 'יותר',
            'אוקיי': 'בסדר',
            'אוקי': 'בסדר',
            'תוד': 'תודה',
            'תודא': 'תודה',
            'אבדוק': 'אבדוק',
            'אבדק': 'אבדוק',
            'אבדקה': 'אבדוק',
            'אישרתי': 'אישרתי',
            'אישור': 'אישרתי',
            'אישרת': 'אישרתי',
            'פגישה': 'פגישה',
            'מפגש': 'פגישה',
            'ישיבה': 'פגישה'
        }
        
        # החלפת שגיאות כתיב
        for wrong, correct in fixes.items():
            text = text.replace(wrong, correct)
        
        return text
    
    def fix_english_text(self, text):
        """תיקון שגיאות כתיב נפוצות באנגלית"""
        # תיקון שגיאות כתיב נפוצות
        fixes = {
            'thnaks': 'thanks',
            'thnak': 'thank',
            'confrimed': 'confirmed',
            'confrim': 'confirm',
            'apporve': 'approve',
            'aproove': 'approve',
            'meeting': 'meeting',
            'meetin': 'meeting',
            'appointment': 'appointment',
            'appointmnet': 'appointment'
        }
        
        # החלפת שגיאות כתיב
        for wrong, correct in fixes.items():
            text = text.replace(wrong, correct)
        
        return text
    
    def clean_response_text(self, text):
        """ניקוי הטקסט מ-JSON/HTML ומטא-דאטה"""
        import re
        
        # הסרת JSON blocks
        text = re.sub(r'```json\s*.*?\s*```', '', text, flags=re.DOTALL)
        text = re.sub(r'```\s*.*?\s*```', '', text, flags=re.DOTALL)
        
        # הסרת JSON objects
        text = re.sub(r'\{[^}]*\}', '', text)
        
        # הסרת HTML tags
        text = re.sub(r'<[^>]+>', '', text)
        
        # הסרת מטא-דאטה נפוצה
        text = re.sub(r'Index:\s*\d+', '', text)
        text = re.sub(r'content\s*\}', '', text)
        text = re.sub(r'role"\s*:model"', '', text)
        text = re.sub(r'finish_reason:\s*\w+', '', text)
        text = re.sub(r'From:\s*.*?<', '', text)
        text = re.sub(r'Sent:\s*.*?PM', '', text)
        
        # הסרת "index: 0 content" וכל השורות שמכילות רק מספרים
        text = re.sub(r'index:\s*\d+\s*content', '', text, flags=re.IGNORECASE)
        text = re.sub(r'^\s*\d+\s*$', '', text, flags=re.MULTILINE)
        
        # הסרת שורות שמכילות רק תווים מיוחדים
        text = re.sub(r'^[^a-zA-Z\u0590-\u05FF]*$', '', text, flags=re.MULTILINE)
        
        # ניקוי שורות ריקות מרובות
        text = re.sub(r'\n\s*\n\s*\n+', '\n\n', text)
        
        # הסרת תווים מיוחדים
        text = text.replace('{', '').replace('}', '')
        
        # אם הטקסט ריק או מכיל רק תווים מיוחדים, החזר טקסט ברירת מחדל
        if not text.strip() or len(text.strip()) < 5:
            return "Thank you for your email. I appreciate your message and will respond accordingly.\n\nBest regards"
        
        return text.strip()
    
    def basic_summary(self, email_data):
        """סיכום בסיסי (fallback) - ניסיון ליצור סיכום אנושי מפורט של כמה משפטים"""
        subject = email_data.get('subject', 'ללא נושא')
        sender = email_data.get('sender', 'שולח לא ידוע')
        body = str(email_data.get('body_preview', '')).lower()
        
        # ניסיון לזהות את סוג המייל וליצור סיכום מפורט של כמה משפטים
        if 'upgrade' in subject.lower() or 'עדכן' in subject.lower():
            return f"הודעה מערכת מ-{sender} המבקשת עדכון או שדרוג של שירות. המייל מכיל הוראות מפורטות לביצוע העדכון ודרישות טכניות. יש צורך לבצע את העדכון כדי להמשיך להשתמש בשירותים. המייל כולל לינקים ומידע טכני נוסף."
        elif 'meeting' in subject.lower() or 'פגישה' in subject.lower():
            return f"הזמנה או תזכורת לפגישה מ-{sender}. המייל כולל פרטי זמן, מקום ותוכן הפגישה המתוכננת. יש צורך לאשר השתתפות או להכין חומרים רלוונטיים. המייל מכיל קישור לקביעת פגישה או פרטי קשר."
        elif 'urgent' in subject.lower() or 'דחוף' in subject.lower():
            return f"הודעה דחופה מ-{sender} שדורשת תשומת לב מיידית. המייל מכיל מידע קריטי או פעולה נדרשת בזמן קצר. יש צורך לטפל במייל זה בהקדם האפשרי. המייל כולל פרטי קשר או הוראות לפעולה מיידית."
        elif 'security' in subject.lower() or 'אבטחה' in subject.lower():
            return f"הודעה בנושא אבטחה מ-{sender}. המייל כולל התראות או הוראות הקשורות לאבטחת החשבון או המערכת. יש צורך לבדוק את מצב האבטחה ולבצע פעולות נדרשות. המייל מכיל מידע על ניסיונות כניסה או שינויים בחשבון."
        elif 'microsoft' in sender.lower() or 'azure' in sender.lower():
            return f"הודעה רשמית מ-Microsoft או Azure בנושא שירותים או עדכונים. המייל מכיל מידע על שינויים בשירותים, עדכוני תוכנה או הודעות מערכת חשובות. יש צורך לעדכן את השירותים או לבצע פעולות נדרשות. המייל כולל מידע טכני מפורט והוראות ביצוע."
        elif 'hotmail' in sender.lower() or 'outlook' in sender.lower():
            return f"הודעה מ-{sender} הקשורה לשירותי דואר אלקטרוני. המייל כולל מידע על שירותים, עדכונים או הוראות שימוש בפלטפורמה. יש צורך להכיר את השינויים החדשים או לבצע עדכונים נדרשים. המייל מכיל מידע על תכונות חדשות, שיפורים או שינויים בממשק."
        elif 'hr' in sender.lower() or 'משאבי אנוש' in sender.lower():
            return f"הודעה ממחלקת משאבי אנוש בנושא מדיניות או נהלים. המייל מכיל מידע על שינויים ארגוניים, נהלים חדשים או הודעות חשובות לעובדים. יש צורך להכיר את המדיניות החדשה או לבצע פעולות נדרשות. המייל כולל מידע על זכויות, חובות או תהליכים ארגוניים."
        elif len(body) > 200:
            # אם יש תוכן ארוך, ננסה לזהות את הנושא
            if 'שלום' in body[:50] or 'hello' in body[:50]:
                return f"הודעה מפורטת מ-{sender} עם תוכן עסקי או אישי. המייל כולל מידע נרחב ודורש קריאה מעמיקה להבנת כל הפרטים. יש צורך לנתח את התוכן ולבצע פעולות נדרשות. המייל מכיל מידע חשוב שדורש תשומת לב מיוחדת."
            else:
                return f"הודעה מ-{sender} עם תוכן מפורט. המייל מכיל מידע רב ופרטים חשובים שדורשים תשומת לב. יש צורך לקרוא את כל התוכן ולהבין את המשמעות המלאה. המייל כולל מידע טכני או עסקי מפורט."
        elif len(body) > 100:
            return f"הודעה בינונית מ-{sender} עם תוכן משמעותי. המייל מכיל מידע חשוב שדורש קריאה והבנה. יש צורך לטפל במייל זה בהתאם לתוכן. המייל כולל פרטים רלוונטיים לנושא הנדון."
        else:
            return f"הודעה קצרה מ-{sender} בנושא {subject}. המייל מכיל מידע בסיסי ונראה כהתראה או הודעה קצרה. יש צורך לקרוא את התוכן ולהבין את המשמעות. המייל כולל מידע חשוב שדורש תשומת לב."
    
    def categorize_email(self, email_data):
        """קטגוריזציה של מייל עם AI"""
        if not self.model:
            return self.basic_category(email_data)
        
        try:
            prompt = f"""
            קטלג: {email_data.get('subject', '')} מ-{email_data.get('sender', '')}
            תשובה: work/personal/marketing/system/urgent/meeting/notification
            """
            
            response = self.model.generate_content(prompt, generation_config={
                'max_output_tokens': 20,
                'temperature': 0.1
            })
            
            # קריאת התגובה מהמודל החדש
            category = ""
            if hasattr(response, 'candidates') and response.candidates:
                candidate = response.candidates[0]
                if hasattr(candidate, 'content') and hasattr(candidate.content, 'parts') and candidate.content.parts:
                    category = candidate.content.parts[0].text.strip().lower()
                else:
                    category = str(candidate).strip().lower()
            elif hasattr(response, 'parts') and response.parts:
                category = response.parts[0].text.strip().lower()
            elif hasattr(response, 'text'):
                category = response.text.strip().lower()
            else:
                category = str(response).strip().lower()
            
            # וידוא שהקטגוריה תקינה
            valid_categories = ['work', 'personal', 'marketing', 'system', 'urgent', 'meeting', 'notification']
            if category not in valid_categories:
                category = 'work'  # ברירת מחדל
            
            # print(f"AI category: {category}")
            return category
            
        except Exception as e:
            # print(f"Error in AI categorization: {e}")
            return self.basic_category(email_data)
    
    def basic_category(self, email_data):
        """קטגוריזציה בסיסית (fallback)"""
        subject = str(email_data.get('subject', '')).lower()
        sender = str(email_data.get('sender', '')).lower()
        
        if any(word in subject for word in ['meeting', 'פגישה', 'call']):
            return 'meeting'
        elif any(word in subject for word in ['urgent', 'דחוף', 'important', 'חשוב']):
            return 'urgent'
        elif any(word in sender for word in ['noreply', 'newsletter', 'marketing']):
            return 'marketing'
        elif any(word in sender for word in ['system', 'admin', 'it']):
            return 'system'
        else:
            return 'work'
    
    def extract_action_items(self, email_data):
        """חילוץ פעולות נדרשות עם AI"""
        if not self.model:
            return []
        
        try:
            prompt = f"""
            פעולות נדרשות ממשיות מ: {email_data.get('subject', '')} - {email_data.get('body_preview', '')}
            תשובה: רשימה קצרה של פעולות אמיתיות או "אין" (רק אם יש פעולות כמו לענות, להתקשר, לשלוח מסמך)
            """
            
            response = self.model.generate_content(prompt, generation_config={
                'max_output_tokens': 100,
                'temperature': 0.1
            })
            
            # קריאת התגובה מהמודל החדש
            response_text = ""
            if hasattr(response, 'candidates') and response.candidates:
                candidate = response.candidates[0]
                if hasattr(candidate, 'content') and hasattr(candidate.content, 'parts') and candidate.content.parts:
                    response_text = candidate.content.parts[0].text.strip()
                else:
                    response_text = str(candidate).strip()
            elif hasattr(response, 'parts') and response.parts:
                response_text = response.parts[0].text.strip()
            elif hasattr(response, 'text'):
                response_text = response.text.strip()
            else:
                response_text = str(response).strip()
            
            action_items = [item.strip() for item in response_text.split('\n') if item.strip() and item.strip() != 'אין' and len(item.strip()) > 3]
            
            # print(f"AI actions: {len(action_items)} actions")
            return action_items
            
        except Exception as e:
            # print(f"Error in AI action extraction: {e}")
            return []
    
    def is_ai_available(self):
        """בדיקה אם AI זמין"""
        return self.model is not None
    
    def analyze_email_with_profile(self, email_data, user_profile, user_preferences, user_categories):
        """ניתוח מייל עם AI כולל נתוני פרופיל משתמש"""
        if not self.model:
            return self.basic_analysis_with_profile(email_data, user_preferences, user_categories)
        
        try:
            # בניית פרומפט מתקדם עם נתוני פרופיל
            profile_context = ""
            if user_preferences:
                profile_context += f"מילות מפתח חשובות למשתמש: {', '.join(user_preferences.keys())}\n"
            
            if user_categories:
                important_cats = [cat for cat, score in user_categories.items() if score > 0.7]
                if important_cats:
                    profile_context += f"קטגוריות חשובות למשתמש: {', '.join(important_cats)}\n"
            
            prompt = f"""
            נתח את המייל הבא עם התחשבות בפרופיל המשתמש:
            
            נושא: {email_data.get('subject', '')}
            שולח: {email_data.get('sender', '')}
            תוכן: {email_data.get('body_preview', '')}
            
            פרופיל משתמש:
            {profile_context}
            
            החזר תשובה ב-JSON בלבד (ללא טקסט נוסף) עם השדות הבאים:
            {{
                "importance_score": ציון חשיבות 0-1,
                "category": קטגוריה (work/personal/marketing/system/urgent/meeting/notification),
                "summary": שני משפטים מלאים בעברית המסבירים את תכולת המייל ואת המשימה העיקרית בצורה אנושית וטבעית (לא לחזור על הכותרת, לא רשימות נקודות),
                "reason": משפט אחד מלא בעברית שמסביר בצורה אנושית וטבעית למה נקבעה רמת העדיפות (למשל: "המייל דורש אישור מיידי לפרויקט חשוב", "יש כאן דדליין קרוב שדורש תשומת לב", "השולח הוא מנהל בכיר שמבקש עדכון דחוף"),
                "action_items": רשימת פעולות נדרשות ממשיות או [] (רק אם יש פעולות אמיתיות כמו "לענות", "להתקשר", "לשלוח מסמך")
            }}
            """
            
            response = self.model.generate_content(prompt, generation_config={
                'max_output_tokens': 2000,
                'temperature': 0.2
            })
            
            # ניסיון לפרסר JSON
            try:
                # קריאת התגובה מהמודל החדש
                response_text = ""
                if hasattr(response, 'candidates') and response.candidates:
                    candidate = response.candidates[0]
                    if hasattr(candidate, 'content') and hasattr(candidate.content, 'parts') and candidate.content.parts:
                        response_text = candidate.content.parts[0].text.strip()
                    else:
                        response_text = str(candidate).strip()
                elif hasattr(response, 'parts') and response.parts:
                    response_text = response.parts[0].text.strip()
                elif hasattr(response, 'text'):
                    response_text = response.text.strip()
                else:
                    response_text = str(response).strip()
                
                analysis = json.loads(response_text)
                
                # וידוא שהערכים תקינים
                importance_score = float(analysis.get('importance_score', 0.5))
                importance_score = max(0.0, min(1.0, importance_score))
                
                category = analysis.get('category', 'work')
                valid_categories = ['work', 'personal', 'marketing', 'system', 'urgent', 'meeting', 'notification']
                if category not in valid_categories:
                    category = 'work'
                
                summary = analysis.get('summary', '')
                action_items = analysis.get('action_items', [])
                reason = analysis.get('reason', '')
                
                # print(f"AI advanced analysis: importance {importance_score}, category {category}")
                
                return {
                    'importance_score': importance_score,
                    'score_source': 'AI',
                    'category': category,
                    'summary': summary,
                    'action_items': action_items,
                    'reason': reason
                }
                
            except json.JSONDecodeError:
                # אם JSON לא תקין, נשתמש בניתוח בסיסי
                # print("AI returned invalid response, using basic analysis")
                return self.basic_analysis_with_profile(email_data, user_preferences, user_categories)
            
        except Exception as e:
            # print(f"Error in advanced AI analysis: {e}")
            return self.basic_analysis_with_profile(email_data, user_preferences, user_categories)
    
    def basic_analysis_with_profile(self, email_data, user_preferences, user_categories):
        """ניתוח בסיסי עם התחשבות בפרופיל ומלל הסבר אנושי"""
        base = 0.5
        importance_factors = []

        # מילות מפתח
        important_keywords = ['חשוב', 'דחוף', 'urgent', 'important', 'meeting', 'פגישה', 'azure', 'microsoft', 'security', 'alert']
        subject = str(email_data.get('subject', '')).lower()
        body = str(email_data.get('body_preview', '')).lower()
        matched_keywords = []
        for kw in important_keywords:
            if kw in subject or kw in body:
                matched_keywords.append(kw)
        if matched_keywords:
            base += 0.09 * len(set([kw for kw in matched_keywords if kw in subject]))
            base += 0.05 * len(set([kw for kw in matched_keywords if kw in body]))
            if 'דחוף' in matched_keywords or 'urgent' in matched_keywords:
                importance_factors.append("המייל מכיל מילות דחיפות")
            elif 'חשוב' in matched_keywords or 'important' in matched_keywords:
                importance_factors.append("המייל מסומן כחשוב")
            elif 'meeting' in matched_keywords or 'פגישה' in matched_keywords:
                importance_factors.append("המייל קשור לפגישה")

        # שולח חשוב
        important_senders = ['manager', 'boss', 'מנהל', 'hr', 'it', 'microsoft', 'azure', 'security', 'admin']
        sender = str(email_data.get('sender', '')).lower()
        matched_senders = [s for s in important_senders if s in sender]
        if matched_senders:
            base += 0.12
            if 'microsoft' in matched_senders or 'azure' in matched_senders:
                importance_factors.append("המייל מגיע מ-Microsoft או Azure")
            elif 'manager' in matched_senders or 'מנהל' in matched_senders:
                importance_factors.append("המייל מגיע ממנהל")
            else:
                importance_factors.append("המייל מגיע מגורם חשוב")

        # קטגוריות משמעותיות לפי בסיס
        category = self.basic_category(email_data)
        if user_categories and category in user_categories:
            base += user_categories[category] * 0.08
        if category == 'urgent':
            importance_factors.append("המייל מסווג כדחוף")
        elif category == 'meeting':
            importance_factors.append("המייל קשור לפגישה")

        # חיזוקים מעדפות מילים של המשתמש
        user_keywords_found = []
        if user_preferences:
            for kw, weight in user_preferences.items():
                if kw.lower() in subject or kw.lower() in body:
                    base += weight * 0.08
                    user_keywords_found.append(kw)
        
        if user_keywords_found:
            importance_factors.append(f"המייל מכיל מילות מפתח חשובות למשתמש: {', '.join(user_keywords_found)}")

        importance_score = min(base, 0.85)

        # יצירת הסבר אנושי שמסביר את שינוי הציון
        original_score = 0.5  # ציון בסיסי
        score_change = importance_score - original_score
        
        if importance_factors:
            if score_change > 0.15:
                if len(importance_factors) == 1:
                    reason = f"הציון עלה כי {importance_factors[0].lower()}"
                elif len(importance_factors) == 2:
                    reason = f"הציון עלה כי {importance_factors[0].lower()} וגם {importance_factors[1].lower()}"
                else:
                    reason = f"הציון עלה כי {', '.join([f.lower() for f in importance_factors[:-1]])} ו{importance_factors[-1].lower()}"
            elif score_change > 0.05:
                reason = f"הציון עלה מעט כי {importance_factors[0].lower()}"
            else:
                reason = f"הציון נשאר דומה כי {importance_factors[0].lower()}"
        else:
            if score_change < -0.1:
                reason = "הציון ירד כי המייל לא מכיל גורמי חשיבות משמעותיים"
            else:
                reason = "הציון נשאר בינוני כי המייל לא מכיל גורמי חשיבות מיוחדים"

        return {
            'importance_score': importance_score,
            'score_source': 'AI',  # גם ניתוח בסיסי נחשב AI אם נקרא מ-analyze_email_with_profile
            'category': category,
            'summary': self.basic_summary(email_data),
            'action_items': [],
            'reason': reason
        }



