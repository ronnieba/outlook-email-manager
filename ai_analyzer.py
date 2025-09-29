"""
AI Email Analyzer using Gemini API
מערכת ניתוח מיילים חכמה עם AI
"""
import google.generativeai as genai
import json
from datetime import datetime
from config import GEMINI_API_KEY

class EmailAnalyzer:
    def __init__(self):
        self.model = None
        self.setup_gemini()
    
    def setup_gemini(self):
        """הגדרת Gemini API"""
        try:
            if GEMINI_API_KEY == 'your_api_key_here':
                print("⚠️ Gemini API Key not configured - AI will not be available")
                return False
            
            genai.configure(api_key=GEMINI_API_KEY)
            self.model = genai.GenerativeModel('gemini-1.5-flash')
            print("✅ Gemini API configured successfully!")
            return True
        except Exception as e:
            print(f"❌ Error configuring Gemini: {e}")
            return False
    
    def analyze_email_importance(self, email_data):
        """ניתוח חשיבות מייל עם AI"""
        if not self.model:
            return self.calculate_basic_importance(email_data)
        
        try:
            prompt = f"""
            נתח את החשיבות של המייל הבא (ציון 0-1):
            
            נושא: {email_data.get('subject', '')}
            שולח: {email_data.get('sender', '')}
            תוכן: {email_data.get('body_preview', '')[:300]}
            
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
            importance_score = float(response.text.strip())
            
            # הגבלת הציון לטווח 0-1
            importance_score = max(0.0, min(1.0, importance_score))
            
            print(f"🤖 AI ניתוח חשיבות: {importance_score}")
            return importance_score
            
        except Exception as e:
            print(f"❌ שגיאה בניתוח AI: {e}")
            return self.calculate_basic_importance(email_data)
    
    def calculate_basic_importance(self, email_data):
        """חישוב בסיסי של חשיבות (fallback)"""
        score = 0.5
        
        # בדיקת מילות מפתח חשובות
        important_keywords = ['חשוב', 'דחוף', 'urgent', 'important', 'meeting', 'פגישה', 'azure', 'microsoft', 'security', 'alert']
        subject = str(email_data.get('subject', '')).lower()
        body = str(email_data.get('body_preview', '')).lower()
        
        for keyword in important_keywords:
            if keyword in subject:
                score += 0.2
            if keyword in body:
                score += 0.1
        
        # בדיקת שולח חשוב
        important_senders = ['manager', 'boss', 'מנהל', 'hr', 'it', 'microsoft', 'azure', 'security', 'admin']
        sender = str(email_data.get('sender', '')).lower()
        
        for important_sender in important_senders:
            if important_sender in sender:
                score += 0.3
        
        return min(score, 1.0)
    
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
            summary = response.text.strip()
            
            print(f"🤖 AI סיכום: {summary[:30]}...")
            return summary
            
        except Exception as e:
            print(f"❌ שגיאה בסיכום AI: {e}")
            return self.basic_summary(email_data)
    
    def basic_summary(self, email_data):
        """סיכום בסיסי (fallback)"""
        subject = email_data.get('subject', 'ללא נושא')
        sender = email_data.get('sender', 'שולח לא ידוע')
        return f"מייל מ-{sender}: {subject}"
    
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
            category = response.text.strip().lower()
            
            # וידוא שהקטגוריה תקינה
            valid_categories = ['work', 'personal', 'marketing', 'system', 'urgent', 'meeting', 'notification']
            if category not in valid_categories:
                category = 'work'  # ברירת מחדל
            
            print(f"🤖 AI קטגוריה: {category}")
            return category
            
        except Exception as e:
            print(f"❌ שגיאה בקטלוג AI: {e}")
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
            פעולות נדרשות ממשיות מ: {email_data.get('subject', '')} - {email_data.get('body_preview', '')[:200]}
            תשובה: רשימה קצרה של פעולות אמיתיות או "אין" (רק אם יש פעולות כמו לענות, להתקשר, לשלוח מסמך)
            """
            
            response = self.model.generate_content(prompt, generation_config={
                'max_output_tokens': 100,
                'temperature': 0.1
            })
            action_items = [item.strip() for item in response.text.strip().split('\n') if item.strip() and item.strip() != 'אין' and len(item.strip()) > 3]
            
            print(f"🤖 AI פעולות: {len(action_items)} פעולות")
            return action_items
            
        except Exception as e:
            print(f"❌ שגיאה בחילוץ פעולות AI: {e}")
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
            תוכן: {email_data.get('body_preview', '')[:400]}
            
            פרופיל משתמש:
            {profile_context}
            
            החזר תשובה ב-JSON עם השדות הבאים:
            {{
                "importance_score": ציון חשיבות 0-1,
                "category": קטגוריה (work/personal/marketing/system/urgent/meeting/notification),
                "summary": סיכום קצר בעברית,
                "action_items": רשימת פעולות נדרשות ממשיות או [] (רק אם יש פעולות אמיתיות כמו "לענות", "להתקשר", "לשלוח מסמך")
            }}
            """
            
            response = self.model.generate_content(prompt, generation_config={
                'max_output_tokens': 300,
                'temperature': 0.2
            })
            
            # ניסיון לפרסר JSON
            try:
                analysis = json.loads(response.text.strip())
                
                # וידוא שהערכים תקינים
                importance_score = float(analysis.get('importance_score', 0.5))
                importance_score = max(0.0, min(1.0, importance_score))
                
                category = analysis.get('category', 'work')
                valid_categories = ['work', 'personal', 'marketing', 'system', 'urgent', 'meeting', 'notification']
                if category not in valid_categories:
                    category = 'work'
                
                summary = analysis.get('summary', '')
                action_items = analysis.get('action_items', [])
                
                print(f"🤖 AI ניתוח מתקדם: חשיבות {importance_score}, קטגוריה {category}")
                
                return {
                    'importance_score': importance_score,
                    'category': category,
                    'summary': summary,
                    'action_items': action_items
                }
                
            except json.JSONDecodeError:
                # אם JSON לא תקין, נשתמש בניתוח בסיסי
                print("⚠️ AI החזיר תשובה לא תקינה, משתמש בניתוח בסיסי")
                return self.basic_analysis_with_profile(email_data, user_preferences, user_categories)
            
        except Exception as e:
            print(f"❌ שגיאה בניתוח AI מתקדם: {e}")
            return self.basic_analysis_with_profile(email_data, user_preferences, user_categories)
    
    def basic_analysis_with_profile(self, email_data, user_preferences, user_categories):
        """ניתוח בסיסי עם התחשבות בפרופיל"""
        # חישוב חשיבות בסיסי
        importance_score = self.calculate_basic_importance(email_data)
        
        # התחשבות בהעדפות המשתמש
        if user_preferences:
            subject = str(email_data.get('subject', '')).lower()
            body = str(email_data.get('body_preview', '')).lower()
            
            for keyword, weight in user_preferences.items():
                if keyword.lower() in subject:
                    importance_score += weight * 0.2
                if keyword.lower() in body:
                    importance_score += weight * 0.1
        
        # התחשבות בקטגוריות חשובות
        if user_categories:
            category = self.basic_category(email_data)
            if category in user_categories:
                importance_score += user_categories[category] * 0.1
        
        importance_score = min(importance_score, 1.0)
        
        return {
            'importance_score': importance_score,
            'category': self.basic_category(email_data),
            'summary': self.basic_summary(email_data),
            'action_items': []
        }



