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
            פעולות נדרשות מ: {email_data.get('subject', '')}
            תשובה: רשימה קצרה או "אין"
            """
            
            response = self.model.generate_content(prompt, generation_config={
                'max_output_tokens': 100,
                'temperature': 0.1
            })
            action_items = [item.strip() for item in response.text.strip().split('\n') if item.strip() and item.strip() != 'אין']
            
            print(f"🤖 AI פעולות: {len(action_items)} פעולות")
            return action_items
            
        except Exception as e:
            print(f"❌ שגיאה בחילוץ פעולות AI: {e}")
            return []
    
    def is_ai_available(self):
        """בדיקה אם AI זמין"""
        return self.model is not None



