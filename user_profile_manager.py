"""
User Profile Manager - מערכת למידה חכמה לפרופיל משתמש
מערכת שמלמדת מההחלטות של המשתמש ומשפרת את ניתוח החשיבות
"""
import sqlite3
import json
from datetime import datetime, timedelta
from collections import defaultdict
import pickle
import os

class UserProfileManager:
    def __init__(self, db_path="email_preferences.db"):
        self.db_path = db_path
        self.profile_data = {}
        self.learning_weights = {}
        self.user_patterns = {}
        self.init_database()
        self.load_user_profile()
    
    def init_database(self):
        """יצירת טבלאות לפרופיל משתמש"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # טבלת משוב משתמש על מיילים
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_feedback (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email_id TEXT NOT NULL,
                subject TEXT,
                sender TEXT,
                user_importance_score REAL,
                ai_importance_score REAL,
                user_category TEXT,
                ai_category TEXT,
                feedback_type TEXT,  -- 'importance', 'category', 'action'
                feedback_value TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # טבלת דפוסי משתמש
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_patterns (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pattern_type TEXT NOT NULL,
                pattern_key TEXT NOT NULL,
                pattern_value TEXT NOT NULL,
                weight REAL DEFAULT 1.0,
                frequency INTEGER DEFAULT 1,
                last_used TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # טבלת העדפות משתמש מתקדמות
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_preferences_advanced (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                preference_type TEXT NOT NULL,
                preference_key TEXT NOT NULL,
                preference_value TEXT NOT NULL,
                confidence_score REAL DEFAULT 0.5,
                usage_count INTEGER DEFAULT 1,
                last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def load_user_profile(self):
        """טעינת פרופיל משתמש קיים"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # טעינת דפוסי משתמש
            cursor.execute('SELECT pattern_type, pattern_key, pattern_value, weight, frequency FROM user_patterns')
            patterns = cursor.fetchall()
            
            for pattern_type, pattern_key, pattern_value, weight, frequency in patterns:
                if pattern_type not in self.user_patterns:
                    self.user_patterns[pattern_type] = {}
                self.user_patterns[pattern_type][pattern_key] = {
                    'value': pattern_value,
                    'weight': weight,
                    'frequency': frequency
                }
            
            # טעינת העדפות מתקדמות
            cursor.execute('SELECT preference_type, preference_key, preference_value, confidence_score, usage_count FROM user_preferences_advanced')
            prefs = cursor.fetchall()
            
            for pref_type, pref_key, pref_value, confidence, usage_count in prefs:
                if pref_type not in self.profile_data:
                    self.profile_data[pref_type] = {}
                self.profile_data[pref_type][pref_key] = {
                    'value': pref_value,
                    'confidence': confidence,
                    'usage_count': usage_count
                }
            
            conn.close()
            # print(f"User profile loaded: {len(self.user_patterns)} patterns, {len(self.profile_data)} preferences")
            
        except Exception as e:
            pass  # Error loading profile - continue with defaults
    
    def record_user_feedback(self, email_data, feedback_type, user_value, ai_value=None):
        """רישום משוב משתמש מתקדם עם למידה מהתנהגות"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO user_feedback 
                (email_id, subject, sender, user_importance_score, ai_importance_score, 
                 user_category, ai_category, feedback_type, feedback_value)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                str(email_data.get('id', '')),
                email_data.get('subject', ''),
                email_data.get('sender', ''),
                user_value if feedback_type == 'importance' else None,
                ai_value if feedback_type == 'importance' else None,
                user_value if feedback_type == 'category' else None,
                ai_value if feedback_type == 'category' else None,
                feedback_type,
                str(user_value)
            ))
            
            conn.commit()
            conn.close()
            
            # עדכון דפוסי למידה
            self.update_learning_patterns(email_data, feedback_type, user_value)
            
            # למידה מהתנהגות משתמש
            self.learn_from_behavior(email_data, feedback_type, user_value, ai_value)
            
        except Exception as e:
            pass  # Error recording feedback - continue
    
    def learn_from_behavior(self, email_data, feedback_type, user_value, ai_value=None):
        """למידה מהתנהגות משתמש"""
        try:
            # למידה מההבדל בין המשתמש ל-AI
            if ai_value is not None and isinstance(user_value, (int, float)) and isinstance(ai_value, (int, float)):
                difference = abs(user_value - ai_value)
                
                # אם ההבדל גדול, זה אומר שהמשתמש חושב אחרת מהמערכת
                if difference > 0.3:
                    # למידה מהקשר מייל
                    self.learn_from_email_context(email_data, user_value, ai_value)
            
            # למידה מדפוסי זמן
            self.learn_from_temporal_patterns(email_data, feedback_type, user_value)
            
            # למידה מדפוסי שולחים
            self.learn_from_sender_patterns(email_data, feedback_type, user_value)
            
        except Exception as e:
            pass  # Error learning from behavior - continue
    
    def learn_from_email_context(self, email_data, user_value, ai_value):
        """למידה מהקשר המייל"""
        try:
            subject = email_data.get('subject', '').lower()
            sender = email_data.get('sender', '').lower()
            
            # אם המשתמש נתן ציון גבוה יותר מהמערכת
            if user_value > ai_value:
                # המערכת לא זיהתה משהו חשוב
                keywords = self.extract_keywords(subject)
                for keyword in keywords:
                    self.update_pattern('keyword_importance', keyword, user_value * 0.1)
                
                # למידה מהשולח
                self.update_pattern('sender_importance', sender, user_value * 0.1)
            
        except Exception as e:
            pass  # Error learning from context - continue
    
    def learn_from_temporal_patterns(self, email_data, feedback_type, user_value):
        """למידה מדפוסי זמן"""
        try:
            from datetime import datetime
            
            current_time = datetime.now()
            hour = current_time.hour
            day_of_week = current_time.weekday()
            
            # למידה מדפוסי זמן
            if feedback_type == 'importance' and isinstance(user_value, (int, float)):
                # אם המשתמש נותן ציונים גבוהים בשעות מסוימות
                if user_value > 0.7:
                    self.update_pattern('time_importance', f'hour_{hour}', user_value)
                    self.update_pattern('time_importance', f'day_{day_of_week}', user_value)
            
        except Exception as e:
            pass  # Error learning from temporal patterns - continue
    
    def learn_from_sender_patterns(self, email_data, feedback_type, user_value):
        """למידה מדפוסי שולחים"""
        try:
            sender = email_data.get('sender', '').lower()
            
            if feedback_type == 'importance' and isinstance(user_value, (int, float)):
                # למידה מדפוסי שולחים
                self.update_pattern('sender_importance', sender, user_value)
                
                # למידה מדפוסי דומיין
                if '@' in sender:
                    domain = sender.split('@')[1]
                    self.update_pattern('domain_importance', domain, user_value)
            
        except Exception as e:
            pass  # Error learning from sender patterns - continue
    
    def update_learning_patterns(self, email_data, feedback_type, user_value):
        """עדכון דפוסי למידה"""
        try:
            # למידה ממילות מפתח
            if feedback_type == 'importance':
                subject = email_data.get('subject', '').lower()
                sender = email_data.get('sender', '').lower()
                
                # חילוץ מילות מפתח מהנושא
                keywords = self.extract_keywords(subject)
                for keyword in keywords:
                    self.update_pattern('keyword_importance', keyword, user_value)
                
                # למידה מהשולח
                self.update_pattern('sender_importance', sender, user_value)
            
            # למידה מקטגוריות
            elif feedback_type == 'category':
                subject = email_data.get('subject', '').lower()
                keywords = self.extract_keywords(subject)
                for keyword in keywords:
                    self.update_pattern('keyword_category', keyword, user_value)
            
        except Exception as e:
            pass  # Error updating patterns - continue
    
    def extract_keywords(self, text):
        """חילוץ מילות מפתח דינמי מטקסט"""
        import re
        
        # מילות מפתח בסיסיות
        base_keywords = [
            'חשוב', 'דחוף', 'urgent', 'important', 'meeting', 'פגישה',
            'azure', 'microsoft', 'security', 'alert', 'error', 'שגיאה',
            'manager', 'מנהל', 'hr', 'it', 'admin', 'boss', 'בוס',
            'deadline', 'תאריך', 'project', 'פרויקט', 'report', 'דוח'
        ]
        
        keywords = []
        text_lower = text.lower()
        
        # חילוץ מילות מפתח בסיסיות
        for word in base_keywords:
            if word in text_lower:
                keywords.append(word)
        
        # חילוץ מילות מפתח דינמיות מהדפוסים הקיימים
        if 'keyword_importance' in self.user_patterns:
            for pattern_key in self.user_patterns['keyword_importance'].keys():
                if pattern_key in text_lower and pattern_key not in keywords:
                    keywords.append(pattern_key)
        
        # חילוץ מילות מפתח חדשות (מילים עם משמעות)
        # חיפוש מילים בעברית (אותיות עבריות)
        hebrew_words = re.findall(r'[\u0590-\u05FF]+', text_lower)
        for word in hebrew_words:
            if len(word) >= 3 and word not in keywords:
                keywords.append(word)
        
        # חיפוש מילים באנגלית (אותיות לטיניות)
        english_words = re.findall(r'[a-zA-Z]+', text_lower)
        for word in english_words:
            if len(word) >= 4 and word not in keywords:
                keywords.append(word)
        
        # חילוץ מספרים וקודים מיוחדים
        special_codes = re.findall(r'\b[A-Z]{2,}\b|\b\d{3,}\b', text)
        for code in special_codes:
            if code.lower() not in keywords:
                keywords.append(code.lower())
        
        return keywords[:20]  # הגבלה ל-20 מילות מפתח
    
    def update_pattern(self, pattern_type, pattern_key, value):
        """עדכון דפוס למידה מתקדם עם התחשבות בזמן"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # בדיקה אם הדפוס קיים
            cursor.execute('''
                SELECT weight, frequency, last_used FROM user_patterns 
                WHERE pattern_type = ? AND pattern_key = ?
            ''', (pattern_type, pattern_key))
            
            result = cursor.fetchone()
            current_time = datetime.now()
            
            if result:
                # עדכון דפוס קיים
                old_weight, old_freq, last_used = result
                new_freq = old_freq + 1
                
                # חישוב משקל חדש עם התחשבות בזמן
                if isinstance(value, (int, float)):
                    # חישוב זמן שעבר מהשימוש האחרון
                    if isinstance(last_used, str):
                        last_used = datetime.fromisoformat(last_used.replace('Z', '+00:00'))
                    
                    days_since_last_use = (current_time - last_used).days
                    
                    # משקל זמן - דפוסים ישנים מאבדים משקל
                    time_decay = max(0.5, 1.0 - (days_since_last_use * 0.01))
                    
                    # ממוצע משוקלל עם התחשבות בזמן
                    new_weight = (old_weight * old_freq * time_decay + value) / (old_freq * time_decay + 1)
                else:
                    new_weight = old_weight  # שמירה על המשקל הקיים
                
                cursor.execute('''
                    UPDATE user_patterns 
                    SET weight = ?, frequency = ?, last_used = CURRENT_TIMESTAMP
                    WHERE pattern_type = ? AND pattern_key = ?
                ''', (new_weight, new_freq, pattern_type, pattern_key))
            else:
                # יצירת דפוס חדש
                weight = value if isinstance(value, (int, float)) else 1.0
                cursor.execute('''
                    INSERT INTO user_patterns (pattern_type, pattern_key, pattern_value, weight, frequency)
                    VALUES (?, ?, ?, ?, 1)
                ''', (pattern_type, pattern_key, str(value), weight))
            
            conn.commit()
            conn.close()
            
            # עדכון זיכרון
            if pattern_type not in self.user_patterns:
                self.user_patterns[pattern_type] = {}
            self.user_patterns[pattern_type][pattern_key] = {
                'value': str(value),
                'weight': weight if isinstance(value, (int, float)) else 1.0,
                'frequency': new_freq if result else 1,
                'last_used': current_time.isoformat()
            }
            
        except Exception as e:
            pass  # Error updating pattern - continue
    
    def get_personalized_importance_score(self, email_data):
        """חישוב ציון חשיבות מותאם אישית מתקדם"""
        base_score = 0.5
        
        try:
            subject = email_data.get('subject', '').lower()
            sender = email_data.get('sender', '').lower()
            
            # למידה ממילות מפתח
            keywords = self.extract_keywords(subject)
            keyword_score = 0
            for keyword in keywords:
                if 'keyword_importance' in self.user_patterns:
                    if keyword in self.user_patterns['keyword_importance']:
                        pattern = self.user_patterns['keyword_importance'][keyword]
                        # התחשבות בתדירות ובזמן
                        frequency_factor = min(1.0, pattern['frequency'] / 10)
                        keyword_score += pattern['weight'] * frequency_factor * 0.1
            
            # למידה מהשולח
            sender_score = 0
            if 'sender_importance' in self.user_patterns:
                if sender in self.user_patterns['sender_importance']:
                    pattern = self.user_patterns['sender_importance'][sender]
                    frequency_factor = min(1.0, pattern['frequency'] / 5)
                    sender_score += pattern['weight'] * frequency_factor * 0.2
            
            # למידה מדומיין
            domain_score = 0
            if '@' in sender:
                domain = sender.split('@')[1]
                if 'domain_importance' in self.user_patterns:
                    if domain in self.user_patterns['domain_importance']:
                        pattern = self.user_patterns['domain_importance'][domain]
                        frequency_factor = min(1.0, pattern['frequency'] / 3)
                        domain_score += pattern['weight'] * frequency_factor * 0.15
            
            # למידה מדפוסי זמן
            time_score = 0
            from datetime import datetime
            current_time = datetime.now()
            hour = current_time.hour
            day_of_week = current_time.weekday()
            
            if 'time_importance' in self.user_patterns:
                hour_key = f'hour_{hour}'
                day_key = f'day_{day_of_week}'
                
                if hour_key in self.user_patterns['time_importance']:
                    pattern = self.user_patterns['time_importance'][hour_key]
                    time_score += pattern['weight'] * 0.05
                
                if day_key in self.user_patterns['time_importance']:
                    pattern = self.user_patterns['time_importance'][day_key]
                    time_score += pattern['weight'] * 0.05
            
            # חישוב ציון סופי
            final_score = base_score + keyword_score + sender_score + domain_score + time_score
            
            # הגבלת הציון לטווח 0-1
            return max(0.0, min(1.0, final_score))
            
        except Exception as e:
            pass  # Error calculating personalized importance - continue
            return base_score
    
    def get_personalized_category(self, email_data):
        """חישוב קטגוריה מותאמת אישית"""
        try:
            subject = email_data.get('subject', '').lower()
            keywords = self.extract_keywords(subject)
            
            # חיפוש דפוסי קטגוריה
            if 'keyword_category' in self.user_patterns:
                category_scores = defaultdict(float)
                
                for keyword in keywords:
                    if keyword in self.user_patterns['keyword_category']:
                        pattern = self.user_patterns['keyword_category'][keyword]
                        category_scores[pattern['value']] += pattern['weight']
                
                if category_scores:
                    # החזרת הקטגוריה עם הציון הגבוה ביותר
                    return max(category_scores.items(), key=lambda x: x[1])[0]
            
            # ברירת מחדל
            return 'work'
            
        except Exception as e:
            pass  # Error calculating personalized category - continue
            return 'work'
    
    def get_user_learning_stats(self):
        """קבלת סטטיסטיקות למידה"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # סטטיסטיקות משוב
            cursor.execute('SELECT COUNT(*) FROM user_feedback')
            total_feedback = cursor.fetchone()[0]
            
            cursor.execute('SELECT COUNT(*) FROM user_feedback WHERE feedback_type = "importance"')
            importance_feedback = cursor.fetchone()[0]
            
            cursor.execute('SELECT COUNT(*) FROM user_feedback WHERE feedback_type = "category"')
            category_feedback = cursor.fetchone()[0]
            
            # דפוסי למידה
            cursor.execute('SELECT COUNT(*) FROM user_patterns')
            total_patterns = cursor.fetchone()[0]
            
            conn.close()
            
            # חישוב דיוק ורמת למידה
            accuracy_rate = self.get_learning_accuracy()
            learning_level = self.get_learning_level()
            
            return {
                'total_feedback': total_feedback,
                'importance_feedback': importance_feedback,
                'category_feedback': category_feedback,
                'total_patterns': total_patterns,
                'learning_active': total_feedback > 0,
                'accuracy_rate': accuracy_rate,
                'learning_level': learning_level
            }
            
        except Exception as e:
            pass  # Error getting statistics - continue
            return {
                'total_feedback': 0,
                'importance_feedback': 0,
                'category_feedback': 0,
                'total_patterns': 0,
                'learning_active': False,
                'accuracy_rate': 0,
                'learning_level': 0
            }
    
    def get_learning_accuracy(self):
        """חישוב דיוק למידה מתקדם"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # קבלת כל המשובים
            cursor.execute('''
                SELECT user_importance_score, ai_importance_score, created_at
                FROM user_feedback 
                WHERE user_importance_score IS NOT NULL AND ai_importance_score IS NOT NULL
                ORDER BY created_at DESC
            ''')
            
            feedbacks = cursor.fetchall()
            conn.close()
            
            if not feedbacks:
                return 0
            
            # חישוב דיוק עם התחשבות בזמן
            total_accuracy = 0
            total_weight = 0
            
            for i, (user_score, ai_score, created_at) in enumerate(feedbacks):
                # משקל זמן - משובים חדשים יותר חשובים יותר
                time_weight = 1.0 - (i * 0.01)  # ירידה הדרגתית
                time_weight = max(0.1, time_weight)
                
                # חישוב דיוק למשוב זה
                difference = abs(user_score - ai_score)
                accuracy = max(0, 1.0 - difference)
                
                total_accuracy += accuracy * time_weight
                total_weight += time_weight
            
            return total_accuracy / total_weight if total_weight > 0 else 0
            
        except Exception as e:
            return 0
    
    def get_learning_level(self):
        """חישוב רמת למידה מתקדמת"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # ספירת משובים
            cursor.execute('SELECT COUNT(*) FROM user_feedback')
            total_feedback = cursor.fetchone()[0]
            
            # ספירת דפוסים
            cursor.execute('SELECT COUNT(*) FROM user_patterns')
            total_patterns = cursor.fetchone()[0]
            
            # ספירת דפוסים פעילים (בחודש האחרון)
            cursor.execute('''
                SELECT COUNT(*) FROM user_patterns 
                WHERE last_used > datetime('now', '-30 days')
            ''')
            active_patterns = cursor.fetchone()[0]
            
            conn.close()
            
            # חישוב רמת למידה
            if total_feedback == 0:
                return 0
            
            # בסיס על כמות משובים
            feedback_level = min(100, (total_feedback / 50) * 50)
            
            # בונוס על דפוסים פעילים
            pattern_bonus = min(30, (active_patterns / 20) * 30)
            
            # בונוס על מגוון דפוסים
            diversity_bonus = min(20, (total_patterns / 50) * 20)
            
            return int(feedback_level + pattern_bonus + diversity_bonus)
            
        except Exception as e:
            return 0
    
    def export_user_profile(self):
        """ייצוא פרופיל משתמש"""
        try:
            profile_export = {
                'user_patterns': self.user_patterns,
                'profile_data': self.profile_data,
                'export_date': datetime.now().isoformat(),
                'stats': self.get_user_learning_stats()
            }
            
            with open('user_profile_backup.json', 'w', encoding='utf-8') as f:
                json.dump(profile_export, f, ensure_ascii=False, indent=2)
            
            # print("User profile exported successfully")
            return True
            
        except Exception as e:
            pass  # Error exporting profile - continue
            return False
    
    def import_user_profile(self, file_path):
        """ייבוא פרופיל משתמש"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                profile_data = json.load(f)
            
            self.user_patterns = profile_data.get('user_patterns', {})
            self.profile_data = profile_data.get('profile_data', {})
            
            # print("User profile imported successfully")
            return True
            
        except Exception as e:
            pass  # Error importing profile - continue
            return False
    
    def get_sender_importance(self, sender):
        """קבלת חשיבות שולח מהפרופיל"""
        if not sender:
            return 0.0
        
        try:
            # בדיקה בדפוסי למידה קיימים
            if 'sender_importance' in self.user_patterns:
                if sender.lower() in self.user_patterns['sender_importance']:
                    pattern = self.user_patterns['sender_importance'][sender.lower()]
                    return float(pattern.get('weight', 0.5))
            
            # בדיקה במילות מפתח חשובות
            important_senders = ['manager', 'boss', 'מנהל', 'hr', 'it', 'microsoft', 'azure', 'security', 'admin', 'ceo', 'cto']
            for important_sender in important_senders:
                if important_sender.lower() in sender.lower():
                    return 0.8
            
            return 0.5
            
        except Exception as e:
            pass  # Error calculating sender importance - continue
            return 0.5
    
    def get_important_keywords(self):
        """קבלת מילות מפתח חשובות מהפרופיל"""
        keywords = {
            'urgent': 0.9,
            'דחוף': 0.9,
            'important': 0.8,
            'חשוב': 0.8,
            'meeting': 0.7,
            'פגישה': 0.7,
            'deadline': 0.8,
            'תאריך יעד': 0.8,
            'review': 0.6,
            'בדוק': 0.6,
            'reply': 0.5,
            'תגובה': 0.5
        }
        
        try:
            # הוספת מילות מפתח מהפרופיל
            if 'keyword_importance' in self.user_patterns:
                for keyword, pattern in self.user_patterns['keyword_importance'].items():
                    keywords[keyword] = float(pattern.get('weight', 0.5))
            
            return keywords
            
        except Exception as e:
            pass  # Error getting keywords - continue
            return keywords
    
    def get_category_importance(self, category):
        """קבלת חשיבות קטגוריה מהפרופיל"""
        if not category:
            return 0.5
        
        try:
            # בדיקה בדפוסי למידה קיימים
            if 'category_importance' in self.user_patterns:
                if category.lower() in self.user_patterns['category_importance']:
                    pattern = self.user_patterns['category_importance'][category.lower()]
                    return float(pattern.get('weight', 0.5))
            
            # חשיבות ברירת מחדל לפי קטגוריה
            default_importance = {
                'urgent': 0.9,
                'meeting': 0.8,
                'project': 0.7,
                'report': 0.6,
                'admin': 0.5,
                'work': 0.5,
                'personal': 0.3,
                'marketing': 0.2,
                'system': 0.6,
                'notification': 0.4
            }
            
            return default_importance.get(category.lower(), 0.5)
            
        except Exception as e:
            pass  # Error calculating category importance - continue
            return 0.5
    
    def get_all_category_importance(self):
        """קבלת חשיבות כל הקטגוריות מהפרופיל"""
        # חשיבות ברירת מחדל לפי קטגוריה
        default_importance = {
            'urgent': 0.9,
            'meeting': 0.8,
            'project': 0.7,
            'report': 0.6,
            'admin': 0.5,
            'work': 0.5,
            'personal': 0.3,
            'marketing': 0.2,
            'system': 0.6,
            'notification': 0.4
        }
        
        try:
            # הוספת קטגוריות מהפרופיל
            if 'category_importance' in self.user_patterns:
                for category, pattern in self.user_patterns['category_importance'].items():
                    default_importance[category] = float(pattern.get('weight', 0.5))
            
            return default_importance
            
        except Exception as e:
            pass  # Error getting category importance - continue
            return default_importance