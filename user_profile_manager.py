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
            print(f"✅ User profile loaded: {len(self.user_patterns)} patterns, {len(self.profile_data)} preferences")
            
        except Exception as e:
            print(f"❌ שגיאה בטעינת פרופיל: {e}")
    
    def record_user_feedback(self, email_data, feedback_type, user_value, ai_value=None):
        """רישום משוב משתמש"""
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
            
            print(f"✅ נרשם משוב: {feedback_type} = {user_value}")
            
        except Exception as e:
            print(f"❌ שגיאה ברישום משוב: {e}")
    
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
            print(f"❌ שגיאה בעדכון דפוסים: {e}")
    
    def extract_keywords(self, text):
        """חילוץ מילות מפתח מטקסט"""
        # מילות מפתח חשובות בעברית ובאנגלית
        important_words = [
            'חשוב', 'דחוף', 'urgent', 'important', 'meeting', 'פגישה',
            'azure', 'microsoft', 'security', 'alert', 'error', 'שגיאה',
            'manager', 'מנהל', 'hr', 'it', 'admin', 'boss', 'בוס',
            'deadline', 'תאריך', 'project', 'פרויקט', 'report', 'דוח'
        ]
        
        keywords = []
        text_lower = text.lower()
        for word in important_words:
            if word in text_lower:
                keywords.append(word)
        
        return keywords
    
    def update_pattern(self, pattern_type, pattern_key, value):
        """עדכון דפוס למידה"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # בדיקה אם הדפוס קיים
            cursor.execute('''
                SELECT weight, frequency FROM user_patterns 
                WHERE pattern_type = ? AND pattern_key = ?
            ''', (pattern_type, pattern_key))
            
            result = cursor.fetchone()
            
            if result:
                # עדכון דפוס קיים
                old_weight, old_freq = result
                new_freq = old_freq + 1
                
                # חישוב משקל חדש (ממוצע משוקלל)
                if isinstance(value, (int, float)):
                    new_weight = (old_weight * old_freq + value) / new_freq
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
                'frequency': new_freq if result else 1
            }
            
        except Exception as e:
            print(f"❌ שגיאה בעדכון דפוס: {e}")
    
    def get_personalized_importance_score(self, email_data):
        """חישוב ציון חשיבות מותאם אישית"""
        base_score = 0.5
        
        try:
            subject = email_data.get('subject', '').lower()
            sender = email_data.get('sender', '').lower()
            
            # למידה ממילות מפתח
            keywords = self.extract_keywords(subject)
            for keyword in keywords:
                if 'keyword_importance' in self.user_patterns:
                    if keyword in self.user_patterns['keyword_importance']:
                        pattern = self.user_patterns['keyword_importance'][keyword]
                        base_score += pattern['weight'] * 0.1
            
            # למידה מהשולח
            if 'sender_importance' in self.user_patterns:
                if sender in self.user_patterns['sender_importance']:
                    pattern = self.user_patterns['sender_importance'][sender]
                    base_score += pattern['weight'] * 0.2
            
            # הגבלת הציון לטווח 0-1
            return max(0.0, min(1.0, base_score))
            
        except Exception as e:
            print(f"❌ שגיאה בחישוב חשיבות מותאמת: {e}")
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
            print(f"❌ שגיאה בחישוב קטגוריה מותאמת: {e}")
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
            
            return {
                'total_feedback': total_feedback,
                'importance_feedback': importance_feedback,
                'category_feedback': category_feedback,
                'total_patterns': total_patterns,
                'learning_active': total_feedback > 0
            }
            
        except Exception as e:
            print(f"❌ שגיאה בקבלת סטטיסטיקות: {e}")
            return {
                'total_feedback': 0,
                'importance_feedback': 0,
                'category_feedback': 0,
                'total_patterns': 0,
                'learning_active': False
            }
    
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
            
            print("✅ פרופיל משתמש יוצא בהצלחה")
            return True
            
        except Exception as e:
            print(f"❌ שגיאה בייצוא פרופיל: {e}")
            return False
    
    def import_user_profile(self, file_path):
        """ייבוא פרופיל משתמש"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                profile_data = json.load(f)
            
            self.user_patterns = profile_data.get('user_patterns', {})
            self.profile_data = profile_data.get('profile_data', {})
            
            print("✅ פרופיל משתמש יובא בהצלחה")
            return True
            
        except Exception as e:
            print(f"❌ שגיאה בייבוא פרופיל: {e}")
            return False
    
    def get_sender_importance(self, sender):
        """קבלת חשיבות שולח מהפרופיל"""
        if not sender:
            return 0.0
        
        # בדיקה בפרופיל
        sender_key = f"sender_{sender.lower()}"
        if sender_key in self.profile_data:
            return self.profile_data[sender_key].get('importance', 0.5)
        
        # בדיקה במילות מפתח חשובות
        important_senders = ['manager', 'boss', 'מנהל', 'hr', 'it', 'microsoft', 'azure', 'security', 'admin', 'ceo', 'cto']
        for important_sender in important_senders:
            if important_sender.lower() in sender.lower():
                return 0.8
        
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
        
        # הוספת מילות מפתח מהפרופיל
        for key, value in self.profile_data.items():
            if key.startswith('keyword_'):
                keyword = key.replace('keyword_', '')
                keywords[keyword] = value.get('weight', 0.5)
        
        return keywords
    
    def get_category_importance(self, category):
        """קבלת חשיבות קטגוריה מהפרופיל"""
        if not category:
            return 0.5
        
        category_key = f"category_{category.lower()}"
        if category_key in self.profile_data:
            return self.profile_data[category_key].get('importance', 0.5)
        
        # חשיבות ברירת מחדל לפי קטגוריה
        default_importance = {
            'urgent': 0.9,
            'meeting': 0.8,
            'project': 0.7,
            'report': 0.6,
            'admin': 0.5,
            'work': 0.5
        }
        
        return default_importance.get(category.lower(), 0.5)
    
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
        
        # הוספת קטגוריות מהפרופיל
        for key, value in self.profile_data.items():
            if key.startswith('category_'):
                category = key.replace('category_', '')
                default_importance[category] = value.get('importance', 0.5)
        
        return default_importance