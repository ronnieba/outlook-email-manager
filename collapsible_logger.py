"""
מערכת לוגים עם בלוקים ניתנים לקפילה
Collapsible Logging System for Console
"""

import time
import uuid
from datetime import datetime
from typing import Dict, List, Optional

class CollapsibleLogger:
    def __init__(self, name: str = "OutlookEmailManager"):
        # פונקציה חיצונית לשליחת הודעות לקונסול
        self.console_logger = None
        # אחסון בלוקים פעילים
        self.active_blocks: Dict[str, Dict] = {}
    
    def set_console_logger(self, console_logger_func):
        """הגדרת פונקציה חיצונית לשליחת הודעות לקונסול"""
        self.console_logger = console_logger_func
    
    def _log_to_console(self, message: str, level: str = "INFO"):
        """שליחת הודעה לקונסול"""
        if self.console_logger:
            self.console_logger(message, level)
        else:
            # fallback להדפסה רגילה
            timestamp = datetime.now().strftime("%H:%M:%S")
            print(f"[{timestamp}] {message}")
    
    def start_block(self, block_name: str, description: str = "") -> str:
        """התחלת בלוק חדש"""
        block_id = str(uuid.uuid4())[:8]
        
        self.active_blocks[block_id] = {
            'name': block_name,
            'description': description,
            'start_time': time.time(),
            'messages': []
        }
        
        # הודעת פתיחה עם מזהה בלוק
        self._log_to_console(f"🔄 [{block_id}] {block_name} - התחלה", "INFO")
        if description:
            self._log_to_console(f"   📝 {description}", "INFO")
        
        return block_id
    
    def add_to_block(self, block_id: str, message: str, level: str = "INFO"):
        """הוספת הודעה לבלוק"""
        if block_id in self.active_blocks:
            self.active_blocks[block_id]['messages'].append({
                'message': message,
                'level': level,
                'time': time.time()
            })
            
            # הצגת ההודעה עם הזחה
            self._log_to_console(f"   └─ {message}", "INFO")
    
    def update_progress(self, block_id: str, current: int, total: int, item_name: str = ""):
        """עדכון התקדמות בבלוק"""
        if block_id in self.active_blocks:
            percentage = (current / total) * 100 if total > 0 else 0
            progress_bar = self._create_progress_bar(percentage)
            
            message = f"התקדמות: {current}/{total} ({percentage:.1f}%) {progress_bar}"
            if item_name:
                message += f" - {item_name}"
            
            self.add_to_block(block_id, message)
    
    def end_block(self, block_id: str, success: bool = True, summary: str = ""):
        """סיום בלוק"""
        if block_id in self.active_blocks:
            block = self.active_blocks[block_id]
            duration = time.time() - block['start_time']
            
            # הודעת סיום
            status_icon = "✅" if success else "❌"
            self._log_to_console(f"{status_icon} [{block_id}] {block['name']} - סיום ({duration:.2f} שניות)", "SUCCESS" if success else "ERROR")
            
            if summary:
                self._log_to_console(f"   📊 סיכום: {summary}", "INFO")
            
            # לא מציגים ספירת הודעות – מיותר למסך
            
            # הסרת הבלוק מהרשימה הפעילה
            del self.active_blocks[block_id]
    
    def _create_progress_bar(self, percentage: float, width: int = 20) -> str:
        """יצירת פס התקדמות"""
        filled = int((percentage / 100) * width)
        bar = "█" * filled + "░" * (width - filled)
        return f"[{bar}]"
    
    def log_info(self, message: str):
        """הודעת מידע רגילה"""
        self._log_to_console(message, "INFO")
    
    def log_warning(self, message: str):
        """הודעת אזהרה"""
        self._log_to_console(f"⚠️ {message}", "WARNING")
    
    def log_error(self, message: str):
        """הודעת שגיאה"""
        self._log_to_console(f"❌ {message}", "ERROR")
    
    def log_success(self, message: str):
        """הודעת הצלחה"""
        self._log_to_console(f"✅ {message}", "SUCCESS")

# יצירת instance גלובלי
logger = CollapsibleLogger()

# דוגמה לשימוש
if __name__ == "__main__":
    # דוגמה של העברת ציונים
    print("דוגמה למערכת לוגים עם בלוקים ניתנים לקפילה")
    print("=" * 50)
    
    # התחלת בלוק העברת ציונים
    block_id = logger.start_block(
        "העברת ציונים", 
        "מעביר ציונים ממיילים לתלמידים"
    )
    
    # סימולציה של העברת ציונים
    students = ["יוסי כהן", "שרה לוי", "דוד ישראלי", "מיכל אברהם", "אליהו רוזן"]
    total_students = len(students)
    
    for i, student in enumerate(students, 1):
        logger.add_to_block(block_id, f"מעביר ציון לתלמיד: {student}")
        logger.update_progress(block_id, i, total_students, student)
        time.sleep(0.5)  # סימולציה של זמן עיבוד
    
    # סיום הבלוק
    logger.end_block(
        block_id, 
        success=True, 
        summary=f"הועברו ציונים ל-{total_students} תלמידים בהצלחה"
    )
    
    print("\n" + "=" * 50)
    print("דוגמה נוספת - בדיקת חיבור")
    
    # דוגמה נוספת
    block_id2 = logger.start_block("בדיקת חיבור", "בודק חיבור לשרת")
    logger.add_to_block(block_id2, "מתחבר לשרת...")
    time.sleep(1)
    logger.add_to_block(block_id2, "חיבור הצליח!")
    logger.end_block(block_id2, success=True, summary="חיבור תקין")
