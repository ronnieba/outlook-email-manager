"""
Wrapper script להפעלת השרת עם השתקת stderr מלאה
זה השיטה היעילה ביותר להסתיר הודעות Google AI
"""
import sys
import os

if __name__ == '__main__':
    # סגירת stderr לפני כל דבר
    sys.stderr.close()
    sys.stderr = open(os.devnull, 'w')
    
    # import והרצת השרת
    import app_with_ai

