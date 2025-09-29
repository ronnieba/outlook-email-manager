import sqlite3

conn = sqlite3.connect('email_manager.db')
cursor = conn.cursor()
cursor.execute("SELECT id, subject, action_items FROM emails WHERE action_items IS NOT NULL AND action_items != '[]' LIMIT 5")
rows = cursor.fetchall()

print("מיילים עם action_items:")
for row in rows:
    print(f'ID: {row[0]}, Subject: {row[1][:50]}..., Action Items: {row[2]}')

conn.close()




