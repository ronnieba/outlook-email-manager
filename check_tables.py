import sqlite3

conn = sqlite3.connect('email_manager.db')
cursor = conn.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cursor.fetchall()

print("טבלאות במסד הנתונים:")
for table in tables:
    print(f"- {table[0]}")

conn.close()




