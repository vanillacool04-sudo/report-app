import sqlite3

conn = sqlite3.connect("report.db")
cursor = conn.cursor()

# เพิ่มคอลัมน์เลขงาน
cursor.execute("ALTER TABLE reports ADD COLUMN work_no TEXT")

# บังคับไม่ให้เลขงานซ้ำ
cursor.execute("CREATE UNIQUE INDEX idx_work_no ON reports(work_no)")

conn.commit()
conn.close()

print("อัปเกรดฐานข้อมูลเรียบร้อย")
