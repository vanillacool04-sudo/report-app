import sqlite3

# สร้าง / เปิดไฟล์ DB
conn = sqlite3.connect("report.db")
cursor = conn.cursor()

# ===============================
# ตารางหน่วยงาน (Department Master)
# ===============================
cursor.execute("""
CREATE TABLE IF NOT EXISTS departments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    short_name TEXT NOT NULL,
    full_name TEXT NOT NULL,
    active INTEGER DEFAULT 1
)
""")

# ===============================
# ตาราง report (ถ้ายังไม่มี)
# ===============================
cursor.execute("""
CREATE TABLE IF NOT EXISTS report (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    department_id INTEGER,
    job_type TEXT,
    asset_no TEXT,
    detail TEXT,
    created_at TEXT
)
""")

# ===============================
# ใส่ข้อมูลตัวอย่างหน่วยงาน
# ===============================
cursor.execute("SELECT COUNT(*) FROM departments")
count = cursor.fetchone()[0]

if count == 0:
    cursor.executemany("""
        INSERT INTO departments (short_name, full_name)
        VALUES (?, ?)
    """, [
        ("กกม.", "กองการจัดการ"),
        ("กพร.", "กองพัฒนาระบบ"),
        ("กสส.", "กองสนับสนุนระบบ"),
        ("กยผ.", "กองยุทธศาสตร์และแผนงาน")
    ])

conn.commit()
conn.close()

print("✅ สร้างฐานข้อมูลเรียบร้อยแล้ว")




