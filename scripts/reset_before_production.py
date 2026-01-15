"""
====================================================
⚠️  WARNING : PRODUCTION RESET SCRIPT
====================================================
ไฟล์นี้ใช้สำหรับ:
- รีเซ็ตข้อมูลทดสอบทั้งหมด
- ใช้ก่อนเปิดใช้งานจริง (Go-Live) เท่านั้น

❌ ห้าม import ไฟล์นี้ใน app.py
❌ ห้ามรันโดยไม่ตั้งใจ
====================================================
"""

import sqlite3
import os

DB_NAME = "report.db"
SIGNATURE_DIR = os.path.join("static", "signatures")

def reset_database():
    print("⚠️  คุณกำลังจะลบข้อมูลทดสอบทั้งหมด")
    confirm = input("พิมพ์ YES เพื่อยืนยัน: ")

    if confirm != "YES":
        print("❌ ยกเลิกการรีเซ็ต")
        return

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    # ลบข้อมูลทั้งหมด
    cursor.execute("DELETE FROM reports")

    # รีเซ็ต autoincrement
    cursor.execute("DELETE FROM sqlite_sequence WHERE name='reports'")

    conn.commit()
    conn.close()

    print("✅ รีเซ็ตฐานข้อมูลเรียบร้อย")

def reset_signatures():
    if not os.path.exists(SIGNATURE_DIR):
        return

    for f in os.listdir(SIGNATURE_DIR):
        try:
            os.remove(os.path.join(SIGNATURE_DIR, f))
        except:
            pass

    print("✅ ลบไฟล์ลายเซ็นเรียบร้อย")

if __name__ == "__main__":
    print("===================================")
    print("  RESET SYSTEM BEFORE PRODUCTION")
    print("===================================")
    reset_database()

    clear_sig = input("ต้องการลบไฟล์ลายเซ็นด้วยหรือไม่ (yes/no): ")
    if clear_sig.lower() == "yes":
        reset_signatures()

    print("🎉 ระบบพร้อมสำหรับใช้งานจริงแล้ว")
