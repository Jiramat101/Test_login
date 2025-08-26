# Lxml works with the encoding on RMUTK Reg website (Windows-874)
# Credit: https://medium.com/@datajournal/web-scraping-with-python-lxml-e9662136cbf5
# Credit: https://reqbin.com/code/python/6mwlgbqa/python-requests-download-file-example
# Credit: https://stackoverflow.com/questions/55508303/how-to-write-a-list-of-list-into-excel-using-python
# Credit: https://xlsxwriter.readthedocs.io/tutorial01.html
# Credit: https://www.geeksforgeeks.org/how-to-convert-python-dictionary-to-json/
# Credit: https://pynative.com/python-get-execution-time-of-program/
# Credit: https://stackoverflow.com/questions/21965484/timeout-for-python-requests-get-entire-response

# ----------------- Library for Task01
import requests
from lxml import html

# ----------------- Library for Task02
import pandas as pd


# ------------------------------------------------------------- Task01

# เข้า Developer Tool หาค่า Cookies ของ reg.rmutk.ac.th แล้วนำมาใส่ด้านล่าง
# ระวัง! โดยทั่วไป ค่า Cookies ควรจะเป็นความลับ  คือไม่ควรแสดงสิ่งนี้ให้คนอื่นเห็น
cookies = {
    รหัสคุกกี้ f12
}

# เข้า Browser หน้าตรวจสอบจบ แล้วนำ URL มาใส่ด้านล่าง
url = "https://reg.rmutk.ac.th/registrar/graduate_check.asp?avs114128872=8"
response = requests.get(url, timeout=60, cookies=cookies)

dom = html.fromstring(response.content)
tables = dom.xpath('.//table[contains(@class, "normalDetail")]')

if not tables:
    print("ไม่พบตารางที่มี class 'normalDetail'")
    exit()

# เลือกตารางแรกที่พบ
table = tables[0]

# ----------------- Task01: แสดงข้อมูล
# headers (กำหนดให้สอดคล้องกับคอลัมน์ในหน้าเว็บ)
headers = ["รหัสวิชา", "ชื่อวิชา", "หน่วยกิต", "เกรด", "เทอม"]

# เก็บข้อมูลในรูปแบบ list
data = []

# วนลูปดึงข้อมูลแต่ละแถวในตาราง (ข้าม header)
for tr in table[1:]:
    columns = tr.getchildren()
    row = [col.text_content().strip() for col in columns]
    
    # ตรวจสอบว่าแถวนี้มีข้อมูลที่ต้องการและมีเกรดที่ไม่เป็น '-' หรือ '--'
    if len(row) >= 5:  # ตรวจสอบว่ามีคอลัมน์เพียงพอ
        grade = row[3]  # เกรดอยู่ในคอลัมน์ที่ 4 (index 3)
        
        # ตรวจสอบเฉพาะวิชาที่มีเกรดแล้ว (ไม่ใช่ '-' หรือ '--')
        if grade and grade not in ["-", "--"]:
            data.append(row[:5])  # เก็บเฉพาะ 5 คอลัมน์แรก

# แสดงข้อมูลที่ได้
print("\nข้อมูลวิชาที่ดึงได้:")
for row in data:
    print(f"รหัสวิชา: {row[0]}, ชื่อวิชา: {row[1]}, หน่วยกิต: {row[2]}, เกรด: {row[3]}, เทอม: {row[4]}")

# ------------------------------------------------------------- Task02 (ตัวอย่างการเก็บใน DataFrame)

