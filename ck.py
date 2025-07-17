# Lxml works with the encoding on RMUTK Reg website (Windows-874)
# Credit: https://medium.com/@datajournal/web-scraping-with-python-lxml-e9662136cbf5
# Credit: https://reqbin.com/code/python/6mwlgbqa/python-requests-download-file-example
# Credit: https://stackoverflow.com/questions/55508303/how-to-write-a-list-of-list-into-excel-using-python
# Credit: https://xlsxwriter.readthedocs.io/tutorial01.html
# Credit: https://www.geeksforgeeks.org/how-to-convert-python-dictionary-to-json/
# Credit: https://pynative.com/python-get-execution-time-of-program/
# Credit: https://stackoverflow.com/questions/21965484/timeout-for-python-requests-get-entire-response

import requests
from lxml import html
import time
import pandas as pd
from playwright.sync_api import sync_playwright, Playwright

st = 0
def printStatus(text):
    print("----------------------- {} ({})".format(text, time.time() - st))

# ------------------------------------------------------------- Global val

st = time.time()

# ------------------------------------------------------------- PlayWright does the login UI
playwright = sync_playwright().start()

browser = playwright.chromium.launch(headless=False)
page = browser.new_page()
page.goto("https://reg.rmutk.ac.th/registrar/login.asp")
#page.screenshot(path="example.png")

# Wait for an event that proofs that log-in has been successful
# Code here
# ทำการเข้าสู่ระบบที่นี่ (กรอกฟอร์มและส่งข้อมูล)

# รอให้มีการนำทางไปยังหน้าใหม่หลังจากเข้าสู่ระบบฃ

page.wait_for_selector('b:has-text("ยินดีต้อนรับเข้าสู่ระบบบริการการศึกษา")', timeout=60000)


# Get cookies from browser
# Code here
#  ดึงคุกกี้จากบราวเซอร์หลังจากการนำทางเสร็จสิ้น
browser_cookies = page.context.cookies()
cookies = {}
for cookie in browser_cookies:
    cookies[cookie['name']] = cookie['value']
print(cookies)
#browser.close()
# รอการตอบสนองจากผู้ใช้ก่อนที่จะปิดบราวเซอร์
input("Enter to exit")

# ปิดบราวเซอร์เมื่อผู้ใช้กด Enter
browser.close()
playwright.stop()
# ------------------------------------------------------------- Scraping
