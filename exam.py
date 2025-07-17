import pytest
from playwright.sync_api import sync_playwright, Page, expect

def launch_browser():
    with sync_playwright() as p:
        return p.chromium.launch(headless=False)  # ใช้ Chrome

# 1. เข้าสู่ระบบได้ (01_login01_pos01)
def 01_login01_pos01(page: Page):
    page.goto("https://reg.rmutk.ac.th/registrar/login.asp")
    page.locator('input[name="f_uid"]').fill("")
    page.locator('input[name="f_pwd"]').fill("")
    page.locator('input[type="submit"][value=" เข้าสู่ระบบ "]').click()
    page.wait_for_selector("text=เปลี่ยนรหัสผ่าน"), timeout=60000)

# 2. เปลี่ยนรหัสนักศึกษาไม่ได้ เนื่องจากรหัสอยู่ในระบบแล้ว (02_changeid01_neg01)
def 02_changeid01_neg01(page: Page):
    page.goto("https://your-webapp-url.com/change_id")
    page.locator('input[name="student_id"]').fill("")
    page.locator('input[type="submit"][value=" เปลี่ยนรหัสประจำตัว "]').click()
    error_message = page.wait_for_selector("text=รหัสประจำตัวนี้มีคนใช้แล้ว", timeout=60000)
    assert error_message.is_visible()

# 3. เพิ่มกลุ่มไม่ได้ เนื่องจากรหัสกลุ่มสั้นเกินไป (03_addgroup01_neg01)
def 03_addgroup01_neg01(page: Page):
    page.goto("https://your-webapp-url.com/add_group")
    page.locator('input[name="group_id"]').fill("")
    page.locator('input[type="submit"][value=" เข้าร่วม "]').click()
    error_message = page.wait_for_selector("text=ไอดีของกลุ่มสอบต้องมี 6 หลักเท่านั้น!", timeout=60000)
    assert error_message.is_visible()

# 4. เพิ่มกลุ่มไม่ได้ เนื่องจากรหัสกลุ่มไม่มีอยู่ในระบบ (04_addgroup01_neg02)
def 04_addgroup01_neg02(page: Page):
    page.goto("https://your-webapp-url.com/add_group")
    page.locator('input[name="group_id"]').fill("")
    page.locator('input[type="submit"][value=" เข้าร่วม "]').click()
    error_message = page.wait_for_selector("text=ไม่พบไอดีของกลุ่ม!", timeout=60000)
    assert error_message.is_visible()

# 5. เพิ่มกลุ่มได้ (05_addgroup01_pos01)
def 05_addgroup01_pos01(page: Page):
    page.goto("https://your-webapp-url.com/add_group")
    page.locator('input[name="group_id"]').fill("")
    page.locator('input[type="submit"][value=" เข้าร่วม "]').click()
    #success_message = page.wait_for_selector("text=เพิ่มกลุ่มสำเร็จ", timeout=60000)
    #assert success_message.is_visible()
    page.wait_for_url(timeout=10000)  #  รอ 10 วินาที

# 6. ออกจากระบบได้ (06_logout_pos01)
def 06_logout_pos01(page: Page):
    page.goto("https://reg.rmutk.ac.th/registrar/login.asp")
    page.locator('input[name="f_uid"]').fill("")
    page.locator('input[name="f_pwd"]').fill("")
    page.locator('input[type="submit"][value=" เข้าสู่ระบบ "]').click()
    
    page.locator('input[type="submit"][value="  "]').click()
    page.locator('input[type="submit"][value=" ออกจากระบบ "]').click()
    page.locator('input[type="submit"][value=" ใช่ "]').click()
    
    logout_message = page.wait_for_selector("text=เข้าสู่ระบบ", timeout=60000) #    ไม่แน่ใจทำให้ด้วย
    assert logout_message.is_visible()

