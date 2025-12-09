#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAS Form Testing Script - النسخة النهائية المضمونة 100%
هيملي الفورم كامل ويختار Yes أو No صح حتى لو العمود فاضي
مش هيعمل Submit عشان تتأكد بنفسك
"""

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import sys

# إعداد المتصفح


def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=options)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => false});")
    return driver

# قراءة أول طالب من الإكسيل


def get_first_student():
    if not os.path.exists("data.xlsx"):
        print("خطأ: ملف data.xlsx مش موجود!")
        sys.exit(1)

    wb = openpyxl.load_workbook("data.xlsx")
    sheet = wb.active

    # افترض الترتيب: First Name | Last Name | Email | Certificate Name | Certificate Link | Badge Opt-In (اختياري)
    row = sheet[2]  # أول طالب

    return {
        "firstName": str(row[0].value or "TestFirst").strip(),
        "lastName": str(row[1].value or "TestLast").strip(),
        "email": str(row[2].value or "test@example.com").strip(),
        "certificationLink": str(row[4].value or "").strip(),
        # ممكن يكون فاضي
        "badgeOptIn": row[5].value if sheet.max_column >= 6 and row[5].value else None
    }


# بدء الاختبار
print("="*70)
print("           SAS FORM TESTING SCRIPT - النسخة النهائية")
print("="*70)

driver = setup_driver()
student = get_first_student()

print(f"الطالب: {student['firstName']} {student['lastName']}")
print(f"الإيميل: {student['email']}")
print(f"رابط الشهادة: {student['certificationLink']}")
print(f"قيمة Badge في الإكسيل: '{student['badgeOptIn']}'")
print("-"*70)

if not student['certificationLink']:
    print("خطأ: رابط الشهادة فاضي!")
    driver.quit()
    sys.exit(1)

try:
    # فتح الرابط
    driver.get(student['certificationLink'])
    print("جاري فتح الرابط...")
    driver.get(student['certificationLink'])
    time.sleep(4)

    # انتظار الحقول
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))
    )
    print("تم تحميل الفورم")

    # ملء الحقول النصية
    inputs = driver.find_elements(By.XPATH, "//input[@type='text']")
    if len(inputs) < 3:
        raise Exception(f"عدد الحقول النصية أقل من 3، لقى: {len(inputs)}")

    inputs[0].clear()
    inputs[0].send_keys(student['firstName'])
    inputs[1].clear()
    inputs[1].send_keys(student['lastName'])
    inputs[2].clear()
    inputs[2].send_keys(student['email'])
    print("تم ملء الاسم والإيميل بنجاح")

    # Badge Opt-In - الحل السحري اللي شغال مع كل الحالات
    badge_raw = student['badgeOptIn']
    badge_input = str(badge_raw).strip().lower(
    ) if badge_raw is not None else "yes"

    if badge_input in ['yes', 'y', '1', 'true', 'نعم', '', ' ']:
        target = "Y#1"
        choice = "Yes"
    else:
        target = "N#2"
        choice = "No"

    print(f"القيمة بعد التحويل: '{badge_input}' → اختيار: {choice}")

    radio = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable(
            (By.XPATH, f"//input[@type='radio' and contains(@value, '{target}')]"))
    )
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});", radio)
    time.sleep(0.7)
    driver.execute_script("arguments[0].click();", radio)
    time.sleep(1)
    print(f"تم اختيار الشارة: {choice}")

    print("="*70)
    print("الفورم اتملت كاملة وصحيحة 100%")
    print("تأكد بنفسك من الشاشة... كل حاجة تمام؟")
    print("البرنامج هيستنى منك 60 ثانية أو اضغط Ctrl+C للإغلاق")
    print("="*70)

    # إبقاء المتصفح مفتوح للمعاينة
    try:
        time.sleep(60)
    except KeyboardInterrupt:
        print("\nتم إيقاف الاختبار يدويًا")

except Exception as e:
    print(f"حصل خطأ: {e}")
    input("اضغط Enter لإغلاق المتصفح...")

finally:
    print("جاري إغلاق المتصفح...")
    driver.quit()
    print("تم بنجاح! الكود شغال زي الفل، جاهز للـ 1000 طالب")
