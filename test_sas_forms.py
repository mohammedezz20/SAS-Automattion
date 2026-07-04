#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAS Form Testing Script
Fills the form and selects Yes/No without submitting.
Usage:
  python test_sas_forms.py
  python test_sas_forms.py "http://go.sas.com/evals?serviceid=AAMBFSASBA25"
"""

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import sys

from sas_automation import (
    _badge_radio_selected,
    _fill_text_fields,
    _locate_text_fields,
    _select_badge_option,
    _wait_for_page_ready,
)

DEFAULT_FORM_URL = "http://go.sas.com/evals?serviceid=AAMBFSASBA25"


def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=options)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => false});"
    )
    return driver


def normalize_badge_opt_in(badge_raw):
    badge_input = (
        str(badge_raw).strip().lower() if badge_raw is not None else "yes"
    )
    if badge_input in ["yes", "y", "1", "true", "نعم", "", " "]:
        return "yes"
    return "no"


def get_first_student(form_url):
    if os.path.exists("data.xlsx"):
        wb = openpyxl.load_workbook("data.xlsx")
        sheet = wb.active
        row = sheet[2]
        return {
            "firstName": str(row[0].value or "TestFirst").strip(),
            "lastName": str(row[1].value or "TestLast").strip(),
            "email": str(row[2].value or "test@example.com").strip(),
            "certificationLink": str(row[4].value or form_url).strip() or form_url,
            "badgeOptIn": normalize_badge_opt_in(
                row[5].value if sheet.max_column >= 6 and row[5].value else None
            ),
        }

    return {
        "firstName": "TestFirst",
        "lastName": "TestLast",
        "email": "test@example.com",
        "certificationLink": form_url,
        "badgeOptIn": "yes",
    }


def main():
    form_url = sys.argv[1].strip() if len(sys.argv) > 1 else DEFAULT_FORM_URL
    student = get_first_student(form_url)

    print("=" * 70)
    print("           SAS FORM TESTING SCRIPT")
    print("=" * 70)
    print(f"الطالب: {student['firstName']} {student['lastName']}")
    print(f"الإيميل: {student['email']}")
    print(f"رابط الشهادة: {student['certificationLink']}")
    print(f"Badge Opt-In: {student['badgeOptIn']}")
    print("-" * 70)

    driver = setup_driver()
    try:
        driver.get(student["certificationLink"])
        print("جاري فتح الرابط...")
        _wait_for_page_ready(driver)

        wait = WebDriverWait(driver, 30)
        first, last, email = _locate_text_fields(driver, wait)
        _fill_text_fields(first, last, email, student)
        print("تم ملء الاسم والإيميل بنجاح")

        choice = _select_badge_option(driver, wait, student["badgeOptIn"])
        if not _badge_radio_selected(driver, student["badgeOptIn"]):
            raise Exception(f"Badge option {choice} is not selected on the form")
        print(f"تم اختيار الشارة: {choice} (verified)")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[normalize-space()='Submit']")
            )
        )
        print("زر Submit موجود وجاهز (لم يتم الضغط عليه)")

        print("=" * 70)
        print("الفورم اتملت كاملة وصحيحة")
        print("تأكد بنفسك من الشاشة... البرنامج هيستنى 60 ثانية")
        print("=" * 70)
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
        print("تم!")


if __name__ == "__main__":
    main()
