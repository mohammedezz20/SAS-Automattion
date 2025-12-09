#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAS Form Automation - Multi-Browser Support
Works with Chrome, Firefox, and Edge
Auto-detects and uses available browser
"""

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
import csv
from datetime import datetime
import sys
import os


class SASFormAutomator:
    def __init__(self, form_url, excel_file, browser_choice='auto'):
        """
        Initialize the automator
        browser_choice: 'auto', 'chrome', 'firefox', or 'edge'
        """
        self.form_url = form_url
        self.excel_file = excel_file
        self.browser_choice = browser_choice
        self.results = []
        self.logs = []
        self.stop_flag = False
        self.driver = None
        self.browser_name = None

        # Start browser automatically
        self.setup_driver()

    def log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {level}: {message}"
        self.logs.append(log_message)
        print(log_message)

    def setup_driver(self):
        """Setup and launch browser - tries multiple browsers if auto mode"""
        self.log("Setting up browser...")

        browsers_to_try = []

        if self.browser_choice == 'auto':
            browsers_to_try = ['chrome', 'edge', 'firefox']
        else:
            browsers_to_try = [self.browser_choice]

        last_error = None

        for browser in browsers_to_try:
            try:
                if browser == 'chrome':
                    self.log("Trying Chrome...")
                    options = webdriver.ChromeOptions()
                    options.add_argument("--start-maximized")
                    options.add_argument(
                        "--disable-blink-features=AutomationControlled")
                    options.add_experimental_option(
                        "excludeSwitches", ["enable-automation"])
                    options.add_experimental_option(
                        "useAutomationExtension", False)
                    options.add_argument("--no-sandbox")
                    options.add_argument("--disable-dev-shm-usage")
                    options.add_argument("--disable-gpu")
                    options.add_argument("--log-level=3")

                    service = ChromeService(ChromeDriverManager().install())
                    self.driver = webdriver.Chrome(
                        service=service, options=options)
                    self.browser_name = "Chrome"

                elif browser == 'firefox':
                    self.log("Trying Firefox...")
                    options = webdriver.FirefoxOptions()
                    options.add_argument("--width=1920")
                    options.add_argument("--height=1080")

                    service = FirefoxService(GeckoDriverManager().install())
                    self.driver = webdriver.Firefox(
                        service=service, options=options)
                    self.driver.maximize_window()
                    self.browser_name = "Firefox"

                elif browser == 'edge':
                    self.log("Trying Edge...")
                    options = webdriver.EdgeOptions()
                    options.add_argument("--start-maximized")
                    options.add_argument(
                        "--disable-blink-features=AutomationControlled")
                    options.add_experimental_option(
                        "excludeSwitches", ["enable-automation"])
                    options.add_argument("--no-sandbox")
                    options.add_argument("--disable-dev-shm-usage")
                    options.add_argument("--log-level=3")

                    service = EdgeService(
                        EdgeChromiumDriverManager().install())
                    self.driver = webdriver.Edge(
                        service=service, options=options)
                    self.browser_name = "Edge"

                # Hide webdriver property
                try:
                    self.driver.execute_script(
                        "Object.defineProperty(navigator, 'webdriver', {get: () => false});")
                except:
                    pass

                self.log(f"✓ Browser setup successful: {self.browser_name}")
                return

            except Exception as e:
                last_error = e
                self.log(f"✗ Failed to launch {browser}: {str(e)}", "WARNING")
                continue

        # If all browsers failed
        error_msg = f"Failed to launch any browser. Last error: {last_error}"
        self.log(error_msg, "ERROR")
        raise Exception(error_msg)

    def read_excel(self):
        """Read student data from Excel file"""
        self.log("Reading Excel file...")
        try:
            workbook = openpyxl.load_workbook(self.excel_file)
            sheet = workbook.active

            data = []
            headers = [cell.value for cell in sheet[1]]
            self.log(f"Column headers: {headers}")

            # Find column indices
            first_name_col = last_name_col = email_col = cert_name_col = cert_link_col = badge_opt_col = None

            for i, header in enumerate(headers):
                if header == 'First Name':
                    first_name_col = i
                elif header == 'Last Name':
                    last_name_col = i
                elif header == 'Email':
                    email_col = i
                elif header == 'Certificate Name':
                    cert_name_col = i
                elif header == 'Certificate Link':
                    cert_link_col = i
                elif header == 'Badge Opt-In':
                    badge_opt_col = i

            for row in sheet.iter_rows(min_row=2, values_only=True):
                if all(cell is None for cell in row):
                    continue

                # Read data with None protection
                first_name = str(row[first_name_col]).strip() if first_name_col is not None and first_name_col < len(
                    row) and row[first_name_col] else "Unknown"
                last_name = str(row[last_name_col]).strip() if last_name_col is not None and last_name_col < len(
                    row) and row[last_name_col] else "User"
                email = str(row[email_col]).strip() if email_col is not None and email_col < len(
                    row) and row[email_col] else "noemail@example.com"
                cert_name = str(row[cert_name_col]).strip() if cert_name_col is not None and cert_name_col < len(
                    row) and row[cert_name_col] else ""
                cert_link = str(row[cert_link_col]).strip(
                ) if cert_link_col is not None and cert_link_col < len(row) and row[cert_link_col] else ""

                # Handle Badge Opt-In intelligently
                raw_badge = row[badge_opt_col] if badge_opt_col is not None and badge_opt_col < len(
                    row) else None
                if raw_badge in [None, "", " ", "None", "none"]:
                    badge_final = "yes"  # Empty = Yes automatically
                else:
                    badge_clean = str(raw_badge).strip().lower()
                    badge_final = "yes" if badge_clean in [
                        'yes', 'y', '1', 'true', 'ok'] else "no"

                if not cert_link:
                    self.log(
                        f"Skipping student {first_name} {last_name} - No certificate link", "WARNING")
                    continue

                row_data = {
                    "firstName": first_name,
                    "lastName": last_name,
                    "email": email,
                    "certificationName": cert_name,
                    "certificationLink": cert_link,
                    "badgeOptIn": badge_final
                }
                data.append(row_data)

            self.log(f"Successfully read {len(data)} students")
            return data

        except Exception as e:
            self.log(f"Error reading Excel: {str(e)}", "ERROR")
            return []

    def fill_form(self, student_data):
        """Fill one form for a specific student"""
        try:
            self.log(f"\n{'='*70}")
            self.log(
                f"Processing: {student_data['firstName']} {student_data['lastName']} | {student_data['email']}")
            self.log(f"Link: {student_data['certificationLink'][:60]}...")

            # Make sure browser is running
            if self.driver is None:
                self.setup_driver()

            self.driver.get(student_data['certificationLink'])
            time.sleep(3)

            # Wait for fields
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//input[@type='text']"))
            )

            # Fill text fields
            inputs = self.driver.find_elements(
                By.XPATH, "//input[@type='text']")
            if len(inputs) < 3:
                raise Exception(
                    f"Expected at least 3 text fields, found: {len(inputs)}")

            inputs[0].clear()
            inputs[0].send_keys(student_data['firstName'])
            inputs[1].clear()
            inputs[1].send_keys(student_data['lastName'])
            inputs[2].clear()
            inputs[2].send_keys(student_data['email'])

            # Select Badge Opt-In intelligently
            target = "Y#1" if student_data['badgeOptIn'] == "yes" else "N#2"
            choice = "Yes" if target == "Y#1" else "No"

            radio = WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//input[@type='radio' and contains(@value, '{target}')]"))
            )
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", radio)
            time.sleep(0.7)
            self.driver.execute_script("arguments[0].click();", radio)
            time.sleep(1)
            self.log(f"Selected badge option: {choice}")

            # Submit safely
            time.sleep(1)
            submit_btn = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[normalize-space()='Submit']"))
            )
            self.driver.execute_script("arguments[0].click();", submit_btn)
            self.log("Submit button clicked successfully")
            time.sleep(4)

            self.log("Form submitted successfully!", "SUCCESS")

            return {
                "email": student_data['email'],
                "firstName": student_data['firstName'],
                "lastName": student_data['lastName'],
                "certificationName": student_data.get('certificationName', ''),
                "status": "Success",
                "message": "Completed successfully"
            }

        except Exception as e:
            error_msg = f"Failed: {str(e)}"
            self.log(error_msg, "ERROR")
            try:
                if self.driver:
                    self.driver.save_screenshot(
                        f"ERROR_{student_data['email'].replace('@', '_')}.png")
            except:
                pass
            return {
                "email": student_data['email'],
                "firstName": student_data['firstName'],
                "lastName": student_data['lastName'],
                "certificationName": student_data.get('certificationName', ''),
                "status": "Failed",
                "message": error_msg
            }

    def close_driver(self):
        """Close browser"""
        try:
            if self.driver:
                self.driver.quit()
                self.log("Browser closed")
        except:
            pass

    def save_results(self):
        """Save results to CSV file"""
        filename = f"SAS_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(
                f, fieldnames=["email", "firstName", "lastName", "certificationName", "status", "message"])
            writer.writeheader()
            writer.writerows(self.results)
        self.log(f"Results saved to: {filename}")

    def __del__(self):
        """Cleanup resources when object is deleted"""
        self.close_driver()
