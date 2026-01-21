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
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading


class SASFormAutomator:
    def __init__(self, form_url, excel_file, browser_choice='auto', checkpoint_dir=None, restart_browser_interval=100):
        """
        Initialize the automator
        browser_choice: 'auto', 'chrome', 'firefox', or 'edge'
        checkpoint_dir: Directory to save progress checkpoints
        restart_browser_interval: Restart browser every N forms to prevent crashes
        """
        self.form_url = form_url
        self.excel_file = excel_file
        self.browser_choice = browser_choice
        self.results = []
        self.logs = []
        self.stop_flag = False
        self.driver = None
        self.browser_name = None
        self.checkpoint_dir = checkpoint_dir or "checkpoints"
        self.restart_browser_interval = restart_browser_interval
        self.forms_processed_since_restart = 0
        self.checkpoint_file = os.path.join(self.checkpoint_dir, "progress.json")
        self.results_file = os.path.join(self.checkpoint_dir, "results.csv")
        
        # Create checkpoint directory
        os.makedirs(self.checkpoint_dir, exist_ok=True)

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

    def fill_form(self, student_data, max_retries=3):
        """Fill one form for a specific student with automatic retry"""
        last_error = None
        
        for attempt in range(1, max_retries + 1):
            try:
                if attempt > 1:
                    self.log(f"Retry attempt {attempt}/{max_retries} for {student_data['email']}", "WARNING")
                    time.sleep(2)  # Wait before retry
                
                return self._fill_form_single(student_data)
                
            except Exception as e:
                last_error = e
                error_msg = f"Attempt {attempt}/{max_retries} failed: {str(e)}"
                self.log(error_msg, "WARNING" if attempt < max_retries else "ERROR")
                
                # If browser seems dead, try to restart it
                if attempt < max_retries:
                    try:
                        # Test if browser is still responsive
                        self.driver.current_url
                    except:
                        self.log("Browser appears unresponsive, restarting...", "WARNING")
                        self.close_driver()
                        time.sleep(2)
                        self.setup_driver()
        
        # All retries failed
        error_msg = f"Failed after {max_retries} attempts: {str(last_error)}"
        self.log(error_msg, "ERROR")
        try:
            if self.driver:
                screenshot_path = os.path.join(self.checkpoint_dir, f"ERROR_{student_data['email'].replace('@', '_')}.png")
                self.driver.save_screenshot(screenshot_path)
        except:
            pass
        
        # Increment counter even on failure
        self.forms_processed_since_restart += 1
        
        return {
            "email": student_data['email'],
            "firstName": student_data['firstName'],
            "lastName": student_data['lastName'],
            "certificationName": student_data.get('certificationName', ''),
            "status": "Failed",
            "message": error_msg
        }
    
    def _fill_form_single(self, student_data):
        """Fill one form for a specific student (single attempt)"""
        try:
            self.log(f"\n{'='*70}")
            self.log(
                f"Processing: {student_data['firstName']} {student_data['lastName']} | {student_data['email']}")
            self.log(f"Link: {student_data['certificationLink'][:60]}...")

            # Make sure browser is running
            if self.driver is None:
                self.setup_driver()

            self.driver.get(student_data['certificationLink'])
            time.sleep(2)  # Reduced from 3 to 2 seconds

            # Wait for fields
            WebDriverWait(self.driver, 25).until(  # Reduced from 30 to 25 seconds
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
            time.sleep(0.5)  # Reduced from 0.7 to 0.5 seconds
            self.driver.execute_script("arguments[0].click();", radio)
            time.sleep(0.5)  # Reduced from 1 to 0.5 seconds
            self.log(f"Selected badge option: {choice}")

            # Submit safely
            time.sleep(0.5)  # Reduced from 1 to 0.5 seconds
            submit_btn = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[normalize-space()='Submit']"))
            )
            self.driver.execute_script("arguments[0].click();", submit_btn)
            self.log("Submit button clicked successfully")
            time.sleep(3)  # Reduced from 4 to 3 seconds

            self.log("Form submitted successfully!", "SUCCESS")
            
            # Increment counter for browser restart
            self.forms_processed_since_restart += 1

            return {
                "email": student_data['email'],
                "firstName": student_data['firstName'],
                "lastName": student_data['lastName'],
                "certificationName": student_data.get('certificationName', ''),
                "status": "Success",
                "message": "Completed successfully"
            }

        except Exception as e:
            # Re-raise exception to be handled by retry logic
            raise

    def close_driver(self):
        """Close browser"""
        try:
            if self.driver:
                self.driver.quit()
                self.log("Browser closed")
        except:
            pass

    def save_checkpoint(self, processed_count, total_count, save_results=True):
        """Save progress checkpoint to resume later"""
        try:
            checkpoint_data = {
                "processed_count": processed_count,
                "total_count": total_count,
                "last_update": datetime.now().isoformat(),
                "results_count": len(self.results)
            }
            with open(self.checkpoint_file, 'w', encoding='utf-8') as f:
                json.dump(checkpoint_data, f, indent=2)
            
            # Also save results incrementally (only if not in parallel mode)
            if save_results:
                self.save_results_incremental()
            
        except Exception as e:
            self.log(f"Error saving checkpoint: {str(e)}", "WARNING")
    
    def load_checkpoint(self):
        """Load progress checkpoint if exists"""
        try:
            if os.path.exists(self.checkpoint_file):
                with open(self.checkpoint_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.log(f"Found checkpoint: {data['processed_count']}/{data['total_count']} processed")
                return data.get('processed_count', 0)
        except Exception as e:
            self.log(f"Error loading checkpoint: {str(e)}", "WARNING")
        return 0
    
    def save_results_incremental(self):
        """Save results incrementally to CSV file (append mode)"""
        try:
            file_exists = os.path.exists(self.results_file)
            with open(self.results_file, 'a', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(
                    f, fieldnames=["email", "firstName", "lastName", "certificationName", "status", "message"])
                if not file_exists:
                    writer.writeheader()
                # Write only new results
                if self.results:
                    writer.writerows(self.results)
                    self.results = []  # Clear after saving to save memory
        except Exception as e:
            self.log(f"Error saving results: {str(e)}", "WARNING")
    
    def save_results(self):
        """Save all results to CSV file (final save)"""
        try:
            if not self.results:
                self.log(f"No results to save", "WARNING")
                return
            
            # Write all results at once (overwrite mode to avoid duplicates)
            with open(self.results_file, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(
                    f, fieldnames=["email", "firstName", "lastName", "certificationName", "status", "message"])
                writer.writeheader()
                writer.writerows(self.results)
            
            # Also create a timestamped copy
            filename = f"SAS_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            import shutil
            shutil.copy(self.results_file, filename)
            self.log(f"Results saved to: {filename}")
        except Exception as e:
            self.log(f"Error saving results: {str(e)}", "ERROR")
    
    def restart_browser_if_needed(self):
        """Restart browser periodically to prevent crashes with large datasets"""
        if self.forms_processed_since_restart >= self.restart_browser_interval:
            self.log(f"Restarting browser after {self.forms_processed_since_restart} forms to prevent crashes...", "INFO")
            self.close_driver()
            time.sleep(2)
            self.setup_driver()
            self.forms_processed_since_restart = 0
            self.log("Browser restarted successfully", "SUCCESS")

    def process_students_parallel(self, students, num_workers=3, log_callback=None, result_callback=None, stop_flag=None):
        """
        Process students in parallel using multiple browsers
        num_workers: Number of parallel browsers (range: 2-10, recommended: 3-4)
        log_callback: Function to call for logging
        result_callback: Function to call for each result
        stop_flag: Thread-safe flag to stop processing
        """
        if log_callback:
            log_callback(f"🚀 Starting parallel processing with {num_workers} workers")
        
        results = []
        lock = threading.Lock()
        processed_count = 0
        total_count = len(students)
        
        def process_single_student(student_data):
            """Process a single student in a worker thread"""
            nonlocal processed_count
            
            # Create a new automator instance for this worker
            worker_automator = SASFormAutomator(
                self.form_url,
                self.excel_file,
                browser_choice=self.browser_choice,
                checkpoint_dir=self.checkpoint_dir,
                restart_browser_interval=200  # Higher for parallel processing
            )
            
            try:
                result = worker_automator.fill_form(student_data)
                
                with lock:
                    processed_count += 1
                    results.append(result)
                    
                    if log_callback:
                        status_emoji = "✅" if result['status'] == "Success" else "❌"
                        log_callback(
                            f"[{processed_count}/{total_count}] {status_emoji} {result['status']}: {result['email']}"
                        )
                    
                    if result_callback:
                        result_callback(result)
                    
                    # Save checkpoint every 50 students (without saving results to avoid duplicates)
                    if processed_count % 50 == 0:
                        self.save_checkpoint(processed_count, total_count, save_results=False)
                        if log_callback:
                            log_callback(f"💾 Checkpoint saved: {processed_count}/{total_count}")
                
                return result
                
            except Exception as e:
                error_result = {
                    "email": student_data['email'],
                    "firstName": student_data['firstName'],
                    "lastName": student_data['lastName'],
                    "certificationName": student_data.get('certificationName', ''),
                    "status": "Failed",
                    "message": f"Worker error: {str(e)}"
                }
                with lock:
                    processed_count += 1
                    results.append(error_result)
                return error_result
            finally:
                worker_automator.close_driver()
        
        # Process students in parallel
        with ThreadPoolExecutor(max_workers=num_workers) as executor:
            # Submit all tasks
            future_to_student = {
                executor.submit(process_single_student, student): student 
                for student in students
            }
            
            # Process completed tasks
            for future in as_completed(future_to_student):
                if stop_flag and stop_flag.is_set():
                    if log_callback:
                        log_callback("⏸️ Stopping parallel processing...")
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                
                try:
                    future.result()
                except Exception as e:
                    if log_callback:
                        log_callback(f"❌ Task error: {str(e)}", "ERROR")
        
        # Final save - clear any existing results file first to avoid duplicates
        if os.path.exists(self.results_file):
            os.remove(self.results_file)
        
        self.results = results
        self.save_checkpoint(processed_count, total_count, save_results=False)
        self.save_results()
        
        if log_callback:
            log_callback(f"🎉 Parallel processing completed: {processed_count}/{total_count}")
        
        return results

    def __del__(self):
        """Cleanup resources when object is deleted"""
        self.close_driver()
