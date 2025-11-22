import re
import time
import threading
import os
import requests
from openpyxl import Workbook

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

def check_link(url):
    try:
        response = requests.get(url, timeout=10)
        return response.status_code, "Success" if response.status_code == 200 else "Failed"
    except requests.exceptions.RequestException as e:
        return None, str(e)

def run_check(activity_url, check_id, report_data, dashboard_data):
    # Create a new headless Chrome driver for each check
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(activity_url)

    wait = WebDriverWait(driver, 10)

    try:
        textarea = wait.until(EC.presence_of_element_located((By.ID, "detail-text")))
        text = textarea.get_attribute("value")  # <-- use 'value' for textarea

        # Extract first URL
        urls = re.findall(r'https?://[^\s"]+', text)
        if not urls:
            raise RuntimeError("No URL found in the textarea text.")
        target_url = urls[0]
        print(f"Check {check_id}: Extracted URL:", target_url)

        # Check link status
        status_code, reason = check_link(target_url)
        report_data.append([target_url, status_code, "Checked" if status_code == 200 else "Not Checked", reason])

        # Open URL in a new tab
        driver.execute_script("window.open(arguments[0], '_blank');", target_url)

        # Switch to the newly opened tab
        driver.switch_to.window(driver.window_handles[-1])

        # Wait for the page to load
        wait.until(lambda d: d.title is not None)

        # Take a screenshot with unique name
        screenshot_path = f"screenshots/screenshot_{check_id}.png"
        os.makedirs(os.path.dirname(screenshot_path), exist_ok=True)
        driver.save_screenshot(screenshot_path)

        # Switch back to the activity tab
        driver.switch_to.window(driver.window_handles[0])

        # Re-locate and interact with elements to avoid stale element reference
        # Fix stale element reference by retrying with a small wait loop
        for _ in range(3):
            try:
                screenshot_element = wait.until(EC.presence_of_element_located((By.ID, "screenshot")))
                screenshot_element.send_keys(os.path.abspath(screenshot_path))
                break
            except Exception as e:
                time.sleep(0.5)
        else:
            raise RuntimeError("Failed to locate screenshot element after retries")

        # Select green radio button
        for _ in range(3):
            try:
                green_radio = wait.until(EC.element_to_be_clickable((By.ID, "green")))
                driver.execute_script("arguments[0].click();", green_radio)
                break
            except Exception as e:
                time.sleep(0.5)
        else:
            raise RuntimeError("Failed to click green radio button after retries")

        # Enter name
        for _ in range(3):
            try:
                name_field = wait.until(EC.presence_of_element_located((By.ID, "name")))
                name_field.send_keys("PyBot")
                break
            except Exception as e:
                time.sleep(0.5)

        # Enable submit button and submit
        driver.execute_script("document.querySelector('.submit-btn').classList.add('enabled');")
        submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'submit-btn')]")))
        driver.execute_script("arguments[0].click();", submit_btn)

        # Collect data for dashboard
        dashboard_data.append({
            'id': check_id,
            'siteName': target_url,
            'responseCode': status_code,
            'status': "Checked" if status_code == 200 else "Not Checked",
            'reason': reason,
            'radioChoice': 'Green'
        })

        print(f"Check {check_id} completed successfully.")

    except Exception as e:
        print(f"Error in check {check_id}: {e}")
        report_data.append([target_url if 'target_url' in locals() else "N/A", None, "Error", str(e)])
        dashboard_data.append({
            'id': check_id,
            'siteName': target_url if 'target_url' in locals() else "N/A",
            'responseCode': None,
            'status': "Error",
            'reason': str(e),
            'radioChoice': 'N/A'
        })
    finally:
        driver.quit()

def automate_login(driver, wait):
    try:
        # Wait for username field
        username_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
        password_field = wait.until(EC.presence_of_element_located((By.ID, "password")))
        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']")))

        # Fill credentials
        username_field.send_keys("admin")
        password_field.send_keys("password")

        # Submit the form
        login_button.click()

        # Wait for redirect to index.html by checking URL contains 'index.html'
        wait.until(EC.url_contains("index.html"))
        print("Login successful, redirected to index.html")

    except Exception as e:
        print("Login failed or timed out:", e)
        driver.quit()
        raise e

def main():
    # Main driver (not headless to observe the login process)
    main_driver = webdriver.Chrome()
    main_driver.maximize_window()
    wait = WebDriverWait(main_driver, 20)

    # Open index.html as the initial page
    main_driver.get("https://kingnstarpancard-code.github.io/axis_automation/index.html")

    # Check if redirected to login page by checking URL or presence of login form elements
    current_url = main_driver.current_url
    if "login.html" in current_url:
        print("Login page detected, performing login automation...")
        automate_login(main_driver, wait)
        # After login, the driver should already be on index.html or will be redirected
    else:
        print("Already logged in, proceeding with activity checks...")

    # Proceed to automate activity checking as before

    # Get all check button URLs
    check_buttons = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//a[@class='check-btn']")))
    activity_urls = [btn.get_attribute("href") for btn in check_buttons]

    # Shared lists for report data and dashboard data
    report_data = []
    dashboard_data = []

    # Run all checks in parallel using threads (headless)
    threads = []
    for i, url in enumerate(activity_urls, start=1):
        t = threading.Thread(target=run_check, args=(url, i, report_data, dashboard_data))
        threads.append(t)
        t.start()

    # Wait for all threads to complete
    for t in threads:
        t.join()

    print("All checks completed.")

    # Generate Excel report
    wb = Workbook()
    ws = wb.active
    ws.title = "Link Check Report"
    ws.append(["Site Name", "Response Code", "Status", "Reason"])

    for row in report_data:
        ws.append(row)

    try:
        wb.save("link_check_report.xlsx")
        print("Excel report generated: link_check_report.xlsx")
    except PermissionError:
        print("Permission denied: Excel file is open. Saving as link_check_report_new.xlsx")
        wb.save("link_check_report_new.xlsx")
        print("Excel report generated: link_check_report_new.xlsx")

    # Save dashboard data to localStorage for the dashboard
    main_driver.execute_script("localStorage.setItem('reportData', JSON.stringify(arguments[0]));", dashboard_data)

    # Update localStorage in main driver to mark all as checked
    checked_ids = list(range(1, len(activity_urls) + 1))
    main_driver.execute_script(f"localStorage.setItem('checked', JSON.stringify({checked_ids}));")

    # Refresh the main page to show checked status
    main_driver.refresh()

    print("Main page refreshed. Check if all activities are marked as checked.")
    input("Press Enter to exit...")

    main_driver.quit()

if __name__ == "__main__":
    main()
