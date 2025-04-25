from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

driver = webdriver.Chrome()
driver.get("http://your-url.com")  # replace with actual URL
wait = WebDriverWait(driver, 10)

def sleep_and_log(field_name):
    time.sleep(3)
    print(f"[✓] {field_name} - Passed")

def log_fail(field_name, error):
    print(f"[✗] {field_name} - Failed: {str(error)}")

# 1. Click "Add Location Quantity" Button
try:
    add_button = wait.until(EC.element_to_be_clickable((By.ID, "addLocationBtn")))
    add_button.click()
    sleep_and_log("Add Location Quantity Button")
except Exception as e:
    log_fail("Add Location Quantity Button", e)

# 2. Select Location dropdown
try:
    location_dropdown = wait.until(EC.presence_of_element_located((By.NAME, "location")))
    Select(location_dropdown).select_by_visible_text("Main Warehouse")  # replace with your option
    sleep_and_log("Location Dropdown")
except Exception as e:
    log_fail("Location Dropdown", e)

# 3. Select Location Bin dropdown
try:
    bin_dropdown = wait.until(EC.presence_of_element_located((By.NAME, "locationBin")))
    Select(bin_dropdown).select_by_visible_text("Bin A")  # replace with your option
    sleep_and_log("Location Bin Dropdown")
except Exception as e:
    log_fail("Location Bin Dropdown", e)

# 4. Enter Batch No
try:
    batch_input = wait.until(EC.presence_of_element_located((By.NAME, "batchNo")))
    batch_input.send_keys("B12345")
    sleep_and_log("Batch No")
except Exception as e:
    log_fail("Batch No", e)

# 5. Enter Opening Quantity
try:
    qty_input = wait.until(EC.presence_of_element_located((By.NAME, "openingQty")))
    qty_input.clear()
    qty_input.send_keys("10")
    sleep_and_log("Opening Quantity")
except Exception as e:
    log_fail("Opening Quantity", e)

# 6. Enter Total Cost
try:
    cost_input = wait.until(EC.presence_of_element_located((By.NAME, "totalCost")))
    cost_input.clear()
    cost_input.send_keys("250")
    sleep_and_log("Total Cost")
except Exception as e:
    log_fail("Total Cost", e)

# 7. Click Serial Button and Enter Serial Number in Popup
try:
    serial_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "serial-btn")))
    serial_btn.click()
    sleep_and_log("Serial Button Click")

    serial_input = wait.until(EC.visibility_of_element_located((By.ID, "serialNo")))
    serial_input.send_keys("SN-987654321")
    sleep_and_log("Serial Number Input")

    save_serial_btn = wait.until(EC.element_to_be_clickable((By.ID, "saveSerial")))
    save_serial_btn.click()
    sleep_and_log("Save Serial Number")
except Exception as e:
    log_fail("Serial Number Flow", e)

# Optional: Close browser
time.sleep(3)
driver.quit()
