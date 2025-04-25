from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# **1. Setup Edge WebDriver**
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-gpu")  # Fix for rendering issues
options.add_argument("--log-level=3")  # Suppress logs
driver = webdriver.Edge(options=options)

# **2. Open Website & Login**
driver.get("https://softwaredevelopmentsolution.com/Contacts/CustomerVendor")

try:
    username_field = driver.find_element(By.ID, "Email")
    password_field = driver.find_element(By.ID, "Password")
    login_button = driver.find_element(By.ID, "LoginSubmit")

    username_field.send_keys("ola123@yopmail.com")
    password_field.send_keys("ola@1234")
    login_button.click()

    # Wait for dashboard to load
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mobile-detailed-view")))
    print("✅ Login successful!")
except:
    print("⚠️ Login not required or already logged in.")

# **3. Wait for Preloader to Disappear**
wait = WebDriverWait(driver, 15)
try:
    wait.until(EC.invisibility_of_element((By.CLASS_NAME, "preloader-it")))
    print("✅ Preloader disappeared.")
except:
    print("⚠️ Preloader took too long to disappear.")

# **4. Click "Import" Button**
import_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Import')]")))
try:
    import_button.click()
except:
    driver.execute_script("arguments[0].click();", import_button)
print("✅ Import button clicked.")

# **5. Click "Import" Option in Dropdown**
import_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Import')]")))
import_option.click()
print("✅ Import option selected.")

# **6. Wait for Import Popup**
wait.until(EC.visibility_of_element_located((By.ID, "ImportPopUp")))
print("✅ Import popup displayed.")

# **7. Upload Correct CSV File**
file_input = driver.find_element(By.ID, "importFile")
file_path = r"C:\Balvant\Ausuma\Ausuma_S/CustomerVendor (19).xlsm"  # Use a correctly formatted CSV file
file_input.send_keys(file_path)
print(f"✅ File uploaded: {file_path}")

# **8. Click "Import" Button**
import_submit = driver.find_element(By.ID, "importData")
import_submit.click()
print("✅ Import started.")

# **9. Wait & Check for Validation Errors**
time.sleep(5)
try:
    validation_popup = driver.find_element(By.ID, "ValidationPopUp")
    if validation_popup.is_displayed():
        print("❌ ERROR: There are invalid records! Check the validation table.")
except:
    print("✅ No validation errors. Import successful!")

# **10. Close Browser**
time.sleep(3)
driver.quit()
print("✅ Browser closed.")
