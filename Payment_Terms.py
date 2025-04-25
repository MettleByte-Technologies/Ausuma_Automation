from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Initialize Edge browser
driver = webdriver.Edge()  # Make sure msedgedriver is in PATH
driver.maximize_window()

# Go to login page
driver.get("https://softwaredevelopmentsolution.com/Accounting/PaymentTerm/Create")

wait = WebDriverWait(driver, 20)

# Step 1: Login if required
try:
    username_field = wait.until(EC.presence_of_element_located((By.ID, "Email")))
    password_field = wait.until(EC.presence_of_element_located((By.ID, "Password")))
    login_button = wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit")))

    username_field.send_keys("ola123@yopmail.com")
    password_field.send_keys("1")
    login_button.click()

    wait.until(EC.presence_of_element_located((By.ID, "mobile-detailed-view")))
    print("✅ Login successful!")
except Exception as e:
    print(f"⚠️ Login not required or already logged in: {e}")

# Step 2: Fill out the Payment Term form
try:
    # Name
    wait.until(EC.presence_of_element_located((By.ID, "pAYMENTTERM_NAME"))).send_keys("Test Payment Term")
    print("✅ Entered Name")

    # Description
    wait.until(EC.presence_of_element_located((By.ID, "pAYMENTTERM_DESCRIPTION"))).send_keys("This is a test description.")
    print("✅ Entered Description")

    # Payment No. of Days
    wait.until(EC.presence_of_element_located((By.ID, "pAYMENTTERM_PAYMENTNOOFDAYS"))).send_keys("30")
    print("✅ Entered Payment No. of Days")
except Exception as e:
    print(f"❌ Form input failed: {e}")

# Step 3: Select from Chosen Surcharge Dropdown
try:
    surcharge_container = wait.until(EC.presence_of_element_located((By.ID, "pAYMENTTERM_PAYMENTSURCHARGEID_chosen")))
    chosen_link = surcharge_container.find_element(By.CLASS_NAME, "chosen-single")
    driver.execute_script("arguments[0].click();", chosen_link)

    # Input search term (e.g., "Account Payment")
    search_box = surcharge_container.find_element(By.XPATH, ".//input")
    search_box.send_keys("Account Payment")
    time.sleep(1.2)

    # Click the matched option
    try:
        option = surcharge_container.find_element(By.XPATH, ".//li[contains(text(), 'Account Payment')]")
        driver.execute_script("arguments[0].click();", option)
        print("✅ Selected Surcharge from dropdown")
    except:
        search_box.send_keys(Keys.ENTER)
        print("✅ Selected Surcharge via ENTER (fallback)")
except Exception as e:
    print(f"❌ Failed to select surcharge: {e}")

# Step 4: Click Submit
try:
    submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Submit')]")))
    driver.execute_script("arguments[0].click();", submit_button)
    print("✅ Form submitted")
except Exception as e:
    print(f"❌ Submit failed: {e}")

# Optional: Close browser
# driver.quit()
