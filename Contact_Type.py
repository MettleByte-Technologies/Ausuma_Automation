from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time 

# ✅ Use Microsoft Edge WebDriver
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")  # Open in maximized mode

# ✅ Initialize Edge WebDriver
driver = webdriver.Edge(options=options)  # EdgeDriver is auto-managed by Selenium

try:
    # Step 1: Open the login page
    driver.get("https://softwaredevelopmentsolution.com")  # Replace with your login page URL
    time.sleep(2)

    # Step 2: Enter login credentials
    username_field = driver.find_element(By.ID, "Email")  # Replace with actual username field ID
    password_field = driver.find_element(By.ID, "Password")  # Replace with actual password field ID
    login_button = driver.find_element(By.ID, "LoginSubmit")  # Replace with actual login button ID

    username_field.send_keys("ola123@yopmail.com")  # Replace with your username
    password_field.send_keys("ola@1234")  # Replace with your password
    login_button.click()
    
    # Step 3: Wait until login completes
    WebDriverWait(driver, 10).until(
        EC.url_contains("Dashboard")  # Adjust based on the redirected URL
    )

    # Step 4: Navigate to the target page
    driver.get("https://softwaredevelopmentsolution.com/Contacts/CustVendContactTypes/Create")  # Replace with your target page URL
    time.sleep(2)

    wait = WebDriverWait(driver, 10)

    # Step 5: Fill the contact form
    name_field = wait.until(EC.presence_of_element_located((By.ID, "cUSTVENDCONTACTTYPE_NAME")))  
    name_field.send_keys("S_HRTEST 4421")

    description_field = wait.until(EC.presence_of_element_located((By.ID, "cUSTVENDCONTACTTYPE_DESCRIPTION")))  
    description_field.send_keys("S_HR_DESCRIPTIONTEST 2441")

    # Step 6: Click the Submit button
    submit_button = driver.find_element(By.XPATH, "//input[@type='button' and @value='Submit']")
    submit_button.click()

    
    
    print("Test completed successfully.")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # Close the browser after a delay (for visibility)
    time.sleep(5)
    driver.quit()
