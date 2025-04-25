from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time

# Initialize the WebDriver (Make sure to set the path of the WebDriver if needed)
driver = webdriver.Chrome()

try:
    # Step 1: Open the login page
    driver.get("https://softwaredevelopmentsolution.com")  # Replace with your login page URL
    
    # Step 2: Enter login credentials
    username_field = driver.find_element(By.ID, "USERNAME_FIELD_ID")  # Replace with actual username field ID
    password_field = driver.find_element(By.ID, "PASSWORD_FIELD_ID")  # Replace with actual password field ID
    login_button = driver.find_element(By.ID, "LOGIN_BUTTON_ID")  # Replace with actual login button ID
    
    username_field.send_keys("YOUR_USERNAME")  # Replace with your username
    password_field.send_keys("YOUR_PASSWORD")  # Replace with your password
    login_button.click()
    
    # Step 3: Wait until redirection happens (adjust the element for the redirected page)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "UNIQUE_ELEMENT_ON_NEXT_PAGE_ID"))
    )
    
    # Step 4: Navigate to the first target page
    driver.get("YOUR_FIRST_TARGET_PAGE_URL")  # Replace with your first target page URL
    
    # Step 5: Select values from two dropdowns
    dropdown1 = Select(driver.find_element(By.CLASS_NAME, "DROPDOWN_1_CLASS"))  # Replace with actual dropdown class
    dropdown1.select_by_visible_text("Option 1")  # Replace with the required option
    
    dropdown2 = Select(driver.find_element(By.CLASS_NAME, "DROPDOWN_2_CLASS"))  # Replace with actual dropdown class
    dropdown2.select_by_visible_text("Option 2")  # Replace with the required option
    
    # Step 6: Navigate to the second target page
    driver.get("YOUR_SECOND_TARGET_PAGE_URL")  # Replace with your second target page URL
    
    # Step 7: Submit the form
    submit_button = driver.find_element(By.ID, "SUBMIT_BUTTON_ID")  # Replace with actual submit button ID
    submit_button.click()
    
    # Step 8: Wait for confirmation (if applicable)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "CONFIRMATION_MESSAGE_ID"))  # Replace with confirmation message ID
    )
    
    print("Test completed successfully.")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # Close the browser after a delay (for visibility)
    time.sleep(5)
    driver.quit()