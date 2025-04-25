from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Initialize the WebDriver (Make sure to set the path of the WebDriver if needed)
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Edge(options=options)

try:
    # Step 1: Open the login page
    driver.get("https://softwaredevelopmentsolution.com")  # Replace with your login page URL
    
    # Step 2: Enter login credentials
    username_field = driver.find_element(By.ID, "Email")  # Replace with actual username field ID
    password_field = driver.find_element(By.ID, "Password")  # Replace with actual password field ID
    login_button = driver.find_element(By.ID, "LoginSubmit")  # Replace with actual login button ID
    
    username_field.send_keys("ola123@yopmail.com")  # Replace with your username
    password_field.send_keys("ola@1234")  # Replace with your password
    login_button.click()
    
    # Step 3: Wait until redirection happens (adjust the element for the redirected page)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "mobile-detailed-view"))
    )
    
    # Step 4: Navigate to the target page (if required)
    driver.get("https://softwaredevelopmentsolution.com/Access/UserSetup/Create")  # Replace with your target page URL
    
    # Step 5: Interact with text fields
    dropdown = Select(driver.find_element(By.ID, "idEmployee"))  # type: ignore # Replace with actual dropdown ID
    dropdown.select_by_visible_text("123 456")  # Replace with the required option
    
    # Step 6: Select a value from a dropdown
    dropdown = Select(driver.find_element(By.ID, "idRole"))  # type: ignore # Replace with actual dropdown ID
    dropdown.select_by_visible_text("New admin")  # Replace with the required option
    
    # Step 7: Submit the form
    submit_button = driver.find_element(By.VALUE, "Submit")  # Replace with actual submit button ID
    submit_button.click()
    
    # Step 8: Wait for confirmation (if applicable)
   # WebDriverWait(driver, 10).until(
     #   EC.presence_of_element_located((By.ID, "CONFIRMATION_MESSAGE_ID"))  # Replace with confirmation message ID
  #  )
    
    print("Test completed successfully.")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # Close the browser after a delay (for visibility)
    time.sleep(5)
    driver.quit()