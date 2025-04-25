from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Set up WebDriver
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Edge(options=options)

try:
    driver.get("https://softwaredevelopmentsolution.com")
    driver.maximize_window()

    # Wait until the username and password fields are visible
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "Email"))
    )
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "Password"))
    )

    # Locate the fields and enter credentials
    username_field = driver.find_element(By.ID, "Email")
    password_field = driver.find_element(By.ID, "Password")
    username_field.send_keys("velonova@yopmail.com")
    password_field.send_keys("Ausuma@123")

    # Wait for the login button and click it
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "LoginSubmit"))
    )
    login_button = driver.find_element(By.ID, "LoginSubmit")
    login_button.click()

    # Wait for the page to load
    WebDriverWait(driver, 10).until(
        EC.url_contains("dashboard")
    )

    # Validate successful login
    if "dashboard" in driver.current_url:
        print("Login successful!")
    else:
        print("Login failed!")

finally:
    driver.quit()
    print("Browser closed.")