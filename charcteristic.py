from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time

# Initialize Microsoft Edge WebDriver
driver = webdriver.Edge()
driver.maximize_window()

try:
    # Open the login page
    driver.get("https://softwaredevelopmentsolution.com")

    # Wait and enter login credentials
    wait = WebDriverWait(driver, 30)
    username_field = wait.until(EC.visibility_of_element_located((By.ID, "Email")))
    password_field = driver.find_element(By.ID, "Password")
    login_button = driver.find_element(By.ID, "LoginSubmit")

    username_field.send_keys("ola123@yopmail.com")
    password_field.send_keys("ola@1234")
    login_button.click()
    print("✅ Successfully logged in!")

    # Navigate to form page
    driver.get("https://softwaredevelopmentsolution.com/Inventory/Characteristics/Create")

    # Switch to iframe if present
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    if iframes:
        print("🔁 Switching to iframe...")
        driver.switch_to.frame(iframes[0])

    # Fill Name
    name_field = wait.until(EC.element_to_be_clickable((By.ID, "cHARACTERISTIC_NAME")))
    name_field.send_keys("WWERTWWW")
    print("✅ Name field filled!")

    # Fill Description
    description_field = wait.until(EC.element_to_be_clickable((By.ID, "cHARACTERISTIC_DESCRIPTION")))
    description_field.send_keys("WWWERTYUW")
    print("✅ Description field filled!")

    # Handle dropdowns
    def select_div_dropdown(dropdown_id, option_text):
        try:
            print(f"🔽 Selecting '{option_text}' from dropdown...")
            dropdown = wait.until(EC.element_to_be_clickable((By.ID, dropdown_id)))
            dropdown.click()
            time.sleep(2)
            option_xpath = f"//div[@id='{dropdown_id}']//li[contains(text(), '{option_text}')]"
            option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
            option.click()
            time.sleep(1)
            print(f"✅ Selected '{option_text}' successfully!")
            driver.find_element(By.TAG_NAME, "body").click()
            time.sleep(1)
        except Exception as e:
            print(f"❌ ERROR selecting {option_text}: {e}")

    select_div_dropdown("cHARACTERISTIC_VALUETYPE_chosen", "Numeric")
    select_div_dropdown("cHARACTERISTIC_VALUERANGE_chosen", "Close ended")
    select_div_dropdown("cHARACTERISTIC_ISUNITREQUIRED_chosen", "No")

    # Submit button logic
    try:
        print("🔍 Looking for Submit button by class...")
        submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-success') and text()='Submit']")))

        driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
        time.sleep(2)

        try:
            ActionChains(driver).move_to_element(submit_button).click().perform()
            print("✅ Submit button clicked using ActionChains!")
        except Exception as e1:
            print(f"⚠️ ActionChains failed: {e1}")
            try:
                submit_button.click()
                print("✅ Submit button clicked using .click()!")
            except Exception as e2:
                print(f"⚠️ .click() failed: {e2}")
                try:
                    driver.execute_script("arguments[0].click();", submit_button)
                    print("✅ Submit button clicked using JavaScript!")
                except Exception as e3:
                    print(f"❌ All click methods failed: {e3}")
                    driver.save_screenshot("submit_error.png")
                    print("📸 Screenshot saved as submit_error.png")

    except Exception as e:
        print(f"❌ ERROR locating Submit button: {e}")
        driver.save_screenshot("submit_error.png")
        print("📸 Screenshot saved as submit_error.png")

except Exception as e:
    print(f"❌ Main ERROR: {e}")

finally:
    time.sleep(5)
    driver.quit()
