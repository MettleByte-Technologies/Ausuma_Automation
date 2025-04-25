from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Initialize Microsoft Edge WebDriver
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Edge(options=options)

# Open the signup page
driver.get("https://staging.ausumaerp.com/account/signup")

# WebDriverWait instance
wait = WebDriverWait(driver, 20)

def safe_find_element(by, value):
    """Safely locate an element with error handling."""
    try:
        return wait.until(EC.presence_of_element_located((by, value)))
    except Exception as e:
        print(f"❌ Error locating element {value}: {e}")
        return None

def safe_click(element):
    """Safely click an element using JavaScript."""
    try:
        driver.execute_script("arguments[0].scrollIntoView();", element)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", element)
    except Exception as e:
        print(f"❌ Error clicking element: {e}")

def set_value_with_js(element, value):
    """Set a value for an input field using JavaScript."""
    try:
        driver.execute_script("arguments[0].value = arguments[1];", element, value)
    except Exception as e:
        print(f"❌ Error setting value for element: {e}")

# Field values
fields = {
    "company_name": "Test Automation Script",
    "company_description": "This is a test description.",
    "address1": "123 Test Street",
    "address2": "Suite 100",
    "city": "Test City",
    "zipcode": "12345",
    "email": "testautomationscript@yopmai.com",
    "first_name": "Testnew",
    "last_name": "AutoS",
    "middle_name": "S",
    "employee_number": "EMP123000",
    "ssn": "000-00-0000",
    "phone": "123-456-7825",
    "mobile": "987-654-3270"
}

# Corrected element mappings
element_mapping = {
    "company_name": (By.ID, "COMPANY_NAME"),
    "company_description": (By.ID, "COMPANY_DESCRIPTION"),
    "financial_year": (By.ID, "ddlMonthName"),
    "address1": (By.ID, "CONTACT_INFORMATION_ADDRESS1"),
    "address2": (By.ID, "CONTACT_INFORMATION_ADDRESS2"),
    "city": (By.ID, "CONTACT_INFORMATION_CITY"),
    "zipcode": (By.ID, "CONTACT_INFORMATION_ZIPCODE"),
    "country": (By.ID, "ddlCompanyCountry"),
    "state": (By.ID, "ddlCompanyState"),
    "state_name": (By.ID, "CONTACT_INFORMATION_STATENAME"),
    "phone": (By.ID, "CONTACT_INFORMATION_PHONE1"),
    "mobile": (By.ID, "CONTACT_INFORMATION_PHONE2"),
    "email": (By.ID, "CONTACT_INFORMATION_EMAIL"),
    "first_name": (By.ID, "eMPLOYEE_FIRSTNAME"),
    "last_name": (By.ID, "eMPLOYEE_LASTNAME"),
    "middle_name": (By.ID, "eMPLOYEE_MIDDLEINITIAL"),
    "employee_number": (By.ID, "eMPLOYEE_NUMBER"),
    "ssn": (By.ID, "eMPLOYEE_SSN"),
    "copy_company_address": (By.ID, "copyAddress"),
    "accept_terms": (By.ID, "term-condition"),
    "submit_button": (By.XPATH, "//input[@type='submit']")
}

# Fill text fields
for field, value in fields.items():
    element = safe_find_element(*element_mapping[field])
    if element:
        set_value_with_js(element, value)

# Select financial year dropdown
try:
    financial_year_dropdown = Select(safe_find_element(*element_mapping["financial_year"]))
    financial_year_dropdown.select_by_visible_text("January")
    print("✅ Financial Year selected: January")
except Exception as e:
    print(f"❌ Error selecting financial year: {e}")

# Select country first (to load states dynamically)
try:
    country_dropdown = Select(safe_find_element(*element_mapping["country"]))
    country_dropdown.select_by_visible_text("USA")
    print("✅ Country selected: USA")
    time.sleep(3)  # Allow state dropdown to update
except Exception as e:
    print(f"❌ Error selecting country: {e}")

# Select state dropdown with fallback mechanism
   # Step 1: Check if dropdown exists
try:
    state_dropdown = wait.until(EC.presence_of_element_located((By.ID, "ddlCompanyState")))
    
    if not state_dropdown.is_enabled():
        driver.execute_script("arguments[0].removeAttribute('disabled');", state_dropdown)

    driver.execute_script("arguments[0].scrollIntoView(true);", state_dropdown)

    try:
        state_dropdown.click()
        option = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[@value='1']")))
        option.click()
    except:
        driver.execute_script("document.getElementById('ddlCompanyState').value = '1';")
        driver.execute_script("document.getElementById('ddlCompanyState').dispatchEvent(new Event('change'));")

except Exception as e:
    print(f"❌ Error: {e}") 

# Click checkboxes
for checkbox in ["copy_company_address", "accept_terms"]:
    element = safe_find_element(*element_mapping[checkbox])
    if element:
        safe_click(element)

# Handle Terms & Conditions popup (if it appears)
try:
    popup = wait.until(EC.visibility_of_element_located((By.ID, "termsAndConditionsModal")))
    close_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "close")))
    close_button.click()
    wait.until(EC.invisibility_of_element((By.ID, "termsAndConditionsModal")))
    print("✅ Terms & Conditions popup handled.")
except Exception:
    print("ℹ️ No Terms & Conditions popup detected.")

# Submit form
submit_button = safe_find_element(*element_mapping["submit_button"])
if submit_button:
    safe_click(submit_button)
    print("✅ Form submitted successfully!")

time.sleep(5)  # Allow time for form submission
driver.quit()