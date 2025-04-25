from selenium import webdriver
from selenium.webdriver.common.by import By
#from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
import time
options = webdriver.EdgeOptions()# Setup
options.add_argument("--start-maximized")
driver = webdriver.Edge(options=options)
driver.get("https://softwaredevelopmentsolution.com/Inventory/Products/Create")  # Change this to your actual URL
try:
    # Wait for login fields to appear
    wait = WebDriverWait(driver, 10)
    username_field = wait.until(EC.presence_of_element_located((By.ID, "Email")))
    password_field = wait.until(EC.presence_of_element_located((By.ID, "Password")))
    login_button = wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit")))
    username_field.send_keys("ola123@yopmail.com")
    password_field.send_keys("1")
    login_button.click()
    # Wait for dashboard to load
    wait.until(EC.presence_of_element_located((By.ID, "mobile-detailed-view")))
    print("✅ Login successful!")
except Exception as e:
    print(f"⚠️ Login not required or already logged in: {e}")
wait = WebDriverWait(driver, 20)
# ========== Step 1: Add Name ==========
try:
    wait.until(EC.presence_of_element_located((By.ID, "pRODUCT_NAME"))).send_keys("ABC TEST  28-03-2025")
    print("✔️ Name added successfully")
except Exception as e:
    print(f"❌ Failed at Step 1 (Name): {e}")
# ========== Step 2: Select Category ==========
try:
    # Step 1: Open the dropdown safely
    category_dropdown = wait.until(EC.presence_of_element_located((By.ID, "ddlCategory")))
    # Scroll slightly to avoid navbar overlap
    driver.execute_script("window.scrollBy(0, 150);")  # Scroll down 150px
    time.sleep(1)
    # Click using JS to avoid click interception
    driver.execute_script("arguments[0].click();", category_dropdown)
    print("✔️ Dropdown opened using JS click")
    # Step 2: Select Bike Tires
    bike_tires_xpath = "//a[@href='#1139']"
    bike_tires_link = wait.until(EC.visibility_of_element_located((By.XPATH, bike_tires_xpath)))
    # Scroll to it
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", bike_tires_link)
    time.sleep(1)
    # Click it
    driver.execute_script("arguments[0].click();", bike_tires_link)
    print("✔️ Bike Tires category selected")
except Exception as e:
    print(f"❌ Error: {e}")
# ========== Step 3: Add Description ==========
try:
    wait.until(EC.presence_of_element_located((By.ID, "pRODUCT_DESCRIPTION"))).send_keys("This is a test description.")
    print("✔️ Description added successfully")
except Exception as e:
    print(f"❌ Failed at Step 3 (Description): {e}")
# ========== Step 6: Minimum Quantity ==========
try:
    wait.until(EC.presence_of_element_located((By.ID, "pRODUCT_MINIMUMQUANTITY"))).send_keys("1000")
    print("✔️ Minimum Quantity '1000' added successfully")
except Exception as e:
    print(f"❌ Failed at Step 6 (Minimum Quantity): {e}")
# ========== Step 7: Reorder Quantity ==========
try:
    wait.until(EC.presence_of_element_located((By.ID, "pRODUCT_REORDERQUANTITY"))).send_keys("1000")
    print("✔️ Reorder Quantity '1000' added successfully")
except Exception as e:
    print(f"❌ Failed at Step 7 (Reorder Quantity): {e}")
# ========== Step 14: Add Location ==========
try:
    print("🔄 Clicking 'Add Location Quantity' button...")
    add_location_button = wait.until(EC.element_to_be_clickable((By.ID, "productFormBtnAddStock")))
    driver.execute_script("arguments[0].click();", add_location_button)
    print("✅ Clicked 'Add Location Quantity' button")
except Exception as e:
    print(f"❌ Failed at Step 1 (Add Location): {e}")
# STEP 2: Select Location
try:
    print("🔄 Selecting Location...")
    location_container = wait.until(EC.presence_of_element_located((By.ID, "ddlWearhouse_Grd_chosen")))
    chosen_link = location_container.find_element(By.CLASS_NAME, "chosen-single")
    driver.execute_script("arguments[0].click();", chosen_link)

    chosen_input = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='ddlWearhouse_Grd_chosen']//input")))
    chosen_input.clear()
    chosen_input.send_keys("OLA Electric Showroom")
    time.sleep(1.5)
    try:
        # Try clicking from list
        location_option = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='ddlWearhouse_Grd_chosen']//li[contains(text(), 'OLA Electric Showroom')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", location_option)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", location_option)
        print("✅ Selected Location: OLA Electric Showroom")
    except:
        # Fallback: Press Enter
        print("⚠️ Option not clickable, sending ENTER as fallback.")
        chosen_input.send_keys(Keys.ENTER)
        print("✅ Selected by ENTER key")
except Exception as e:
    print(f"❌ Failed at Step 2 (Location dropdown): {e}")
# STEP 3: Select Bin
try:
    print("🔄 Selecting Location Bin...")
    bin_container = wait.until(EC.presence_of_element_located((By.ID, "ddlLocation_Grd_chosen")))
    bin_link = bin_container.find_element(By.CLASS_NAME, "chosen-single")
    driver.execute_script("arguments[0].click();", bin_link)
    bin_input = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='ddlLocation_Grd_chosen']//input")))
    bin_input.clear()
    bin_input.send_keys("EV Warehouse 001")
    time.sleep(1.5)
    try:
        bin_option = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='ddlLocation_Grd_chosen']//li[contains(text(), 'EV Warehouse 001')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", bin_option)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", bin_option)
        print("✅ Selected Bin: EV Warehouse 001")
    except:
        print("⚠️ Bin option not clickable, sending ENTER as fallback.")
        bin_input.send_keys(Keys.ENTER)
        print("✅ Bin selected by ENTER key")
except Exception as e:
    print(f"❌ Failed at Step 3 (Location Bin dropdown): {e}")
# STEP 4: Quantity
try:
    print("🔄 Entering Quantity...")
    quantity_input = wait.until(EC.element_to_be_clickable((By.ID, "txtLocationQuantity_Grd")))
    quantity_input.clear()
    quantity_input.send_keys("100")
    print("✅ Entered Quantity: 100")
except Exception as e:
    print(f"❌ Failed at Step 4 (Quantity): {e}")
# STEP 5: Cost
try:
    print("🔄 Entering Cost...")
    cost_input = wait.until(EC.element_to_be_clickable((By.ID, "txtCost_Grd")))
    cost_input.clear()
    cost_input.send_keys("1000")
    print("✅ Entered Cost: 1000")
except Exception as e:
    print(f"❌ Failed at Step 5 (Cost): {e}")
# STEP 6: Unit Cost (optional)
try:
    print("🔄 Entering Opening Unit Cost...")
    unit_cost_input = wait.until(EC.presence_of_element_located((By.ID, "pRODUCT_OPENINGUNITCOST")))
    unit_cost_input.clear()
    unit_cost_input.send_keys("10")
    print("✅ Entered Opening Unit Cost: 10")
except Exception as e:
    print(f"⚠️ Skipped Opening Unit Cost: {e}")
# STEP 5 - Enter Cost
try:
    print("🔄 Entering Cost...")
    cost_input = wait.until(EC.element_to_be_clickable((By.ID, "txtCost_Grd")))
    cost_input.clear()
    cost_input.send_keys("1000")
    print("✅ Entered Cost: 1000")
except Exception as e:
    print(f"❌ Failed at Step 5 (Cost): {e}")
# STEP 6 - Enter Opening Unit Cost (Optional)
try:
    print("🔄 Entering Opening Unit Cost...")
    unit_cost_input = wait.until(EC.presence_of_element_located((By.ID, "pRODUCT_OPENINGUNITCOST")))
    unit_cost_input.clear()
    unit_cost_input.send_keys("10")
    print("✅ Entered Opening Unit Cost: 10")
except Exception as e:
    print(f"⚠️ Skipped Opening Unit Cost (optional): {e}")
# ========== Step 16: Add Unit ==========
try:
    wait.until(EC.element_to_be_clickable((By.ID, "productFormBtnAddUnit"))).click()
    print("✔️ Clicked on 'Add Unit' button successfully")

except Exception as e:
    print(f"❌ Failed at Click Unit Button: {e}")
try:
    print("🔄 Selecting Unit...")
    unit_container = wait.until(EC.presence_of_element_located((By.ID, "ddlUnit_Grd_chosen")))
    chosen_link = unit_container.find_element(By.CLASS_NAME, "chosen-single")
    driver.execute_script("arguments[0].click();", chosen_link)

    chosen_input = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='ddlUnit_Grd_chosen']//input")))
    chosen_input.clear()
    chosen_input.send_keys("piece")
    time.sleep(1.5)
    try:
        # Try clicking from list
        unit_option = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='ddlUnit_Grd_chosen']//li[contains(text(), 'piece')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", unit_option)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", unit_option)
        print("✅ Selected Unit: piece")
    except:
        # Fallback: Press Enter
        print("⚠️ Option not clickable, sending ENTER as fallback.")
        chosen_input.send_keys(Keys.ENTER)
        print("✅ Selected Unit by ENTER key")
except Exception as e:
    print(f"❌ Failed at Unit dropdown: {e}")
try:
    # Fill SKU
    wait.until(EC.presence_of_element_located((By.ID, "txtSku_Grd"))).send_keys("10")
    print("✔️ SKU '10' added successfully")
    # Fill UPC
    wait.until(EC.presence_of_element_located((By.ID, "txtUpc_Grd"))).send_keys("10")
    print("✔️ UPC '10' added successfully")
    # Fill Sales Price
    wait.until(EC.presence_of_element_located((By.ID, "txtSalesPrice_Grd"))).send_keys("15")
    print("✔️ Sales price '15' added successfully")
    # Click Save
    wait.until(EC.element_to_be_clickable((By.ID, "Addnewrow"))).click()
    print("✔️ Unit saved successfully")
except Exception as e:
    print(f"❌ Failed at Step 16 (Unit Details): {e}")
# ========== Step 17: Characteristics ==========
### 1. Select Unit: "Electric Motor"
try:
    print("🔄 Selecting Unit: Electric Motor...")
    unit_container = wait.until(EC.presence_of_element_located((By.ID, "ddlUnit_0_chosen")))
    chosen_link = unit_container.find_element(By.CLASS_NAME, "chosen-single")
    driver.execute_script("arguments[0].click();", chosen_link)
    chosen_input = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='ddlUnit_0_chosen']//input")))
    chosen_input.clear()
    chosen_input.send_keys("Electric Motor")
    time.sleep(1.5)
    try:
        unit_option = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='ddlUnit_0_chosen']//li[contains(text(), 'Electric Motor')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", unit_option)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", unit_option)
        print("✅ Selected Unit: Electric Motor")
    except:
        print("⚠️ Option not clickable, sending ENTER as fallback.")
        chosen_input.send_keys(Keys.ENTER)
        print("✅ Selected Unit by ENTER key")
except Exception as e:
    print(f"❌ Failed at Unit dropdown: {e}")
# === SELECT VALUE: "1000" ===
try:
    print("🔄 Selecting Value: 1000...")
    value_container = wait.until(EC.presence_of_element_located((By.ID, "ddlValue_0_chosen")))
    chosen_link = value_container.find_element(By.CLASS_NAME, "chosen-single")
    driver.execute_script("arguments[0].click();", chosen_link)
    chosen_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@id='ddlValue_0_chosen']//input")))
    chosen_input.clear()
    chosen_input.send_keys("1000")
    time.sleep(1.5)
    try:
        value_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='ddlValue_0_chosen']//li[contains(text(), '1000')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", value_option)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", value_option)
        chosen_input.send_keys(Keys.TAB)  # ✅ Ensure value is committed
        print("✅ Selected Value: 1000")
    except Exception as inner_e:
        print(f"⚠️ Option not clickable, sending ENTER as fallback. {inner_e}")
        chosen_input.send_keys(Keys.ENTER)
        chosen_input.send_keys(Keys.TAB)
        print("✅ Selected Value by ENTER key")
except Exception as e:
    print(f"❌ Failed at Value dropdown: {e}")
# ========== Step 18: Submit ==========
try:
    # Wait for the Submit button to be present and clickable
    submit_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-success') and text()='Submit']")))
    # Click the Submit button
    submit_button.click()
    print("✔️ Product Submitted Successfully ✅")
except Exception as e:
    print(f"❌ Failed at Step 18 (Submit): {e}")
# Close Browser
time.sleep(3)
driver.quit()   