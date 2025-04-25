import time
from tkinter import Tk, Label, Button, StringVar, OptionMenu
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# --- Tkinter Popup ---
def show_popup():
    global selected_option
    popup = Tk()
    popup.title("Select Type")

    Label(popup, text="Select Customer or Vendor:").pack(pady=10)

    selected_option = StringVar(popup)
    selected_option.set("Customer")  # default value
    OptionMenu(popup, selected_option, "Customer", "Vendor").pack(pady=5)

    Button(popup, text="Submit", command=popup.destroy).pack(pady=10)
    popup.mainloop()

# --- Show the popup first ---
show_popup()
selection = selected_option.get()

# --- Set up Edge WebDriver ---
edge_options = EdgeOptions()
edge_options.add_argument("--start-maximized")
driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=edge_options)

# --- Open login page ---
url = "https://softwaredevelopmentsolution.com/Accounting/JournalEntry/Create"
driver.get(url)

# --- Wait until we are on the expected URL or the login form is loaded ---
wait = WebDriverWait(driver, 20)

try:
    # Wait until login field is visible
    wait.until(EC.presence_of_element_located((By.ID, "email"))).send_keys("ola123@yopmail.com")
    driver.find_element(By.ID, "password").send_keys("1")

    # Click the login button
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Log')]")))
    login_button.click()

    # Wait until Journal Entry page loads
    wait.until(EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Journal Entry')]")))

    # Handle Customer/Vendor Dropdown Selection
    dropdown_element = wait.until(EC.presence_of_element_located((By.ID, "CustomerVendor")))
    select = Select(dropdown_element)

    # Match dropdown value
    for option in select.options:
        if selection.lower() in option.text.lower():
            select.select_by_visible_text(option.text)
            break

    print(f"✅ Selected '{selection}' from Customer/Vendor dropdown successfully.")

except Exception as e:
    print(f"❌ Error occurred: {e}")
    print(f"Page title: {driver.title}")
    print(f"Page URL: {driver.current_url}")

# --- Optional: pause for manual check ---
time.sleep(5)
driver.quit()
