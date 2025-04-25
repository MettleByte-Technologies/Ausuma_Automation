import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import threading

# Dropdown values for filters
account_class_list = [
    "All", "Asset", "Expense", "Liability", "Equity", "Revenue"
]

account_type_options = [
    "All", "Current Liabilities", "Fixed Assets", "Intangible Assets", "Current Assets", "Other Assets",
    "Long Term Liabilities", "Owners Equity", "Sales Income", "Other Income", "Direct Cost", "Admin Expenses"
]

account_detail_type_options = [
    "All", "Goodwill", "Cash", "Inventories", "Accounts Receivables", "Advance & PrePayments", "Other Current Asset"
]

account_group_options = [
    "All", "Expenses", "Purchase Expense", "Liability", "current", "saving"
]

# Global variables for selected values
selected_account_class = None
selected_account_type = None
selected_account_detail_type = None
selected_account_group = None
exclude_zero_balance = False
enable_currency = False

# Tkinter popup to select filters
def show_popup():
    global selected_account_class, selected_account_type, selected_account_detail_type, selected_account_group, exclude_zero_balance, enable_currency

    def submit():
        global selected_account_class, selected_account_type, selected_account_detail_type, selected_account_group, exclude_zero_balance, enable_currency
        selected_account_class = account_class_combo.get()
        selected_account_type = account_type_combo.get()
        selected_account_detail_type = account_detail_type_combo.get()
        selected_account_group = account_group_combo.get()
        exclude_zero_balance = zero_balance_var.get()
        enable_currency = enable_currency_var.get()
        root.destroy()

    root = tk.Tk()
    root.title("Select Filters")

    # Account Class Dropdown
    tk.Label(root, text="Choose an Account Class:").pack(pady=5)
    account_class_combo = ttk.Combobox(root, values=account_class_list, width=40)
    account_class_combo.pack()
    account_class_combo.current(0)  # Default to 'All'

    # Account Type Dropdown
    tk.Label(root, text="Choose an Account Type:").pack(pady=5)
    account_type_combo = ttk.Combobox(root, values=account_type_options, width=40)
    account_type_combo.pack()
    account_type_combo.current(0)

    # Account Detail Type Dropdown
    tk.Label(root, text="Choose an Account Detail Type:").pack(pady=5)
    account_detail_type_combo = ttk.Combobox(root, values=account_detail_type_options, width=40)
    account_detail_type_combo.pack()
    account_detail_type_combo.current(0)

    # Account Group Dropdown
    tk.Label(root, text="Choose an Account Group:").pack(pady=5)
    account_group_combo = ttk.Combobox(root, values=account_group_options, width=40)
    account_group_combo.pack()
    account_group_combo.current(0)

    # Exclude Zero Balance Checkbox
    zero_balance_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Exclude Zero Balance", variable=zero_balance_var).pack(pady=5)

    # Enable Currency Checkbox
    enable_currency_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Enable Currency", variable=enable_currency_var).pack(pady=5)

    # Submit Button
    tk.Button(root, text="Run", command=submit).pack(pady=10)
    root.mainloop()

# Selenium automation
def run_automation():
    global selected_account_class, selected_account_type, selected_account_detail_type, selected_account_group, exclude_zero_balance, enable_currency

    # Setup WebDriver (Edge in this example)
    options = webdriver.EdgeOptions()
    options.add_argument("start-maximized")
    driver = webdriver.Edge(options=options)

    # Open the URL
    driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/AccountList")

    # Wait setup
    wait = WebDriverWait(driver, 30)  # Increased wait time

    # Login
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
    print("✅ Logged in")

    # Wait for the page to load (checking for the visibility of the preloader instead of invisibility)
    try:
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "preloader-it")))
    except Exception as e:
        print(f"Warning: {e}, attempting a longer wait.")

    # Wait for the Account Class dropdown to be present
    wait.until(EC.presence_of_element_located((By.ID, "ddlAccountClass_chosen")))

    # Open the Account Class dropdown (Chosen.js based dropdown)
    wait.until(EC.element_to_be_clickable((By.ID, "ddlAccountClass_chosen"))).click()
    time.sleep(1)

    # Select the Account Class (from popup selection)
    account_class_xpath = f"//ul[@class='chosen-results']//li[contains(text(), '{selected_account_class}')]"
    wait.until(EC.element_to_be_clickable((By.XPATH, account_class_xpath))).click()
    print(f"✅ Account Class selected: {selected_account_class}")

    # Repeat for Account Type
    wait.until(EC.element_to_be_clickable((By.ID, "ddlAccountType_chosen"))).click()
    account_type_xpath = f"//ul[@class='chosen-results']//li[contains(text(), '{selected_account_type}')]"
    try:
        # Click using JavaScript to avoid click interception
        driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, account_type_xpath))))
    except Exception as e:
        print(f"Error: {e}")

    # Repeat for Account Detail Type
    wait.until(EC.element_to_be_clickable((By.ID, "ddlAccountDetailType_chosen"))).click()
    account_detail_type_xpath = f"//ul[@class='chosen-results']//li[contains(text(), '{selected_account_detail_type}')]"
    try:
        driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, account_detail_type_xpath))))
    except Exception as e:
        print(f"Error: {e}")

    # Repeat for Account Group
    wait.until(EC.element_to_be_clickable((By.ID, "ddlAccountGroup_chosen"))).click()
    account_group_xpath = f"//ul[@class='chosen-results']//li[contains(text(), '{selected_account_group}')]"
    try:
        driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, account_group_xpath))))
    except Exception as e:
        print(f"Error: {e}")

    # Click both checkboxes if selected
    if exclude_zero_balance:
        driver.find_element(By.ID, "ZeroBalanceAccount").click()
    if enable_currency:
        driver.find_element(By.ID, "EnableCurrency").click()

    # Wait for the Generate Report button to be present
    generate_btn = wait.until(EC.presence_of_element_located((By.ID, "submitbtn")))

    # Scroll into view and click using JavaScript to avoid interception
    driver.execute_script("arguments[0].scrollIntoView(true);", generate_btn)
    time.sleep(1)  # slight pause to ensure UI stabilizes
    driver.execute_script("arguments[0].click();", generate_btn)
    print("📄 Generate Report button clicked using JavaScript")

    # Optional: wait to see the report load
    time.sleep(25)

    # Close the browser
    driver.quit()

# Show popup in a separate thread so it doesn't block the automation
threading.Thread(target=show_popup).start()

# Wait until user makes a selection in the popup
while selected_account_class is None:
    time.sleep(1)

# Start automation after the popup selection
run_automation()
