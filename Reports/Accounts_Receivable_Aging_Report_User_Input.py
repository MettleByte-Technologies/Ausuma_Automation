import threading
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Global form data
form_data = {}

# Dropdown values
customer_list = [
    "All", "Lucas Adams", "Harry Smith", "John White", "Jimmy Turner", "Ola Shop001",
    "BroToMotive", "OLA S&P", "William Turner", "OLA MFR Vendor", "TEST CUSTOMER",
    "OLA CUSTOMER/VENDOR CV", "OLA CUSTOMER 001362XZT", "Test Customer.0", "Jems OLA Customer",
    "OLA Service Customer/Vendor", "OLA SERVICE  CUSTOMER", "JS OLA Customer"
]

unit_options = ["Days", "Months", "Years"]

# Tkinter popup for user input
def show_popup():
    def submit():
        form_data["customer"] = customer_combo.get()
        for i in range(1, 6):
            form_data[f'age{i}'] = age_vars[i - 1].get()
            form_data[f'unit{i}'] = unit_vars[i - 1].get()
        form_data["show_invoice_count"] = show_invoice_count_var.get()
        form_data["exclude_zero_balance"] = exclude_zero_balance_var.get()
        form_data["include_returns"] = include_returns_var.get()
        form_data["enable_currency"] = enable_currency_var.get()
        popup.destroy()

    popup = tk.Tk()
    popup.title("Accounts Receivable Aging Filter")
    popup.geometry("320x600")

    tk.Label(popup, text="Select Customer:").pack(pady=5)
    customer_combo = ttk.Combobox(popup, values=customer_list, state="readonly")
    customer_combo.pack()
    customer_combo.set("All")

    age_vars = []
    unit_vars = []

    for i in range(1, 6):
        tk.Label(popup, text=f"Age {i} Value:").pack()
        age_var = tk.StringVar()
        tk.Entry(popup, textvariable=age_var).pack()

        tk.Label(popup, text=f"Age {i} Unit:").pack()
        unit_var = tk.StringVar(value="Days")
        ttk.Combobox(popup, textvariable=unit_var, values=unit_options, state="readonly").pack()

        age_vars.append(age_var)
        unit_vars.append(unit_var)

    # Checkboxes
    show_invoice_count_var = tk.BooleanVar()
    exclude_zero_balance_var = tk.BooleanVar()
    include_returns_var = tk.BooleanVar(value=True)
    enable_currency_var = tk.BooleanVar(value=True)

    tk.Checkbutton(popup, text="Show Invoice Count", variable=show_invoice_count_var).pack(pady=5)
    tk.Checkbutton(popup, text="Exclude Zero Balance", variable=exclude_zero_balance_var).pack(pady=5)
    tk.Checkbutton(popup, text="Include Returns", variable=include_returns_var).pack(pady=5)
    tk.Checkbutton(popup, text="Enable Currency", variable=enable_currency_var).pack(pady=5)

    tk.Button(popup, text="Submit", command=submit).pack(pady=15)
    popup.mainloop()

# Selenium automation
def run_selenium():
    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 15)

    driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/AccountsReceivableAging")

    # Login
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
    print("✅ Logged in")

    # Wait for page to load and ensure no preloader is visible (increased timeout)
    try:
        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.CLASS_NAME, "preloader-it")))
        print("✅ Preloader is gone")
    except Exception as e:
        print(f"⚠️ Timeout while waiting for preloader: {e}")

    # Wait for customer dropdown to be clickable
    customer_dropdown = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-id='ddlCustVend']")))

    # Click to open the dropdown if a customer other than 'All' is selected
    if form_data["customer"] != "All":
        customer_dropdown.click()
        options = driver.find_elements(By.CSS_SELECTOR, "ul.dropdown-menu.inner li a span.title")
        for option in options:
            if form_data["customer"] in option.text:
                option.click()
                print(f"✅ Selected Customer: {form_data['customer']}")
                break
        else:
            print(f"❌ Customer '{form_data['customer']}' not found!")
            return
    else:
        print("✅ Customer 'All' selected, no customer filter applied.")

    # Fill age fields
    for i in range(1, 6):
        age_input_id = f"daytype{i}Input"
        try:
            age_input = wait.until(EC.presence_of_element_located((By.ID, age_input_id)))
            age_input.clear()
            age_input.send_keys(form_data[f'age{i}'])
            print(f"📝 Filled Age {i}: {form_data[f'age{i}']}")
        except Exception as e:
            print(f"⚠️ Error filling {age_input_id}: {e}")

        # Handle Chosen.js unit dropdown
        chosen_id = f"DaysType{i}{i}_chosen" if i != 4 else "DaysType4_chosen"  # Special case for Age 4
        try:
            # Wait for any possible overlay or popup to disappear
            WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.CLASS_NAME, "sweet-alert")))

            dropdown_div = wait.until(EC.element_to_be_clickable((By.ID, chosen_id)))
            dropdown_div.click()

            unit_value = form_data[f'unit{i}']
            option_xpath = f"//div[@id='{chosen_id}']//ul/li[text()='{unit_value}']"
            unit_option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
            unit_option.click()
            print(f"✅ Selected Unit {i}: {unit_value}")
        except Exception as e:
            print(f"⚠️ Error selecting unit for Age {i}: {e}")

    # Handle checkboxes
    if form_data["show_invoice_count"]:
        wait.until(EC.presence_of_element_located((By.ID, "showInvoiceCount"))).click()
        print("✅ Show Invoice Count checked")
    
    if form_data["exclude_zero_balance"]:
        wait.until(EC.presence_of_element_located((By.ID, "ExculdeZeroOutstandingCustomer"))).click()
        print("✅ Exclude Zero Balance checked")
    
    if form_data["include_returns"]:
        wait.until(EC.presence_of_element_located((By.ID, "IncludeReturns"))).click()
        print("✅ Include Returns checked")
    
    if form_data["enable_currency"]:
        wait.until(EC.presence_of_element_located((By.ID, "EnableCurrency"))).click()
        print("✅ Enable Currency checked")

    # Click Generate Report
    wait.until(EC.element_to_be_clickable((By.ID, "submitbtn"))).click()
    print("📄 Report generated")
    time.sleep(10)
    driver.quit()

# Run popup first
popup_thread = threading.Thread(target=show_popup)
popup_thread.start()
popup_thread.join()

# Start automation
if "customer" in form_data:
    run_selenium()
else:
    print("❌ No customer selected. Automation cancelled.")
