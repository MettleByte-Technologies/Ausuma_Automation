import time
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

# Hardcoded vendor options from HTML
VENDOR_OPTIONS = {
    "All": "0",
    "Arice CV": "10639",
    "Ausuma Vendor 27-03-2025": "10635",
    "BroToMotive - testvendor@yopmail.com": "6622",
    "BroToMotive - prajapatibalwant12345@gmail.com": "1081",
    "Jems OLA Vendor": "4096",
    "Jimmy Turner": "77",
    "Jonny Deep": "4121",
    "JS OLA Vendor": "4104",
    "Liam White": "74",
    "Mango Co. Pvt. Ltd.": "7625"
}

EXPORT_OPTIONS = ["Pdf", "Excel", "Print"]

# Tkinter popup for user input
def get_user_inputs():
    def submit():
        nonlocal selected_vendor, from_date, to_date, enable_currency, export_format
        selected_vendor = vendor_combo.get()
        from_date = from_entry.get()
        to_date = to_entry.get()
        enable_currency = currency_var.get()
        export_format = export_combo.get()
        root.destroy()

    selected_vendor = None
    from_date = None
    to_date = None
    enable_currency = None
    export_format = None

    root = tk.Tk()
    root.title("Open Purchase Order - Input")

    tk.Label(root, text="Choose Vendor:").pack(pady=5)
    vendor_combo = ttk.Combobox(root, values=list(VENDOR_OPTIONS.keys()), state="readonly", width=50)
    vendor_combo.set("All")
    vendor_combo.pack()

    tk.Label(root, text="From Date (MM/DD/YYYY):").pack(pady=5)
    from_entry = tk.Entry(root, width=50)
    from_entry.pack()

    tk.Label(root, text="To Date (MM/DD/YYYY):").pack(pady=5)
    to_entry = tk.Entry(root, width=50)
    to_entry.pack()

    currency_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Enable Currency", variable=currency_var).pack(pady=5)

    tk.Label(root, text="Export Format:").pack(pady=5)
    export_combo = ttk.Combobox(root, values=EXPORT_OPTIONS, state="readonly", width=50)
    export_combo.set("Excel")
    export_combo.pack()

    tk.Button(root, text="Submit", command=submit).pack(pady=10)
    root.mainloop()

    return selected_vendor, from_date, to_date, enable_currency, export_format

# Selenium automation
def run_automation():
    selected_vendor, from_date, to_date, enable_currency, export_format = get_user_inputs()
    vendor_value = VENDOR_OPTIONS[selected_vendor]

    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 20)
    driver.maximize_window()
    driver.get("https://softwaredevelopmentsolution.com/Purchase/Reports/OpenPurchaseOrder")

    # Login
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()

    # Wait for preloader to disappear
    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "preloader-it")))

    # Wait for dropdown button to appear
    dropdown_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.btn-group.bootstrap-select button")))
    dropdown_btn.click()
    time.sleep(2)
    # Select vendor
    vendor_items = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.dropdown-menu.inner li")))
    for item in vendor_items:
        if selected_vendor in item.text.strip():
            item.click()
            break
    time.sleep(2)       
    # Set From Date
    from_input = wait.until(EC.presence_of_element_located((By.ID, "FromDate")))
    from_input.clear()
    from_input.send_keys(from_date)
    time.sleep(2)
    # Set To Date
    to_input = wait.until(EC.presence_of_element_located((By.ID, "ToDate")))
    to_input.clear()
    to_input.send_keys(to_date)
    time.sleep(2)
    # Enable Currency checkbox
    currency_checkbox = wait.until(EC.presence_of_element_located((By.ID, "EnableCurrency")))
    if enable_currency != currency_checkbox.is_selected():
        currency_checkbox.click()
    time.sleep(2)
    # Click Generate Report
    wait.until(EC.element_to_be_clickable((By.ID, "submitbtn"))).click()

    # Wait for report to generate
    time.sleep(2)

    # Click Export button
    export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Export')]")))
    export_btn.click()
    print("✅ Export menu opened")

    # Choose export format
    export_xpath = f"//ul[@class='dropdown-menu show']//a[contains(text(), '{export_format}')]"
    wait.until(EC.element_to_be_clickable((By.XPATH, export_xpath))).click()
    print(f"✅ Exported as: {export_format}")

    # Keep browser open for observation
    time.sleep(10)
    driver.quit()

# Run the script
if __name__ == "__main__":
    run_automation()
