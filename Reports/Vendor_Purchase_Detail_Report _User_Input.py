import time
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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

def get_user_inputs():
    def submit():
        nonlocal selected_vendor, from_date, to_date, enable_currency, include_returns, exclude_fully_returned, export_format
        selected_vendor = vendor_combo.get()
        from_date = from_entry.get()
        to_date = to_entry.get()
        enable_currency = currency_var.get()
        include_returns = returns_var.get()
        exclude_fully_returned = exclude_returns_var.get()
        export_format = export_combo.get()
        root.destroy()

    selected_vendor = None
    from_date = None
    to_date = None
    enable_currency = None
    include_returns = None
    exclude_fully_returned = None
    export_format = None

    root = tk.Tk()
    root.title("Open Purchase Order - Input")

    returns_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Include Returns?", variable=returns_var).pack(pady=5)

    exclude_returns_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Exclude Fully Returned Invoices?", variable=exclude_returns_var).pack(pady=5)

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

    return selected_vendor, from_date, to_date, enable_currency, include_returns, exclude_fully_returned, export_format

def run_automation():
    selected_vendor, from_date, to_date, enable_currency, include_returns, exclude_fully_returned, selected_export = get_user_inputs()

    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 20)
    driver.maximize_window()
    driver.get("https://softwaredevelopmentsolution.com/Purchase/Reports/VendorPurchaseDetail")

    # Login
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()

    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "preloader-it")))

    # Include Returns checkbox
    include_returns_checkbox = wait.until(EC.presence_of_element_located((By.ID, "includeReturns")))
    if include_returns != include_returns_checkbox.is_selected():
        include_returns_checkbox.click()
    time.sleep(2)
    # Exclude Fully Returned Invoices checkbox
    exclude_returns_checkbox = wait.until(EC.presence_of_element_located((By.ID, "excludeFullyReturnedInvoice")))
    if exclude_fully_returned != exclude_returns_checkbox.is_selected():
        exclude_returns_checkbox.click()

    # Vendor dropdown
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.btn-group.bootstrap-select button"))).click()
    time.sleep(1)
    vendor_items = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.dropdown-menu.inner li")))

    for item in vendor_items:
        if selected_vendor in item.text.strip():
            item.click()
            break
    time.sleep(1)

    # Date fields
    wait.until(EC.presence_of_element_located((By.ID, "FromDate"))).clear()
    wait.until(EC.presence_of_element_located((By.ID, "FromDate"))).send_keys(from_date)
    wait.until(EC.presence_of_element_located((By.ID, "ToDate"))).clear()
    wait.until(EC.presence_of_element_located((By.ID, "ToDate"))).send_keys(to_date)

    # Enable currency
    currency_checkbox = wait.until(EC.presence_of_element_located((By.ID, "EnableCurrency")))
    if enable_currency != currency_checkbox.is_selected():
        currency_checkbox.click()

    # Generate Report
    wait.until(EC.element_to_be_clickable((By.ID, "submitbtn"))).click()
    time.sleep(3)

    # Export dropdown
    export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Export')]"))).click()
    print("✅ Export menu opened")

    # Select export option
    export_xpath = f"//ul[@class='dropdown-menu show']//a[contains(text(), '{selected_export}')]"
    wait.until(EC.element_to_be_clickable((By.XPATH, export_xpath))).click()
    print(f"✅ Exported as: {selected_export}")

    time.sleep(10)
    driver.quit()

if __name__ == "__main__":
    run_automation()
