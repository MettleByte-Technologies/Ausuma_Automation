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

# Tkinter popup for vendor selection
def get_vendor_selection():
    def submit():
        nonlocal selected_vendor
        selected_vendor = vendor_combo.get()
        root.destroy()

    selected_vendor = None
    root = tk.Tk()
    root.title("Select Vendor")
    tk.Label(root, text="Choose Vendor:").pack(pady=10)

    vendor_combo = ttk.Combobox(root, values=list(VENDOR_OPTIONS.keys()), state="readonly")
    vendor_combo.set("All")
    vendor_combo.pack(pady=5)

    tk.Button(root, text="Submit", command=submit).pack(pady=10)
    root.mainloop()
    return selected_vendor

# Selenium automation
def run_automation():
    selected_vendor = get_vendor_selection()
    vendor_value = VENDOR_OPTIONS[selected_vendor]

    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 20)
    driver.maximize_window()
    driver.get("https://softwaredevelopmentsolution.com/Purchase/Reports/OpenPurchaseOrder")

    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()

    # Wait for page to load after login
    wait.until(EC.presence_of_element_located((By.ID, "ddlCustVend")))

    # Open dropdown
    dropdown_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.btn-group.bootstrap-select button")))
    dropdown_btn.click()

    # Search and click matching vendor
    vendor_items = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.dropdown-menu.inner li")))
    for item in vendor_items:
        text = item.text.strip()
        if selected_vendor in text:
            item.click()
            break

    time.sleep(1)

    # Click Generate Report
    wait.until(EC.element_to_be_clickable((By.ID, "submitbtn"))).click()

    # Keep browser open for observation
    time.sleep(15)
    driver.quit()

# Run the script
run_automation()
