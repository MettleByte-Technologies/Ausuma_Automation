import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# ------------------------- CONFIG -------------------------
EMAIL = "ola123@yopmail.com"
PASSWORD = "1"
LOGIN_URL = "https://softwaredevelopmentsolution.com/Accounting/Reports/PartyLedger"
CUSTOMERS = [
    "Ausuma CUST",
    "Customer Mails",
    "Jems OLA Vendor",
    "Jimmy Turner",
    "Jonny Deep",
    "JS OLA Vendor",
    "Liam  White",
    "Mango Co. Pvt. Ltd.",
    "Adam Adams"
]
# ---------------------------------------------------------


def get_user_input():
    """Tkinter single popup to collect inputs (excluding email/password)"""
    def submit():
        creds['customer_name'] = customer_var.get()
        creds['from_date'] = from_date_var.get()
        creds['to_date'] = to_date_var.get()
        creds['enable_currency'] = currency_var.get() == 'Yes'
        root.quit()
        root.destroy()

    creds = {}
    root = tk.Tk()
    root.title("Party Ledger Report Inputs")

    tk.Label(root, text="Customer Name").grid(row=0, column=0, sticky="w")
    customer_var = tk.StringVar(value=CUSTOMERS[0])
    customer_dropdown = ttk.Combobox(root, textvariable=customer_var, values=CUSTOMERS, width=37)
    customer_dropdown.grid(row=0, column=1)

    tk.Label(root, text="From Date (MM/DD/YYYY)").grid(row=1, column=0, sticky="w")
    from_date_var = tk.StringVar(value="01/01/2024")
    tk.Entry(root, textvariable=from_date_var, width=40).grid(row=1, column=1)

    tk.Label(root, text="To Date (MM/DD/YYYY)").grid(row=2, column=0, sticky="w")
    to_date_var = tk.StringVar(value="04/30/2025")
    tk.Entry(root, textvariable=to_date_var, width=40).grid(row=2, column=1)

    tk.Label(root, text="Enable Currency").grid(row=3, column=0, sticky="w")
    currency_var = tk.StringVar(value="Yes")
    ttk.Combobox(root, textvariable=currency_var, values=["Yes", "No"], width=37).grid(row=3, column=1)

    tk.Button(root, text="Submit", command=submit).grid(row=4, columnspan=2, pady=10)

    root.mainloop()
    return creds


def run_script():
    user_input = get_user_input()

    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 20)
    driver.maximize_window()
    driver.get(LOGIN_URL)

    # Login
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys(EMAIL)
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys(PASSWORD)
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
    print("✅ Logged in successfully")

    # Wait for Party Ledger page
    wait.until(EC.presence_of_element_located((By.ID, "radio1")))

    # Select "Customer" radio button
    radio_button = driver.find_element(By.ID, "radio1")
    if not radio_button.is_selected():
        radio_button.click()

    # Open Bootstrap dropdown
    dropdown_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-id='sl']")))
    dropdown_btn.click()
    print("🔽 Dropdown opened")

    # Search for customer
    search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".bs-searchbox input")))
    search_input.clear()
    search_input.send_keys(user_input['customer_name'])
    print(f"🔍 Searching for customer: {user_input['customer_name']}")
    time.sleep(1.5)  # Wait for filtering to complete

    # Select the customer
    customer_xpath = f"//span[contains(text(), '{user_input['customer_name']}')]"
    try:
        customer_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, customer_xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", customer_option)
        customer_option.click()
        print(f"✅ Customer '{user_input['customer_name']}' selected")
    except:
        print(f"❌ Failed to find customer: {user_input['customer_name']}")
        driver.quit()
        return

    # Fill in dates
    from_input = wait.until(EC.presence_of_element_located((By.ID, "FromDate")))
    from_input.clear()
    from_input.send_keys(user_input['from_date'])

    to_input = wait.until(EC.presence_of_element_located((By.ID, "ToDate")))
    to_input.clear()
    to_input.send_keys(user_input['to_date'])

    # Currency checkbox
    currency_checkbox = driver.find_element(By.ID, "EnableCurrency")
    if user_input['enable_currency'] != currency_checkbox.is_selected():
        currency_checkbox.click()
    print(f"💱 Currency checkbox set to: {user_input['enable_currency']}")

    # Generate Report
    wait.until(EC.element_to_be_clickable((By.ID, "submitbtn"))).click()
    print("📊 Report generation triggered")

    # Wait for loading to finish
    try:
        wait.until(EC.invisibility_of_element_located((By.ID, "processing-scroll")))
    except:
        print("⚠️ Timeout waiting for loading overlay.")

    # Export to Excel
    export_btn = wait.until(EC.element_to_be_clickable((By.ID, "exportbtn")))
    driver.execute_script("arguments[0].click();", export_btn)
    print("⬇️ Export dropdown opened")

    excel_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Excel')]")))
    driver.execute_script("arguments[0].click();", excel_option)
    print("📁 Excel export triggered")

    time.sleep(15)  # Give time for download to finish
    driver.quit()
    print("🚪 Browser closed")


if __name__ == "__main__":
    run_script()
