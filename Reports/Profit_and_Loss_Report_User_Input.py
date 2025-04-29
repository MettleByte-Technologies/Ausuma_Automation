import time
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def run_script(from_date, to_date, exclude_currency, enable_currency):
    driver = webdriver.Edge()
    driver.maximize_window()
    driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/ProfitAndLoss")

    wait = WebDriverWait(driver, 20)

    # Login
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
    print("✅ Logged in successfully")

    # Fill From Date
    wait.until(EC.presence_of_element_located((By.ID, "FromDate"))).send_keys(from_date)
    print(f"From Date set to {from_date}")

    # Fill To Date
    wait.until(EC.presence_of_element_located((By.ID, "ToDate"))).send_keys(to_date)
    print(f"To Date set to {to_date}")

    # Handle checkboxes using JavaScript to avoid click interception
    if exclude_currency:
        driver.execute_script("""
            let el = document.getElementById("ZeroBalanceAccount");
            if (!el.checked) el.click();
        """)
        print("✅ Exclude Currency checked")
    else:
        driver.execute_script("""
            let el = document.getElementById("ZeroBalanceAccount");
            if (el.checked) el.click();
        """)

    if enable_currency:
        driver.execute_script("""
            let el = document.getElementById("EnableCurrency");
            if (!el.checked) el.click();
        """)
        print("✅ Enable Currency checked")
    else:
        driver.execute_script("""
            let el = document.getElementById("EnableCurrency");
            if (el.checked) el.click();
        """)

    # Click Generate Report
    wait.until(EC.element_to_be_clickable((By.ID, "submitbtn"))).click()
    time.sleep(15)
    print("✅ Report generated")

def open_popup():
    def submit():
        from_date = entry_from_date.get()
        to_date = entry_to_date.get()
        exclude_currency = var_exclude_currency.get()
        enable_currency = var_enable_currency.get()

        if not from_date or not to_date:
            messagebox.showerror("Error", "Please fill all fields.")
            return

        popup.destroy()
        run_script(from_date, to_date, exclude_currency, enable_currency)

    popup = tk.Tk()
    popup.title("Enter Report Details")

    tk.Label(popup, text="From Date (MM/DD/YYYY):").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    entry_from_date = tk.Entry(popup)
    entry_from_date.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(popup, text="To Date (MM/DD/YYYY):").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    entry_to_date = tk.Entry(popup)
    entry_to_date.grid(row=1, column=1, padx=10, pady=5)

    var_exclude_currency = tk.BooleanVar()
    tk.Checkbutton(popup, text="Exclude Currency", variable=var_exclude_currency).grid(row=2, columnspan=2, pady=5)

    var_enable_currency = tk.BooleanVar()
    tk.Checkbutton(popup, text="Enable Currency", variable=var_enable_currency).grid(row=3, columnspan=2, pady=5)

    tk.Button(popup, text="Submit", command=submit).grid(row=4, columnspan=2, pady=10)

    popup.mainloop()

open_popup()
