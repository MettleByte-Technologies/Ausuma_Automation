import time
import tkinter as tk
from tkinter import messagebox, ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def run_script(from_date, to_date, exclude_currency, enable_currency, export_format):
    driver = webdriver.Edge()
    driver.maximize_window()
    driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/ProfitAndLoss")

    wait = WebDriverWait(driver, 30)

    # Login
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
    print("✅ Logged in successfully")

    # Wait for page to load and date fields to appear
    wait.until(EC.presence_of_element_located((By.ID, "FromDate")))
    wait.until(EC.presence_of_element_located((By.ID, "ToDate")))

    # Set dates using JavaScript
    driver.execute_script(f"document.getElementById('FromDate').value = '{from_date}';")
    driver.execute_script(f"document.getElementById('ToDate').value = '{to_date}';")
    print(f"📅 From Date: {from_date}, To Date: {to_date}")

    # Checkboxes
    if exclude_currency:
        driver.execute_script("""
            let el = document.getElementById("ZeroBalanceAccount");
            if (!el.checked) el.click();
        """)
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
    else:
        driver.execute_script("""
            let el = document.getElementById("EnableCurrency");
            if (el.checked) el.click();
        """)

    # Wait for preloader to disappear (opacity = 0 or element gone)
    try:
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "preloader-it")))
        print("⏳ Preloader disappeared")
    except TimeoutException:
        print("⚠️ Preloader timeout; attempting to proceed anyway")

    # Click Generate Report via JS (to avoid click interception)
    generate_btn = wait.until(EC.presence_of_element_located((By.ID, "submitbtn")))
    driver.execute_script("arguments[0].click();", generate_btn)
    print("✅ Report generated")

    time.sleep(2)  # Allow time for report to load

    # Export report
    if export_format.lower() in ["pdf", "excel", "print"]:
        driver.execute_script(f"GenerateReport('{export_format.capitalize()}')")
        print(f"✅ Exported as {export_format.capitalize()}")
    else:
        print("⚠️ Invalid export format selected.")
    time.sleep(10)  # Allow time for export to complete
def open_popup():
    def submit():
        from_date = entry_from_date.get()
        to_date = entry_to_date.get()
        exclude_currency = var_exclude_currency.get()
        enable_currency = var_enable_currency.get()
        export_format = export_var.get()

        if not from_date or not to_date or export_format == "":
            messagebox.showerror("Error", "Please fill all fields.")
            return

        popup.destroy()
        run_script(from_date, to_date, exclude_currency, enable_currency, export_format)

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

    tk.Label(popup, text="Export Format:").grid(row=4, column=0, padx=10, pady=5, sticky="e")
    export_var = tk.StringVar()
    export_dropdown = ttk.Combobox(popup, textvariable=export_var, values=["Pdf", "Excel", "Print"], state="readonly")
    export_dropdown.grid(row=4, column=1, padx=10, pady=5)
    export_dropdown.set("Pdf")

    tk.Button(popup, text="Submit", command=submit).grid(row=5, columnspan=2, pady=10)

    popup.mainloop()

open_popup()
