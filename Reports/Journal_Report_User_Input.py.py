import threading
import time
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options

# Dictionary to hold user input from GUI
user_input = {}

# Tkinter GUI to collect user input
def start_gui():
    def submit():
        user_input['journal_type'] = journal_type_var.get()
        user_input['from_date'] = from_date_entry.get()
        user_input['to_date'] = to_date_entry.get()
        user_input['export_format'] = export_var.get()
        root.destroy()

    root = tk.Tk()
    root.title("Journal Report Filter")

    ttk.Label(root, text="Journal Type:").grid(row=0, column=0, padx=5, pady=5)
    journal_type_var = tk.StringVar()
    journal_type_dropdown = ttk.Combobox(root, textvariable=journal_type_var, values=["Journal", "Payment", "Receipt"])
    journal_type_dropdown.grid(row=0, column=1, padx=5, pady=5)
    journal_type_dropdown.current(0)

    ttk.Label(root, text="From Date (dd/mm/yyyy):").grid(row=1, column=0, padx=5, pady=5)
    from_date_entry = ttk.Entry(root)
    from_date_entry.grid(row=1, column=1, padx=5, pady=5)

    ttk.Label(root, text="To Date (dd/mm/yyyy):").grid(row=2, column=0, padx=5, pady=5)
    to_date_entry = ttk.Entry(root)
    to_date_entry.grid(row=2, column=1, padx=5, pady=5)

    ttk.Label(root, text="Export Format:").grid(row=3, column=0, padx=5, pady=5)
    export_var = tk.StringVar()
    export_dropdown = ttk.Combobox(root, textvariable=export_var, values=["PDF", "Excel", "Print"])
    export_dropdown.grid(row=3, column=1, padx=5, pady=5)
    export_dropdown.current(0)

    ttk.Button(root, text="Submit", command=submit).grid(row=4, columnspan=2, pady=10)
    root.mainloop()

# Selenium automation for the Journal Report
def run_selenium():
    driver = None  # Initialize driver to avoid UnboundLocalError
    try:
        # Set up Edge options and Service for correct driver initialization
        edge_options = Options()
        edge_options.add_argument("--headless")  # Optional: run headless for no UI

        # Update this with your actual Edge WebDriver path
        edgedriver_path = "C:/path/to/msedgedriver"  # Set the correct path for Edge WebDriver
        service = Service(edgedriver_path)
        driver = webdriver.Edge(service=service, options=edge_options)
        
        wait = WebDriverWait(driver, 30)

        driver = webdriver.Edge()
        wait = WebDriverWait(driver, 15)

        driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/AccountsReceivableAging")


        # Login
        wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
        driver.find_element(By.ID, "Password").send_keys("1")
        login_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnLogin")))
        login_btn.click()
        print("✅ Logged in")

        # Wait for journal type dropdown to load
        wait.until(EC.presence_of_element_located((By.ID, "ddlJournalType_chosen"))).click()
        options = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#ddlJournalType_chosen .chosen-results li")))
        for opt in options:
            if opt.text.strip() == user_input['journal_type']:
                opt.click()
                break

        # Fill in date filters
        from_date = wait.until(EC.presence_of_element_located((By.ID, "txtFromDate")))
        from_date.clear()
        from_date.send_keys(user_input['from_date'])

        to_date = wait.until(EC.presence_of_element_located((By.ID, "txtToDate")))
        to_date.clear()
        to_date.send_keys(user_input['to_date'])

        # Generate report
        wait.until(EC.element_to_be_clickable((By.ID, "btnView"))).click()
        print("📄 Report generated")

        # Wait and open export options
        time.sleep(3)
        export_btn = wait.until(EC.element_to_be_clickable((By.ID, "exportbtn")))
        driver.execute_script("arguments[0].click();", export_btn)

        export_map = {
            "PDF": "#exportDropdown li a[data-original-title='Export to PDF']",
            "Excel": "#exportDropdown li a[data-original-title='Export to Excel']",
            "Print": "#exportDropdown li a[data-original-title='Print']"
        }

        export_selector = export_map[user_input['export_format']]
        export_option = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, export_selector)))
        driver.execute_script("arguments[0].click();", export_option)
        print(f"✅ Report exported as {user_input['export_format']}")

    except Exception as e:
        print("❌ Error during automation:", e)
    finally:
        if driver:
            time.sleep(5)
            driver.quit()

# Start GUI first
start_gui()

# Then start Selenium in a separate thread
selenium_thread = threading.Thread(target=run_selenium)
selenium_thread.start()
