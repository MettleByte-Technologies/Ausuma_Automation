import time
import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# ------------------------
# Tkinter GUI for Input
# ------------------------
def submit_form():
    global user_inputs
    user_inputs = {
        'journal_type': journal_type_var.get(),
        'from_date': from_date_var.get(),
        'to_date': to_date_var.get(),
        'enable_currency': enable_currency_var.get(),
        'export_type': export_type_var.get()
    }
    root.destroy()

root = tk.Tk()
root.title("Report Generation Input")

# Journal Type Dropdown
tk.Label(root, text="Journal Type:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
journal_type_var = tk.StringVar()
journal_types = [
    "All", "Purchase Invoice", "Purchase Return", "Sale Invoice", 
    "Sale Return", "Stock Adjustment", "Stock Transfer", "Opening Entry", "Journal Entry"
]
journal_type_dropdown = ttk.Combobox(root, textvariable=journal_type_var, values=journal_types, state="readonly")
journal_type_dropdown.grid(row=0, column=1, padx=10, pady=5)
journal_type_dropdown.current(0)

# From Date
tk.Label(root, text="From Date (dd/mm/yyyy):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
from_date_var = tk.StringVar()
from_date_entry = tk.Entry(root, textvariable=from_date_var)
from_date_entry.grid(row=1, column=1, padx=10, pady=5)

# To Date
tk.Label(root, text="To Date (dd/mm/yyyy):").grid(row=2, column=0, padx=10, pady=5, sticky="w")
to_date_var = tk.StringVar()
to_date_entry = tk.Entry(root, textvariable=to_date_var)
to_date_entry.grid(row=2, column=1, padx=10, pady=5)

# Enable Currency Checkbox
enable_currency_var = tk.BooleanVar()
enable_currency_checkbox = tk.Checkbutton(root, text="Enable Currency", variable=enable_currency_var)
enable_currency_checkbox.grid(row=3, column=1, padx=10, pady=5, sticky="w")

# Export Type Dropdown
tk.Label(root, text="Export Type:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
export_type_var = tk.StringVar()
export_types = ["Pdf", "Excel", "Print"]
export_type_dropdown = ttk.Combobox(root, textvariable=export_type_var, values=export_types, state="readonly")
export_type_dropdown.grid(row=4, column=1, padx=10, pady=5)
export_type_dropdown.current(0)

# Submit Button
submit_button = tk.Button(root, text="Submit", command=submit_form)
submit_button.grid(row=5, column=0, columnspan=2, pady=10)

root.mainloop()

# ------------------------
# Selenium Automation
# ------------------------

# Setup WebDriver
driver = webdriver.Edge()
wait = WebDriverWait(driver, 20)
driver.maximize_window()

# Open the website
driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/Journal")

# Login
wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
print("✅ Logged in successfully")

# Wait for page load
time.sleep(3)

# Select Journal Type
journal_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "ddlJournalType_chosen")))
journal_dropdown.click()

options = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#ddlJournalType_chosen ul.chosen-results li")))

for option in options:
    if option.text.strip().lower() == user_inputs['journal_type'].lower():
        option.click()
        break
else:
    print(f"⚠️ Journal Type '{user_inputs['journal_type']}' not found, selecting default 'All'.")
    options[0].click()

# Fill From Date
from_date_field = wait.until(EC.presence_of_element_located((By.ID, "FromDate")))
from_date_field.clear()
from_date_field.send_keys(user_inputs['from_date'])

# Fill To Date
to_date_field = wait.until(EC.presence_of_element_located((By.ID, "ToDate")))
to_date_field.clear()
to_date_field.send_keys(user_inputs['to_date'])

# Enable Currency Checkbox if selected
if user_inputs['enable_currency']:
    currency_checkbox = wait.until(EC.element_to_be_clickable((By.ID, "EnableCurrency")))
    if not currency_checkbox.is_selected():
        currency_checkbox.click()

# Generate Report
generate_button = wait.until(EC.element_to_be_clickable((By.ID, "submitbtn")))
generate_button.click()
print("✅ Report generated")

# Wait for report generation
time.sleep(5)

# Export
export_btn = wait.until(EC.element_to_be_clickable((By.ID, "exportbtn")))
ActionChains(driver).move_to_element(export_btn).click().perform()

time.sleep(2)

export_type_script = {
    "pdf": "GenerateReport('Pdf')",
    "excel": "GenerateReport('Excel')",
    "print": "GenerateReport('Print')"
}

export_choice = user_inputs['export_type'].lower()
if export_choice in export_type_script:
    driver.execute_script(export_type_script[export_choice])
    print(f"✅ Report exported as {user_inputs['export_type']}")
else:
    print(f"⚠️ Invalid export type '{user_inputs['export_type']}' selected.")

# Close the browser after some time
time.sleep(5)
driver.quit()
