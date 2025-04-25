import threading
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# --- Global data container ---
form_data = {}

# --- Tkinter Popup ---
def show_popup():
    def update_fields():
        if date_type_var.get() == "d":
            from_date_label.grid()
            from_date_entry.grid()
            to_date_label.grid()
            to_date_entry.grid()
            include_returns_check.grid()
            enable_currency_check.grid()

            start_from_label.grid_remove()
            start_from_entry.grid_remove()
            age_label.grid_remove()
            age_entry.grid_remove()
            show_currency_days_check.grid_remove()
        else:
            from_date_label.grid_remove()
            from_date_entry.grid_remove()
            to_date_label.grid_remove()
            to_date_entry.grid_remove()

            include_returns_check.grid()
            enable_currency_check.grid()
            start_from_label.grid()
            start_from_entry.grid()
            age_label.grid()
            age_entry.grid()
            show_currency_days_check.grid()

    def submit():
        form_data['customer'] = customer_var.get()
        form_data['date_type'] = date_type_var.get()
        form_data['from_date'] = from_date_var.get()
        form_data['to_date'] = to_date_var.get()
        form_data['start_from'] = start_from_var.get()
        form_data['age'] = age_var.get()
        form_data['include_returns'] = include_returns_var.get()
        form_data['enable_currency'] = enable_currency_var.get()
        form_data['show_currency_days'] = show_currency_days_var.get()
        form_data['export_format'] = export_format_var.get()
        root.destroy()

    root = tk.Tk()
    root.title("Accounts Receivable Report Filters")

    # Customer Dropdown
    tk.Label(root, text="Customer:").grid(row=0, column=0, sticky="w")
    customer_var = tk.StringVar(value="0")
    customer_dropdown = ttk.Combobox(root, textvariable=customer_var, width=50, state="readonly")
    customer_dropdown['values'] = [
        ("0", "All"),
        ("70", "Lucas Adams"),
        ("71", "Harry Smith"),
        ("72", "John White"),
        ("77", "Jimmy Turner"),
        ("80", "Ola Shop001"),
        ("1081", "BroToMotive"),
        ("1082", "OLA S&P"),
        ("2084", "William Turner"),
        ("2086", "OLA MFR Vendor"),
        ("10633", "Unknown"),
        ("10634", "IMPORT 03-03-2025"),
        ("10636", "Ausuma Customer 28-03-2025"),
        ("10637", "Bhopendra Customer"),
        ("10639", "Arice CV"),
        ("10642", "AUSUMA ABC CV")
    ]
    customer_dropdown.grid(row=0, column=1)

    # Date / Age Toggle
    tk.Label(root, text="Select by:").grid(row=1, column=0, sticky="w")
    date_type_var = tk.StringVar(value="d")
    tk.Radiobutton(root, text="Date", variable=date_type_var, value="d", command=update_fields).grid(row=1, column=1, sticky="w")
    tk.Radiobutton(root, text="Age", variable=date_type_var, value="a", command=update_fields).grid(row=1, column=2, sticky="w")

    # From Date and To Date
    from_date_var = tk.StringVar()
    to_date_var = tk.StringVar()
    from_date_label = tk.Label(root, text="From Date (dd-mm-yyyy):")
    from_date_label.grid(row=2, column=0, sticky="w")
    from_date_entry = tk.Entry(root, textvariable=from_date_var)
    from_date_entry.grid(row=2, column=1)

    to_date_label = tk.Label(root, text="To Date (dd-mm-yyyy):")
    to_date_label.grid(row=3, column=0, sticky="w")
    to_date_entry = tk.Entry(root, textvariable=to_date_var)
    to_date_entry.grid(row=3, column=1)

    # Start From & Age (Hidden initially)
    start_from_var = tk.StringVar()
    age_var = tk.StringVar()
    start_from_label = tk.Label(root, text="Start From:")
    start_from_entry = tk.Entry(root, textvariable=start_from_var)
    age_label = tk.Label(root, text="Age:")
    age_entry = tk.Entry(root, textvariable=age_var)
    start_from_label.grid(row=4, column=0, sticky="w")
    start_from_entry.grid(row=4, column=1)
    age_label.grid(row=5, column=0, sticky="w")
    age_entry.grid(row=5, column=1)
    start_from_label.grid_remove()
    start_from_entry.grid_remove()
    age_label.grid_remove()
    age_entry.grid_remove()

    # Checkboxes
    include_returns_var = tk.BooleanVar()
    enable_currency_var = tk.BooleanVar()
    show_currency_days_var = tk.BooleanVar()
    include_returns_check = tk.Checkbutton(root, text="Include Returns", variable=include_returns_var)
    include_returns_check.grid(row=6, column=0, sticky="w")
    enable_currency_check = tk.Checkbutton(root, text="Enable Currency", variable=enable_currency_var)
    enable_currency_check.grid(row=6, column=1, sticky="w")
    show_currency_days_check = tk.Checkbutton(root, text="Show Currency Days", variable=show_currency_days_var)
    show_currency_days_check.grid(row=7, column=0, sticky="w")
    show_currency_days_check.grid_remove()

    # Export format dropdown
    tk.Label(root, text="Export Format:").grid(row=8, column=0, sticky="w")
    export_format_var = tk.StringVar(value="Pdf")
    export_dropdown = ttk.Combobox(root, textvariable=export_format_var, state="readonly")
    export_dropdown['values'] = ("Pdf", "Excel", "Print")
    export_dropdown.grid(row=8, column=1)

    # Submit
    tk.Button(root, text="Submit", command=submit).grid(row=9, column=1, pady=10)
    root.mainloop()

# Show the popup first
popup_thread = threading.Thread(target=show_popup)
popup_thread.start()
popup_thread.join()

# --- Selenium Automation ---
driver = webdriver.Edge()
driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/AccountsReceivableOutstanding")
wait = WebDriverWait(driver, 20)

# Login
wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
print("✅ Logged in")

# Wait for dropdown
wait.until(EC.presence_of_element_located((By.ID, "ddlCustVend")))
time.sleep(2)

# Set customer using JavaScript
customer_value = form_data['customer']
driver.execute_script(f"""
    var select = document.getElementById('ddlCustVend');
    for (var i = 0; i < select.options.length; i++) {{
        if (select.options[i].value === '{customer_value}') {{
            select.selectedIndex = i;
            break;
        }}
    }}
    $('#ddlCustVend').selectpicker('refresh');
""")

# Date or Age selection
if form_data['date_type'] == "d":
    driver.find_element(By.ID, "radio1").click()
    driver.find_element(By.ID, "FromDate").clear()
    driver.find_element(By.ID, "FromDate").send_keys(form_data['from_date'])
    driver.find_element(By.ID, "ToDate").clear()
    driver.find_element(By.ID, "ToDate").send_keys(form_data['to_date'])
else:
    driver.find_element(By.ID, "radio2").click()
    wait.until(EC.presence_of_element_located((By.ID, "startdaytypeInput"))).send_keys(form_data['start_from'])
    driver.find_element(By.ID, "daytype1Input").send_keys(form_data['age'])

# Checkboxes
if form_data['include_returns']:
    checkbox = driver.find_element(By.ID, "IncludeReturns")
    if not checkbox.is_selected():
        checkbox.click()

if form_data['enable_currency']:
    checkbox = driver.find_element(By.ID, "EnableCurrency")
    if not checkbox.is_selected():
        checkbox.click()

if form_data['show_currency_days']:
    checkbox = driver.find_element(By.ID, "showCurrentDays")
    if not checkbox.is_selected():
        checkbox.click()

# Generate Report
driver.find_element(By.ID, "submitbtn").click()
time.sleep(5)  # Wait for report to generate
print("📄 Report generation triggered")

# Export based on user selection
export_format = form_data['export_format']
driver.execute_script(f"GenerateReport('{export_format}')")
print(f"📤 Exported as {export_format}")
