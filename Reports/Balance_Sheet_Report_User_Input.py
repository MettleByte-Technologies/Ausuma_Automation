import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Dropdown options
class_options = ["All", "Asset", "Expense", "Liability", "Equity", "Revenue"]
type_options = ["All", "Intangible Assets", "Current Assets", "Fixed Assets", "Other Assets", "Current Liabilities", "Long Term Liabilities", "Owners Equity", "Sales Income", "Other Income", "Direct Cost", "Admin Expenses", "Depr. & Amortisation Expenses", "Taxes", "Utilities", "Other Expenses", "Balance Sheet", "Asset Accounts", "Income Statement", "ACC. Account Edited*", "ACC. Asset", "Income Statement of Sales", "Account test date", "New Edit Account", "Income", "ITGST   03-02-2025", "TEST ACCOUNT TYPE", "TEST ACCOUNT TYPEc", "TEST AT 05-03-2025", "Mango", "Apple Digital Production", "test", "INCOME 123", "A TEST AT", "ABC TEST AT"]
detail_type_options = ["All", "Goodwill", "Cash", "Inventories", "Accounts Receivables", "Advance & PrePayments", "Other Current Asset", "Land & Buildings", "Furniture & Fixtures", "Office Equipment & Computers", "Motor Vehicles", "Tools & Machinery", "Other  Fixed  Assets", "Miscellaneous Assets", "Accounts Payables", "Accruals", "Long Term Liabilities", "Opening Balance Equity", "Retained Earnings", "Sales Revenue Trading", "Services Revenue Trading", "Other Primary Income", "Miscellaneous Income", "Interest Earned", "Non Profit Income", "Cost Of Sales Trading"]
group_options = ["All", "Expenses", "Purchase Expense", "Liability", "current", "saving", "Show", "Cost of Goods Sold", "Financial Groups", "Financial Groups1", "Financial Groups0", "Cost of Product", "Acc Gorup", "Cost of Goods Sold Group", "Financial Groups of Sales", "Acc Financial Groups", "Test Edit Groups", "TEST GROUP AUSUMA", "GST GROUP   03-02-2025", "TEST GROUP", "TEST GROUP1", "TEST G 05-03-2025", "Apple Digital Production", "INCOME GROUP", "err", "A TEST GROUP"]
export_options = ["None", "PDF", "Excel", "Print"]

# GUI for user input
selected_values = {}

def toggle_date_fields():
    if date_range_var.get() == "d":
        from_date_entry.grid(row=10, column=1, padx=10, pady=5)
        to_date_entry.grid(row=11, column=1, padx=10, pady=5)
        column_format_cb.grid_remove()
    else:
        from_date_entry.grid_remove()
        to_date_entry.grid_remove()
        column_format_cb.grid(row=12, column=0, columnspan=2, sticky="w", padx=10)

def on_submit():
    selected_values['Class'] = class_cb.get()
    selected_values['Account Type'] = type_cb.get()
    selected_values['Account Details Type'] = detail_type_cb.get()
    selected_values['Group'] = group_cb.get()
    selected_values['Exclude Zero Balance'] = exclude_zero_var.get()
    selected_values['Hide Class'] = hide_class_var.get()
    selected_values['Hide Type'] = hide_type_var.get()
    selected_values['Hide Details Type'] = hide_detail_type_var.get()
    selected_values['Date Range'] = date_range_var.get()
    selected_values['From Date'] = from_date_var.get()
    selected_values['To Date'] = to_date_var.get()
    selected_values['Column Format'] = column_format_var.get()
    selected_values['Enable Currency'] = enable_currency_var.get()
    selected_values['Export Format'] = export_cb.get()
    root.destroy()

root = tk.Tk()
root.title("Select Filters")

ttk.Label(root, text="Class").grid(row=0, column=0, padx=10, pady=5, sticky="w")
class_cb = ttk.Combobox(root, values=class_options)
class_cb.grid(row=0, column=1, padx=10, pady=5)
class_cb.set("All")

ttk.Label(root, text="Account Type").grid(row=1, column=0, padx=10, pady=5, sticky="w")
type_cb = ttk.Combobox(root, values=type_options)
type_cb.grid(row=1, column=1, padx=10, pady=5)
type_cb.set("All")

ttk.Label(root, text="Account Details Type").grid(row=2, column=0, padx=10, pady=5, sticky="w")
detail_type_cb = ttk.Combobox(root, values=detail_type_options)
detail_type_cb.grid(row=2, column=1, padx=10, pady=5)
detail_type_cb.set("All")

ttk.Label(root, text="Group").grid(row=3, column=0, padx=10, pady=5, sticky="w")
group_cb = ttk.Combobox(root, values=group_options)
group_cb.grid(row=3, column=1, padx=10, pady=5)
group_cb.set("All")

exclude_zero_var = tk.BooleanVar()
hide_class_var = tk.BooleanVar()
hide_type_var = tk.BooleanVar()
hide_detail_type_var = tk.BooleanVar()
column_format_var = tk.BooleanVar()
enable_currency_var = tk.BooleanVar()

ttk.Checkbutton(root, text="Exclude Zero Balance", variable=exclude_zero_var).grid(row=4, column=0, columnspan=2, sticky="w", padx=10)
ttk.Checkbutton(root, text="Hide Class", variable=hide_class_var).grid(row=5, column=0, columnspan=2, sticky="w", padx=10)
ttk.Checkbutton(root, text="Hide Type", variable=hide_type_var).grid(row=6, column=0, columnspan=2, sticky="w", padx=10)
ttk.Checkbutton(root, text="Hide Details Type", variable=hide_detail_type_var).grid(row=7, column=0, columnspan=2, sticky="w", padx=10)

date_range_var = tk.StringVar(value="d")
ttk.Label(root, text="Date Selection").grid(row=8, column=0, sticky="w", padx=10)
tk.Radiobutton(root, text="Date Range", variable=date_range_var, value="d", command=toggle_date_fields).grid(row=8, column=1, sticky="w")
tk.Radiobutton(root, text="As of Today", variable=date_range_var, value="a", command=toggle_date_fields).grid(row=9, column=1, sticky="w")

from_date_var = tk.StringVar()
to_date_var = tk.StringVar()
from_date_entry = ttk.Entry(root, textvariable=from_date_var)
to_date_entry = ttk.Entry(root, textvariable=to_date_var)
from_date_entry.grid(row=10, column=1, padx=10, pady=5)
to_date_entry.grid(row=11, column=1, padx=10, pady=5)

column_format_cb = ttk.Checkbutton(root, text="Column Format", variable=column_format_var)
column_format_cb.grid(row=12, column=0, columnspan=2, sticky="w", padx=10)

ttk.Checkbutton(root, text="Enable Currency", variable=enable_currency_var).grid(row=13, column=0, columnspan=2, sticky="w", padx=10)

ttk.Label(root, text="Export Format").grid(row=14, column=0, padx=10, pady=5, sticky="w")
export_cb = ttk.Combobox(root, values=export_options)
export_cb.grid(row=14, column=1, padx=10, pady=5)
export_cb.set("None")

ttk.Button(root, text="Submit", command=on_submit).grid(row=15, column=0, columnspan=2, pady=10)

toggle_date_fields()
root.mainloop()

# === Selenium Logic ===
driver = webdriver.Edge()
wait = WebDriverWait(driver, 20)
driver.maximize_window()
driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/BalanceSheet")

# Login
wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
print("✅ Logged in successfully")

try:
    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "preloader-it")))
except Exception:
    print("⚠️ preloader-it not found or took too long to disappear.")

def select_chosen_option(chosen_base_id, visible_text):
    chosen_div = wait.until(EC.element_to_be_clickable((By.ID, f"{chosen_base_id}_chosen")))
    driver.execute_script("arguments[0].scrollIntoView(true);", chosen_div)
    chosen_div.click()
    time.sleep(0.5)
    option_xpath = f'//div[@id="{chosen_base_id}_chosen"]//li[contains(@class,"active-result") and normalize-space()="{visible_text}"]'
    wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath))).click()
    time.sleep(0.5)

select_chosen_option("ddlAccountClass", selected_values['Class'])
select_chosen_option("ddlAccountType", selected_values['Account Type'])
select_chosen_option("ddlAccountDetailType", selected_values['Account Details Type'])
select_chosen_option("ddlAccountGroup", selected_values['Group'])

def set_checkbox(checkbox_id, should_check):
    checkbox = wait.until(EC.presence_of_element_located((By.ID, checkbox_id)))
    is_checked = checkbox.is_selected()
    if should_check != is_checked:
        driver.execute_script("arguments[0].click();", checkbox)

set_checkbox("ZeroBalanceAccount", selected_values['Exclude Zero Balance'])
set_checkbox("classGrouping", selected_values['Hide Class'])
set_checkbox("typeGrouping", selected_values['Hide Type'])
set_checkbox("dtypeGrouping", selected_values['Hide Details Type'])
set_checkbox("columnarReport", selected_values['Column Format'])
set_checkbox("columnarReport", selected_values['Enable Currency'])

if selected_values['Date Range'] == 'd':
    from_input = driver.find_element(By.ID, "FromDate")
    to_input = driver.find_element(By.ID, "ToDate")
    from_input.clear()
    from_input.send_keys(selected_values['From Date'])
    to_input.clear()
    to_input.send_keys(selected_values['To Date'])
else:
    try:
        radio_button = wait.until(EC.element_to_be_clickable((By.ID, "radio2")))
        driver.execute_script("arguments[0].click();", radio_button)
        print("📅 'As of Today' radio selected.")
    except Exception as e:
        print("❌ Failed to select 'As of Today' radio button:", e)

# Generate Report
try:
    print("⚙️ Triggering Generate Report function via JS...")
    driver.execute_script("GenerateReport('generate')")
    print("📄 Generate Report triggered.")
except Exception as e:
    print("❌ Failed to generate report:", e)
    with open("report_error.html", "w", encoding="utf-8") as f:
        f.write(driver.page_source)

time.sleep(10)

# Export if selected
export_format = selected_values.get('Export Format', 'None')
if export_format.lower() in ["pdf", "excel", "print"]:
    try:
        driver.execute_script(f"GenerateReport('{export_format.capitalize()}')")
        print(f"✅ Exported as {export_format.capitalize()}")
    except Exception as e:
        print(f"❌ Failed to export as {export_format}:", e)
    time.sleep(10)
else:
    print("ℹ️ No export selected or invalid option.")

driver.quit()
print("✅ Done. Browser closed.")
