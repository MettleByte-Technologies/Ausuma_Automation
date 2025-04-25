import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
import random

form_data = {}

def generate_unique_name(prefix="Test"):
    return f"{prefix}_{int(time.time())}_{random.randint(100, 999)}"

def generate_random_number(length=10):
    return ''.join(random.choices("0123456789", k=length))

def get_form_data():
    def on_acc_detail_change(event=None):
        if acc_detail_var.get() == "Bank":
            bank_name_label.grid(row=5, column=0, padx=10, pady=5, sticky="w")
            bank_name_entry.grid(row=5, column=1, padx=10, pady=5)
            bank_type_label.grid(row=6, column=0, padx=10, pady=5, sticky="w")
            bank_type_combo.grid(row=6, column=1, padx=10, pady=5)
            account_no_label.grid(row=7, column=0, padx=10, pady=5, sticky="w")
            account_no_entry.grid(row=7, column=1, padx=10, pady=5)
            routing_no_label.grid(row=8, column=0, padx=10, pady=5, sticky="w")
            routing_no_entry.grid(row=8, column=1, padx=10, pady=5)
        else:
            bank_name_label.grid_remove()
            bank_name_entry.grid_remove()
            bank_type_label.grid_remove()
            bank_type_combo.grid_remove()
            account_no_label.grid_remove()
            account_no_entry.grid_remove()
            routing_no_label.grid_remove()
            routing_no_entry.grid_remove()

    def submit():
        form_data["name"] = name_var.get()
        form_data["description"] = desc_var.get()
        form_data["acc_detail"] = acc_detail_var.get()
        form_data["group"] = group_var.get()
        form_data["dc"] = dc_var.get()
        if form_data["acc_detail"] == "Bank":
            form_data["bank_name"] = bank_name_var.get()
            form_data["bank_type"] = bank_type_var.get()
            form_data["account_no"] = account_no_var.get()
            form_data["routing_no"] = routing_no_var.get()
        root.destroy()

    root = tk.Tk()
    root.title("Enter Account Info")

    default_name = generate_unique_name("Account")
    default_desc = "Auto-generated description"
    default_bank_name = generate_unique_name("Bank")
    default_acc_no = generate_random_number()
    default_routing = generate_random_number()

    tk.Label(root, text="Account Name").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    name_var = tk.StringVar(value=default_name)
    tk.Entry(root, textvariable=name_var).grid(row=0, column=1, padx=10, pady=5)

    tk.Label(root, text="Description").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    desc_var = tk.StringVar(value=default_desc)
    tk.Entry(root, textvariable=desc_var).grid(row=1, column=1, padx=10, pady=5)

    tk.Label(root, text="Account Detail Type").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    acc_detail_var = tk.StringVar(value="Bank")
    acc_detail_options = [
        "Acc.Financial Accounting", "Accounts Receivables", "Accruals", "Advance & PrePayments",
        "Bank", "Bank Charges", "Cash", "Cost Of Sales Services", "Cost Of Sales Trading",
        "Depreciation Expenses", "Discounts & Refunds", "Financial Accounting",
        "General & Administrative Expenses", "Furniture & Fixtures", "Miscellaneous Income",
        "Miscellaneous Expense", "Office Equipment & Computers", "Non Profit Income",
        "Payment Gateway Charges"
    ]
    acc_detail_combo = ttk.Combobox(root, textvariable=acc_detail_var, values=acc_detail_options, state="readonly")
    acc_detail_combo.grid(row=2, column=1, padx=10, pady=5)
    acc_detail_combo.bind("<<ComboboxSelected>>", on_acc_detail_change)

    tk.Label(root, text="Group").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    group_var = tk.StringVar(value="INCOME GROUP")
    group_options = [
        "Cost of Goods Sold", "Cost of Product", "Acc Gorup", "Financial Groups of Sales", "INCOME GROUP"
    ]
    ttk.Combobox(root, textvariable=group_var, values=group_options, state="readonly").grid(row=3, column=1, padx=10, pady=5)

    tk.Label(root, text="Default Debit/Credit").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    dc_var = tk.StringVar(value="Credit")
    ttk.Combobox(root, textvariable=dc_var, values=["Credit", "Debit"], state="readonly").grid(row=4, column=1, padx=10, pady=5)

    # Bank fields
    bank_name_var = tk.StringVar(value=default_bank_name)
    bank_name_label = tk.Label(root, text="Bank Name")
    bank_name_entry = tk.Entry(root, textvariable=bank_name_var)

    bank_type_var = tk.StringVar(value="Checking")
    bank_type_label = tk.Label(root, text="Bank Account Type")
    bank_type_combo = ttk.Combobox(root, textvariable=bank_type_var, values=["Checking", "Savings"], state="readonly")

    account_no_var = tk.StringVar(value=default_acc_no)
    account_no_label = tk.Label(root, text="Bank Account No.")
    account_no_entry = tk.Entry(root, textvariable=account_no_var)

    routing_no_var = tk.StringVar(value=default_routing)
    routing_no_label = tk.Label(root, text="Routing No.")
    routing_no_entry = tk.Entry(root, textvariable=routing_no_var)

    tk.Button(root, text="Submit", command=submit).grid(row=9, columnspan=2, pady=10)

    on_acc_detail_change()
    root.mainloop()

# --- Run the form ---
get_form_data()

# --- Start Selenium ---
driver = webdriver.Edge()
driver.get("https://softwaredevelopmentsolution.com/Accounting/Account/Create")
wait = WebDriverWait(driver, 15)

# --- Login block ---
try:
    username_field = wait.until(EC.presence_of_element_located((By.ID, "Email")))
    password_field = wait.until(EC.presence_of_element_located((By.ID, "Password")))
    login_button = wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit")))

    username_field.send_keys("ola123@yopmail.com")
    password_field.send_keys("1")
    login_button.click()

    wait.until(EC.presence_of_element_located((By.ID, "mobile-detailed-view")))
    print("✅ Login successful!")
except Exception as e:
    print(f"⚠️ Login error (or already logged in): {e}")

# --- Fill the form ---
wait.until(EC.presence_of_element_located((By.ID, "aCCOUNT_NAME"))).send_keys(form_data["name"])
driver.find_element(By.ID, "aCCOUNT_DESCRIPTION").send_keys(form_data["description"])

# Account Detail Type
acc_detail_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "ddlAccountDetailType_chosen")))
acc_detail_dropdown.click()
acc_detail_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[text()='{form_data['acc_detail']}']")))
driver.execute_script("arguments[0].scrollIntoView(true);", acc_detail_option)
time.sleep(0.5)
acc_detail_option.click()

# Group
group_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "ddlGroup_chosen")))
group_dropdown.click()
group_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[text()='{form_data['group']}']")))
driver.execute_script("arguments[0].scrollIntoView(true);", group_option)
time.sleep(0.5)
group_option.click()

# Debit/Credit
dc_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "ddlDC_chosen")))
dc_dropdown.click()
dc_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[text()='{form_data['dc']}']")))
driver.execute_script("arguments[0].scrollIntoView(true);", dc_option)
time.sleep(0.5)
dc_option.click()

# --- Bank fields ---
if form_data["acc_detail"] == "Bank":
    wait.until(EC.presence_of_element_located((By.ID, "bANK_ACCOUNT_NAME"))).send_keys(form_data["bank_name"])
    bank_type_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "ddlBankAccountType_chosen")))
    bank_type_dropdown.click()
    bank_type_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[text()='{form_data['bank_type']}']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", bank_type_option)
    time.sleep(0.5)
    bank_type_option.click()
    driver.find_element(By.ID, "bANK_ACCOUNT_ACCOUNTNO").send_keys(form_data["account_no"])
    driver.find_element(By.ID, "bANK_ACCOUNT_ROUTINGNUMBER").send_keys(form_data["routing_no"])
    time.sleep(2)

# --- Submit form ---
try:
    submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @value='Submit']")))
    submit_button.click()
    print("✅ Form submitted successfully.")
except Exception as e:
    print(f"⚠️ Submit error: {e}")

time.sleep(5)
driver.quit()