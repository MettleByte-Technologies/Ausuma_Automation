import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random
from datetime import datetime

# === AUTO-GENERATE UNIQUE NAME AND DESCRIPTION ===
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
default_name = f"MiscItem_{timestamp}"
default_description = f"Test description for miscellaneous item {timestamp}"

# === TKINTER POPUP ===
user_input = {}

def get_user_input():
    def submit():
        user_input['name'] = name_entry.get()
        user_input['description'] = description_entry.get()
        user_input['account'] = account_var.get()
        root.destroy()

    root = tk.Tk()
    root.title("Miscellaneous Item Entry")
    root.geometry("300x300")

    tk.Label(root, text="Name:").pack()
    name_entry = tk.Entry(root)
    name_entry.insert(0, default_name)
    name_entry.pack()

    tk.Label(root, text="Description:").pack()
    description_entry = tk.Entry(root)
    description_entry.insert(0, default_description)
    description_entry.pack()

    tk.Label(root, text="Account:").pack()
    account_var = tk.StringVar()
    account_dropdown = ttk.Combobox(root, textvariable=account_var)
    account_dropdown['values'] = [
        "Bank of America", "Bank of Ausuma", "Commissions & fees", "Depreciation", 
        "Development Income", "Discount Given - COS", "Discounts received", "Fuel Expenses", 
        "Gas Bills", "Acc. Account Date", "Acc. Dep. Furniture & Fixtures", 
        "Acc. Dep. Land & Buildings", "Acc. Dep. Motor Vehicles"
    ]
    account_dropdown.current(0)
    account_dropdown.pack()

    tk.Button(root, text="Submit", command=submit).pack(pady=10)

    root.mainloop()

# Call popup first to get user input
get_user_input()

# === SELENIUM SCRIPT ===
driver = webdriver.Edge()
driver.maximize_window()
wait = WebDriverWait(driver, 20)

# Login
driver.get("https://softwaredevelopmentsolution.com/Accounting/MiscellaneousItem/Create")

try:
    username_field = wait.until(EC.presence_of_element_located((By.ID, "Email")))
    password_field = wait.until(EC.presence_of_element_located((By.ID, "Password")))
    login_button = wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit")))

    username_field.send_keys("ola123@yopmail.com")
    password_field.send_keys("1")
    login_button.click()

    # Wait for redirect
    wait.until(EC.presence_of_element_located((By.ID, "mobile-detailed-view")))
    print("✅ Login successful!")

except Exception as e:
    print(f"⚠️ Login error or already logged in: {e}")

# Load the Create Miscellaneous Item form


# Fill in Name
name_field = wait.until(EC.presence_of_element_located((By.ID, "mISCELLANEOU_NAME")))
name_value = user_input['name']  # Use the user-provided name
name_field.send_keys(name_value)

# Fill in Description
description_field = driver.find_element(By.ID, "mISCELLANEOU_DESCRIPTION")
description_value = user_input['description']  # Use the user-provided description
description_field.send_keys(description_value)

# Handle Chosen dropdown for Account
account_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "ddlAccount_chosen")))
account_dropdown.click()

# Search inside the dropdown
search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#ddlAccount_chosen .chosen-search input")))
search_input.send_keys(user_input['account'])
time.sleep(1)  # Allow dropdown to filter

# Click on the desired item
options = driver.find_elements(By.CSS_SELECTOR, "#ddlAccount_chosen .chosen-results li")
for option in options:
    if user_input['account'] in option.text:
        option.click()
        break

# Submit form with updated waiting method
try:
    # Use XPath if ID doesn't work
    submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Submit')]")))
    submit_button.click()
    print("✅ Form submitted successfully!")
except Exception as e:
    print(f"⚠️ Error submitting form: {e}")

time.sleep(5)
driver.quit()
