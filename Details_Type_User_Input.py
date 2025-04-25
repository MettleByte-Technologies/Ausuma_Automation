from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import simpledialog
import time

# ✅ Step 1: Unified popup form for Name, Description, Account Type
def get_user_input():
    user_data = {}

    def submit_data():
        user_data['name'] = name_entry.get()
        user_data['description'] = description_entry.get()
        user_data['account_type'] = account_type_entry.get()
        root.destroy()

    root = tk.Tk()
    root.title("Enter Account Detail Type Info")
    root.geometry("350x180")
    root.resizable(False, False)

    tk.Label(root, text="Name:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    name_entry = tk.Entry(root, width=30)
    name_entry.grid(row=0, column=1)

    tk.Label(root, text="Description:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    description_entry = tk.Entry(root, width=30)
    description_entry.grid(row=1, column=1)

    tk.Label(root, text="Account Type:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
    account_type_entry = tk.Entry(root, width=30)
    account_type_entry.grid(row=2, column=1)

    submit_btn = tk.Button(root, text="Submit", command=submit_data)
    submit_btn.grid(row=3, columnspan=2, pady=20)

    root.mainloop()
    return user_data['name'], user_data['description'], user_data['account_type']

# ✅ Step 2: Get input from user first
name, description, account_type_input = get_user_input()

# ✅ Step 3: Set up Edge
options = Options()
options.use_chromium = True
driver = webdriver.Edge(options=options)
driver.get("https://softwaredevelopmentsolution.com/Accounting/AccountDetailType/Create")
wait = WebDriverWait(driver, 15)

# ✅ Step 4: Log in if needed
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
    print(f"⚠️ Login not required or already logged in: {e}")

# ✅ Step 5: Fill the form
name_field = wait.until(EC.presence_of_element_located((By.ID, "aCCOUNTDETAILTYPE_NAME")))
name_field.clear()
name_field.send_keys(name)

desc_field = wait.until(EC.presence_of_element_located((By.ID, "aCCOUNTDETAILTYPE_DESCRIPTION")))
desc_field.clear()
desc_field.send_keys(description)

# ✅ Step 6: Handle Account Type dropdown (Chosen plugin dropdown)
chosen_dropdown = wait.until(EC.presence_of_element_located((By.ID, "ddlAccountType_chosen")))
chosen_dropdown.click()  # Click to open the dropdown

# Now, search for the correct option in the dropdown
options = driver.find_elements(By.XPATH, "//ul[@class='chosen-results']/li")
matched = False
for option in options:
    if account_type_input.lower() in option.text.lower():
        option.click()  # Select the matched option
        matched = True
        break

if not matched:
    print(f"⚠️ No matching Account Type found for: {account_type_input}")

# ✅ Step 7: Submit form
submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Submit']")))
submit_button.click()

print("✅ Form submitted successfully!")
time.sleep(3)
driver.quit()
