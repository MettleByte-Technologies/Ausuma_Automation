import time
import random
from datetime import datetime
from tkinter import Tk, Label, Entry, Button, StringVar, OptionMenu
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options

# === Generate unique values ===
def generate_unique_value(base, max_val=None):
    suffix = datetime.now().strftime("%H%M%S")
    if max_val:
        return str(random.randint(1, int(max_val)))
    return f"{base}_{suffix}"

# === Tkinter popup for input ===
def show_popup():
    def submit():
        popup_data["name"] = name_var.get()
        popup_data["grace_period"] = grace_var.get()
        popup_data["desc1"] = desc1_var.get()
        popup_data["type"] = type_var.get()
        popup_data["max_days"] = max_days_var.get()
        popup_data["value"] = value_var.get()
        popup_data["desc2"] = desc2_var.get()
        root.destroy()

    popup_data = {}
    root = Tk()
    root.title("Enter Payment Surcharge Details")

    # Dropdown values
    type_options = ["Flat", "Percentage", "Interest"]

    # Default values
    name_default = generate_unique_value("Name")
    grace_default = generate_unique_value("", 364)
    desc1_default = generate_unique_value("Description1")
    max_days_default = generate_unique_value("", 364)
    value_default = generate_unique_value("", 99)
    desc2_default = generate_unique_value("Description2")

    # Variables
    name_var = StringVar(value=name_default)
    grace_var = StringVar(value=grace_default)
    desc1_var = StringVar(value=desc1_default)
    type_var = StringVar(value=type_options[0])
    max_days_var = StringVar(value=max_days_default)
    value_var = StringVar(value=value_default)
    desc2_var = StringVar(value=desc2_default)

    # Layout
    Label(root, text="Name:").grid(row=0, column=0)
    Entry(root, textvariable=name_var).grid(row=0, column=1)

    Label(root, text="Grace Period (Days):").grid(row=1, column=0)
    Entry(root, textvariable=grace_var).grid(row=1, column=1)

    Label(root, text="Description 1:").grid(row=2, column=0)
    Entry(root, textvariable=desc1_var).grid(row=2, column=1)

    Label(root, text="Type:").grid(row=3, column=0)
    OptionMenu(root, type_var, *type_options).grid(row=3, column=1)

    Label(root, text="Max Days:").grid(row=4, column=0)
    Entry(root, textvariable=max_days_var).grid(row=4, column=1)

    Label(root, text="Value:").grid(row=5, column=0)
    Entry(root, textvariable=value_var).grid(row=5, column=1)

    Label(root, text="Description 2:").grid(row=6, column=0)
    Entry(root, textvariable=desc2_var).grid(row=6, column=1)

    Button(root, text="Submit", command=submit).grid(row=7, column=0, columnspan=2)

    root.mainloop()
    return popup_data

# === Step 1: Get data from popup ===
user_input = show_popup()

# === Step 2: Launch Edge and run automation ===
edge_options = Options()
edge_options.use_chromium = True
driver = webdriver.Edge(options=edge_options)
driver.maximize_window()
wait = WebDriverWait(driver, 20)

# === Step 3: Open Target URL ===
driver.get("https://softwaredevelopmentsolution.com/Accounting/PaymentSurcharge/Create")

# === Step 4: Login ===
try:
    username_field = wait.until(EC.presence_of_element_located((By.ID, "Email")))
    password_field = wait.until(EC.presence_of_element_located((By.ID, "Password")))
    login_button = wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit")))

    username_field.send_keys("ola123@yopmail.com")
    password_field.send_keys("1")
    login_button.click()

    # Wait for redirect to dashboard or form
    wait.until(EC.presence_of_element_located((By.ID, "mobile-detailed-view")))
    print("✅ Login successful!")
except Exception as e:
    print(f"⚠️ Login error (or already logged in): {e}")

# === Step 5: Fill Payment Surcharge Form ===
try:
    # Wait for form to appear
    wait.until(EC.presence_of_element_located((By.ID, "pAYMENTSURCHARGE_NAME"))).send_keys(user_input["name"])
    driver.find_element(By.ID, "pAYMENTSURCHARGE_GRACEPERIODDAYS").send_keys(user_input["grace_period"])
    driver.find_element(By.ID, "pAYMENTSURCHARGE_DESCRIPTION").send_keys(user_input["desc1"])

    # Handle dropdown
    driver.find_element(By.ID, "txtItemType_chosen").click()
    time.sleep(1)  # Allow options to load
    options = driver.find_elements(By.CSS_SELECTOR, "#txtItemType_chosen .chosen-drop ul.chosen-results li")
    for opt in options:
        if opt.text.strip() == user_input["type"]:
            opt.click()
            break

    # Fill rest of the form
    driver.find_element(By.ID, "txtGridMaxNoofDays").send_keys(user_input["max_days"])
    driver.find_element(By.ID, "txtGridValue").send_keys(user_input["value"])
    driver.find_element(By.ID, "description").send_keys(user_input["desc2"])

    # Submit the form
    submit_button = driver.find_element(By.ID, "submitbtn")
    submit_button.click()
    print("✅ Form submitted successfully!")

except Exception as e:
    print(f"⚠️ Error while filling or submitting the form: {e}")
