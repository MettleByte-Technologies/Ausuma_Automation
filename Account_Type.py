from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import Tk, Label, Entry, Button, StringVar, OptionMenu
import time
# ----- Popup for User Input -----
def get_user_input():
    def submit():
        nonlocal name, description, selected_class
        name = name_var.get()
        description = desc_var.get()
        selected_class = class_var.get()
        root.destroy()
    name = description = selected_class = ""
    root = Tk()
    root.title("Enter Account Type Info")
    root.geometry("350x200")
    Label(root, text="Name:").pack()
    name_var = StringVar()
    Entry(root, textvariable=name_var).pack()
    Label(root, text="Description:").pack()
    desc_var = StringVar()
    Entry(root, textvariable=desc_var).pack()
    Label(root, text="Class:").pack()
    class_var = StringVar(value="Asset")
    options = ["Asset", "Liability", "Equity", "Income", "Expense"]
    OptionMenu(root, class_var, *options).pack()
    Button(root, text="Submit", command=submit).pack(pady=10)
    root.mainloop()
    return name, description, selected_class
# ----- Get Input from User -----
account_name, account_description, class_input = get_user_input()
# ----- Start Edge WebDriver -----
options = webdriver.EdgeOptions()
options.add_experimental_option("detach", True)  # Keeps browser open after script ends
driver = webdriver.Edge(options=options)
wait = WebDriverWait(driver, 10)
driver.get("https://softwaredevelopmentsolution.com/Accounting/AccountType/Create")
# ----- Login -----
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
# ----- Fill Form -----
try:
    name_field = wait.until(EC.presence_of_element_located((By.ID, "aCCOUNTTYPE_NAME")))
    desc_field = wait.until(EC.presence_of_element_located((By.ID, "aCCOUNTTYPE_DESCRIPTION")))
    name_field.clear()
    name_field.send_keys(account_name)
    desc_field.clear()
    desc_field.send_keys(account_description)
    print("✅ Filled Name and Description")
except Exception as e:
    print(f"❌ Failed to fill fields: {e}")
# ---- Select Class Radio Button -----
try:
    # Mapping class input to the value in the radio button
    class_map = {
        "Asset": "A",
        "Liability": "L",
        "Equity": "Q",
        "Income": "R",
        "Expense": "E"
    }
    value_to_select = class_map.get(class_input)
    if value_to_select is None:
        raise Exception(f"Class '{class_input}' not found in map.")
    # Select the correct radio button based on the value
    rb = driver.find_element(By.XPATH, f"//input[@name='aCCOUNTTYPE.CLASS' and @value='{value_to_select}']")
    driver.execute_script("arguments[0].click();", rb)
    print(f"✅ Selected class radio button (value={value_to_select})")
except Exception as e:
    print(f"❌ Failed to select class radio button: {e}")
# ----- Submit Form -----
try:
    submit_btn = driver.find_element(By.XPATH, "//input[@type='button' and @value='Submit']")
    driver.execute_script("arguments[0].click();", submit_btn)
    print("✅ Form submitted")
except Exception as e:
    print(f"❌ Submit failed: {e}")
# ----- Optionally keep browser open -----
time.sleep(5)
