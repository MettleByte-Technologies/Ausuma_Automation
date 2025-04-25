from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk

# ✅ Step 1: Global variables to hold form data
user_input = {"name": "", "description": ""}

# ✅ Step 2: Tkinter form for input
def get_input_from_user():
    def submit():
        name = name_entry.get().strip()
        description = desc_entry.get().strip()

        if not name:
            tk.messagebox.showerror("Error", "Name is required!")
            return

        user_input["name"] = name
        user_input["description"] = description
        root.destroy()

    root = tk.Tk()
    root.title("Enter Account Details")
    root.geometry("350x200")
    root.resizable(False, False)

    tk.Label(root, text="Name:").pack(pady=(10, 0))
    name_entry = tk.Entry(root, width=40)
    name_entry.pack()

    tk.Label(root, text="Description:").pack(pady=(10, 0))
    desc_entry = tk.Entry(root, width=40)
    desc_entry.pack()

    submit_btn = tk.Button(root, text="Submit", command=submit)
    submit_btn.pack(pady=20)

    root.mainloop()

# ✅ Step 3: Run Tkinter popup and wait for input
get_input_from_user()
name = user_input["name"]
description = user_input["description"]

# ✅ Step 4: Selenium setup
options = Options()
options.use_chromium = True
driver = webdriver.Edge(options=options)
driver.get("https://softwaredevelopmentsolution.com/Accounting/Group/Create")
wait = WebDriverWait(driver, 15)

# ✅ Step 5: Login if needed
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

# ✅ Step 6: Fill and submit the form
wait.until(EC.presence_of_element_located((By.ID, "aCCOUNTGROUP_NAME"))).send_keys(name)
wait.until(EC.presence_of_element_located((By.ID, "aCCOUNTGROUP_DESCRIPTION"))).send_keys(description)

submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @value='Submit']")))
submit_button.click()

print("✅ Form submitted successfully!")
driver.quit()