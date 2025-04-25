from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import tkinter as tk
from tkinter import messagebox
# --- Custom Tkinter Form ---
user_input = {}
def on_submit():
    name = name_entry.get().strip()
    desc = desc_entry.get().strip()
    days = days_entry.get().strip()
    if not name or not desc or not days:
        messagebox.showerror("Error", "All fields are required.")
        return
    user_input['name'] = name
    user_input['description'] = desc
    user_input['days'] = days
    form.destroy()
# Create popup form
form = tk.Tk()
form.title("Enter Payment Term Info")
tk.Label(form, text="Payment Term Name:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
name_entry = tk.Entry(form, width=40)
name_entry.grid(row=0, column=1, pady=5)
tk.Label(form, text="Description:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
desc_entry = tk.Entry(form, width=40)
desc_entry.grid(row=1, column=1, pady=5)
tk.Label(form, text="No. of Payment Days:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
days_entry = tk.Entry(form, width=40)
days_entry.grid(row=2, column=1, pady=5)
submit_button = tk.Button(form, text="Submit", command=on_submit)
submit_button.grid(row=3, column=0, columnspan=2, pady=10)
form.mainloop()
# If input was not filled, exit
if not user_input:
    exit()
# --- Initialize Edge browser ---
driver = webdriver.Edge()  # Ensure msedgedriver is in PATH
driver.maximize_window()
# Go to login page
driver.get("https://softwaredevelopmentsolution.com/Accounting/PaymentTerm/Create")
wait = WebDriverWait(driver, 20)
# Step 1: Login if required
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
# Step 2: Fill out the Payment Term form using user input
try:
    wait.until(EC.presence_of_element_located((By.ID, "pAYMENTTERM_NAME"))).send_keys(user_input['name'])
    print("✅ Entered Name")
    wait.until(EC.presence_of_element_located((By.ID, "pAYMENTTERM_DESCRIPTION"))).send_keys(user_input['description'])
    print("✅ Entered Description")
    wait.until(EC.presence_of_element_located((By.ID, "pAYMENTTERM_PAYMENTNOOFDAYS"))).send_keys(user_input['days'])
    print("✅ Entered Payment No. of Days")
except Exception as e:
    print(f"❌ Form input failed: {e}")
# Step 3: Select from Chosen Surcharge Dropdown
try:
    surcharge_container = wait.until(EC.presence_of_element_located((By.ID, "pAYMENTTERM_PAYMENTSURCHARGEID_chosen")))
    chosen_link = surcharge_container.find_element(By.CLASS_NAME, "chosen-single")
    driver.execute_script("arguments[0].click();", chosen_link)
    search_box = surcharge_container.find_element(By.XPATH, ".//input")
    search_box.send_keys("Account Payment")
    time.sleep(1.2)
    try:
        option = surcharge_container.find_element(By.XPATH, ".//li[contains(text(), 'Account Payment')]")
        driver.execute_script("arguments[0].click();", option)
        print("✅ Selected Surcharge from dropdown")
    except:
        search_box.send_keys(Keys.ENTER)
        print("✅ Selected Surcharge via ENTER (fallback)")
except Exception as e:
    print(f"❌ Failed to select surcharge: {e}")
# Step 4: Click Submit
try:
    submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Submit')]")))
    driver.execute_script("arguments[0].click();", submit_button)
    print("✅ Form submitted")
    messagebox.showinfo("Success", f"Form submitted with:\n\nName: {user_input['name']}\nDescription: {user_input['description']}\nDays: {user_input['days']}")
except Exception as e:
    print(f"❌ Submit failed: {e}")
    messagebox.showerror("Submission Error", f"Something went wrong:\n{e}")
driver.quit()