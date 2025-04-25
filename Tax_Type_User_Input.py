from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from tkinter import Tk, Label, Button, Entry

# --- Get user input using Tkinter ---
def get_user_input():
    def submit():
        nonlocal user_data
        user_data = {
            'account1': account1_entry.get(),
            'account2': account2_entry.get(),
        }
        root.destroy()

    user_data = {}
    root = Tk()
    root.title("User Input")

    Label(root, text="Account 1:").grid(row=0, column=0)
    account1_entry = Entry(root)
    account1_entry.grid(row=0, column=1)

    Label(root, text="Account 2:").grid(row=1, column=0)
    account2_entry = Entry(root)
    account2_entry.grid(row=1, column=1)

    Button(root, text="Submit", command=submit).grid(row=2, column=1)

    root.mainloop()
    return user_data

# --- Start Selenium automation ---
user_data = get_user_input()

# Setup Microsoft Edge WebDriver
driver = webdriver.Edge()

# Open Journal Entry creation screen
driver.get("https://softwaredevelopmentsolution.com/Accounting/JournalEntry/Create")

wait = WebDriverWait(driver, 20)

# Wait for the first Account dropdown to appear
account1_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[1]//td[1]//input[contains(@class, 'account-select')]")))
account1_input.click()
account1_input.send_keys(user_data['account1'])
wait.until(EC.visibility_of_element_located((By.XPATH, f"//li[contains(text(), '{user_data['account1']}')]"))).click()

# Fill in Amount
amount_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[1]//td[2]//input")))
amount_input.send_keys("1000")

# Select Credit or Debit from dropdown
type_dropdown = Select(wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[1]//td[3]//select"))))
type_dropdown.select_by_visible_text("Debit")

# Fill Description and press Tab to generate second row
desc_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[1]//td[4]//input")))
desc_input.send_keys("Auto-entry")
desc_input.send_keys(Keys.TAB)

# Second row: wait until it appears
account2_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[2]//td[1]//input[contains(@class, 'account-select')]")))
account2_input.click()
account2_input.send_keys(user_data['account2'])
wait.until(EC.visibility_of_element_located((By.XPATH, f"//li[contains(text(), '{user_data['account2']}')]"))).click()
amount2_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[2]//td[2]//input")))
amount2_input.send_keys("1000")
type2_dropdown = Select(wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[2]//td[3]//select"))))
type2_dropdown.select_by_visible_text("Credit")
driver.quit()
