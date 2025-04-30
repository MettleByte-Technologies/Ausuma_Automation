import time
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Step 1: Tkinter Popup for User Input (Single Page)
def get_user_input():
    # Initialize Tkinter root window
    root = tk.Tk()
    root.title("Report Generation Inputs")

    # Labels and Entries for From Date and To Date
    tk.Label(root, text="Enter From Date (MM/DD/YYYY):").grid(row=0, column=0, padx=10, pady=10)
    from_date = tk.Entry(root)
    from_date.grid(row=0, column=1, padx=10, pady=10)

    tk.Label(root, text="Enter To Date (MM/DD/YYYY):").grid(row=1, column=0, padx=10, pady=10)
    to_date = tk.Entry(root)
    to_date.grid(row=1, column=1, padx=10, pady=10)

    # Dropdown for Export Option
    tk.Label(root, text="Select Export Option:").grid(row=2, column=0, padx=10, pady=10)
    export_option = ttk.Combobox(root, values=["Pdf", "Excel", "Print"])
    export_option.grid(row=2, column=1, padx=10, pady=10)
    export_option.set("Pdf")  # Set default to Pdf

    # Function to close and get the inputs when the button is clicked
    def submit():
        root.quit()

    # Submit button
    tk.Button(root, text="Submit", command=submit).grid(row=3, column=0, columnspan=2, pady=10)

    # Start the Tkinter loop
    root.mainloop()

    # Get values from the input fields
    return from_date.get(), to_date.get(), export_option.get()

# Step 2: Selenium Automation
def automate_report(from_date, to_date, export_option):
    # Start the Edge WebDriver
    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 20)
    driver.maximize_window()
    driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/JournalEntryList")

    # Step 2.1: Login
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
    print("✅ Logged in successfully")

    # Step 2.2: Fill the From Date and To Date fields
    wait.until(EC.presence_of_element_located((By.ID, "FromDate"))).clear()
    wait.until(EC.presence_of_element_located((By.ID, "FromDate"))).send_keys(from_date)
    wait.until(EC.presence_of_element_located((By.ID, "ToDate"))).clear()
    wait.until(EC.presence_of_element_located((By.ID, "ToDate"))).send_keys(to_date)

    # Step 2.3: Click the Generate Report button
    wait.until(EC.element_to_be_clickable((By.ID, "submitbtn"))).click()
    print("✅ Report generated successfully")
    time.sleep(5)
    # Step 2.4: Wait for the processing div to disappear (blocking element)
    wait.until(EC.invisibility_of_element_located((By.ID, "processing-scroll")))

    # Step 2.5: Ensure the Export button is clickable and then click it
    wait.until(EC.element_to_be_clickable((By.ID, "exportbtn"))).click()
    time.sleep(1)  # Wait for the dropdown to open

    if export_option == "Pdf":
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Pdf')]"))).click()
        print("✅ Export to PDF initiated")
    elif export_option == "Excel":
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Excel')]"))).click()
        print("✅ Export to Excel initiated")
    elif export_option == "Print":
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Print')]"))).click()
        print("✅ Print initiated")

    # Step 3: Close the browser after operation
    time.sleep(5)  # Wait a few seconds to see the results
    driver.quit()

# Main function to execute the script
if __name__ == "__main__":
    from_date, to_date, export_option = get_user_input()  # Get input from the user
    if from_date and to_date:  # Proceed if valid input is provided
        automate_report(from_date, to_date, export_option)
    else:
        print("❌ Invalid input. Please provide valid From Date and To Date.")
