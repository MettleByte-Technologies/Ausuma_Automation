import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time

# Function to run the script with user input
def run_script(account_class, account_type, account_detail_type, account_group, zero_balance, enable_currency, export_option):
    # Setup Edge driver (make sure msedgedriver is in PATH or specify executable_path)
    driver = webdriver.Edge()  # Or use executable_path='path_to_msedgedriver' if necessary

    # Maximize the window
    driver.maximize_window()

    # Open the target URL
    driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/GeneralLedger")

    # Wait until page loads
    wait = WebDriverWait(driver, 20)

    # Login steps
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
    print("✅ Logged in successfully")

    # Function to select from Chosen dropdowns
    def select_chosen_dropdown(dropdown_id, value):
        # Wait for the preloader to disappear (if present)
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "preloader-it")))

        # Wait until the dropdown is visible and click to open it
        dropdown = wait.until(EC.element_to_be_clickable((By.ID, dropdown_id)))
        dropdown.click()  # Open dropdown

        # Wait until the options list appears and is clickable
        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "chosen-results")))

        # Search for the option matching the user input
        options = driver.find_elements(By.XPATH, f"//ul[@class='chosen-results']//li[text()='{value}']")

        if options:
            option = options[0]
            # Use ActionChains to click on the option, ensuring no obstruction by other elements
            actions = ActionChains(driver)
            actions.move_to_element(option).click().perform()
            print(f"✅ {dropdown_id} selected: {value}")
        else:
            print(f"❌ Option '{value}' not found in {dropdown_id}. Please check the input.")

        # Close the dropdown by clicking outside (click on the body)
        driver.find_element(By.TAG_NAME, 'body').click()

    # Select Account Class dropdown
    select_chosen_dropdown("ddlAccountClass_chosen", account_class)

    # Select Account Type dropdown
    select_chosen_dropdown("ddlAccountType_chosen", account_type)

    # Select Account Detail Type dropdown
    select_chosen_dropdown("ddlAccountDetailType_chosen", account_detail_type)

    # Select Account Group dropdown
    select_chosen_dropdown("ddlAccountGroup_chosen", account_group)

    # Check the "Exclude Zero Balance" checkbox if selected
    if zero_balance:
        zero_balance_checkbox = wait.until(EC.element_to_be_clickable((By.ID, "ZeroBalanceAccount")))
        zero_balance_checkbox.click()
        print("✅ Zero Balance checkbox selected.")

    # Check the "Enable Currency" checkbox if selected
    if enable_currency:
        enable_currency_checkbox = wait.until(EC.element_to_be_clickable((By.ID, "EnableCurrency")))
        enable_currency_checkbox.click()
        print("✅ Enable Currency checkbox selected.")

    # Wait for the "Generate Report" button to be clickable
    generate_btn = wait.until(EC.element_to_be_clickable((By.ID, "submitbtn")))

    # Scroll the "Generate Report" button into view (ensure it's visible)
    driver.execute_script("arguments[0].scrollIntoView(true);", generate_btn)

    # Wait for a moment to make sure the button is fully visible
    time.sleep(2)

    # Attempt to click the "Generate Report" button using JavaScript to bypass potential issues
    try:
        driver.execute_script("arguments[0].click();", generate_btn)
        print("📄 Generate Report button clicked.")
    except Exception as e:
        print(f"❌ Error clicking the button: {e}")

    # Wait for the report to generate or page to change (Adjust the wait time as needed)
    time.sleep(15)  # Increase or decrease the time as needed based on your report generation time

    # Now click the Export button to export the report in the desired format (PDF, Excel, or Print)
    export_btn = wait.until(EC.element_to_be_clickable((By.ID, "exportbtn")))

    # Scroll the "Export" button into view (ensure it's visible)
    driver.execute_script("arguments[0].scrollIntoView(true);", export_btn)

    # Click the Export button to open the dropdown
    export_btn.click()
    print("✅ Export dropdown opened.")

    # Wait for the export options to be clickable and select the user's choice (Pdf, Excel, Print)
    if export_option == 'Pdf':
        pdf_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//ul[@class='dropdown-menu']//a[text()='Pdf']")))
        pdf_option.click()
        print("✅ Exporting as PDF.")
    elif export_option == 'Excel':
        excel_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//ul[@class='dropdown-menu']//a[text()='Excel']")))
        excel_option.click()
        print("✅ Exporting as Excel.")
    elif export_option == 'Print':
        print_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//ul[@class='dropdown-menu']//a[text()='Print']")))
        print_option.click()
        print("✅ Printing report.")

    # Wait for the report to export or print (Adjust the wait time as needed)
    time.sleep(10)  # Adjust this wait time if necessary

    print("✅ Report export process completed.")

    # Optionally: close the browser after the process
    # driver.quit()

# Function to open the popup and get user input
def open_popup():
    def on_button_click():
        selected_class = class_var.get()
        selected_type = type_var.get()
        selected_detail_type = detail_type_var.get()
        selected_group = group_var.get()
        zero_balance = zero_balance_var.get()
        enable_currency = enable_currency_var.get()
        export_option = export_var.get()  # Get the selected export option
        
        if selected_class and selected_type and selected_detail_type and selected_group:
            run_script(selected_class, selected_type, selected_detail_type, selected_group, zero_balance, enable_currency, export_option)  # Run the Selenium script with the selected values
            root.quit()  # Close the popup after the script runs
        else:
            messagebox.showerror("Error", "Please select Account Class, Account Type, Account Detail Type, and Account Group.")

    # Create the main window (popup)
    root = tk.Tk()
    root.title("Select Options")

    # Account Class Dropdown
    class_label = tk.Label(root, text="Select Account Class:")
    class_label.pack(padx=10, pady=5)
    class_options = ["All", "Asset", "Expense", "Liability", "Equity", "Revenue"]
    class_var = tk.StringVar(root)
    class_var.set(class_options[0])  # Set default value
    class_dropdown = tk.OptionMenu(root, class_var, *class_options)
    class_dropdown.pack(padx=10, pady=5)

    # Account Type Dropdown
    type_label = tk.Label(root, text="Select Account Type:")
    type_label.pack(padx=10, pady=5)
    type_options = [
        "All", "Intangible Assets", "Current Assets", "Fixed Assets", "Other Assets", "Current Liabilities",
        "Long Term Liabilities", "Owners Equity", "Sales Income", "Other Income", "Direct Cost",
        "Admin Expenses", "Depr. & Amortisation Expenses", "Taxes", "Utilities", "Other Expenses",
        "Balance Sheet", "Asset Accounts", "Income Statement", "ACC. Account Edited*", "ACC. Asset",
        "Income Statement of Sales", "Account test date", "New Edit Account", "Income", "ITGST 03-02-2025",
        "TEST ACCOUNT TYPE", "TEST ACCOUNT TYPEc", "TEST AT 05-03-2025", "Mango", "Apple Digital Production",
        "test", "INCOME 123", "A TEST AT", "ABC TEST AT"
    ]
    type_var = tk.StringVar(root)
    type_var.set(type_options[0])  # Set default value
    type_dropdown = tk.OptionMenu(root, type_var, *type_options)
    type_dropdown.pack(padx=10, pady=5)

    # Account Detail Type Dropdown
    detail_type_label = tk.Label(root, text="Select Account Detail Type:")
    detail_type_label.pack(padx=10, pady=5)
    detail_type_options = [
        "All", "Goodwill", "Cash", "Inventories", "Accounts Receivables", "Advance & PrePayments", "Other Current Asset",
        "Land & Buildings", "Furniture & Fixtures", "Office Equipment & Computers", "Motor Vehicles", "Tools & Machinery",
        "Other Fixed Assets", "Miscellaneous Assets", "Accounts Payables", "Accruals", "Long Term Liabilities",
        "Opening Balance Equity", "Retained Earnings", "Sales Revenue Trading", "Services Revenue Trading", "Other Primary Income",
        "Miscellaneous Income", "Interest Earned", "Non Profit Income", "Cost Of Sales Trading", "Cost Of Sales Services",
        "Bank Charges", "Vehicle Expenses", "General & Administrative Expenses", "Legal Professional Fees", "Rent/Lease Of Buildings",
        "Depreciation Expenses", "Taxes Paid", "Utility Bills", "Discounts & Refunds", "Other Business Expenses",
        "Promotion & Advertising", "Meals & Entertainment", "Insurance", "Miscellaneous Expense", "Bank", "Payment Gateway Charges",
        "Asset Accounts", "Financial Accounting", "Financial Accounting1", "Financial Year", "Acc. Test", "Acc.Financial Accounting",
        "Financial Accounting of Sales", "Account Test Date", "Edit Details Type", "Financial Accounting000",
        "TEST DETAIL TYPE 24-01-2025", "GST DETAILS 03-02-2025", "TEST DETAILS TYPE", "Expenses", "Aladin", "001", "TEST DETAIL TYPE",
        "Mango", "TEST DT 05-03-2025", "002", "Apple Digital Production", "FINANCIAL INCOME", "TEST DT"
    ]
    detail_type_var = tk.StringVar(root)
    detail_type_var.set(detail_type_options[0])  # Set default value
    detail_type_dropdown = tk.OptionMenu(root, detail_type_var, *detail_type_options)
    detail_type_dropdown.pack(padx=10, pady=5)

    # Account Group Dropdown
    group_label = tk.Label(root, text="Select Account Group:")
    group_label.pack(padx=10, pady=5)
    group_options = [
        "All", "Expenses", "Purchase Expense", "Liability", "current", "saving", "Show", "Cost of Goods Sold", "Financial Groups",
        "Financial Groups1", "Acc Group", "Cost of Goods Sold Group", "Test Edit Groups", "TEST GROUP AUSUMA", "GST GROUP 03-02-2025"
    ]
    group_var = tk.StringVar(root)
    group_var.set(group_options[0])  # Set default value
    group_dropdown = tk.OptionMenu(root, group_var, *group_options)
    group_dropdown.pack(padx=10, pady=5)

    # Exclude Zero Balance Checkbox
    zero_balance_var = tk.BooleanVar()
    zero_balance_checkbox = tk.Checkbutton(root, text="Exclude Zero Balance", variable=zero_balance_var)
    zero_balance_checkbox.pack(padx=10, pady=5)

    # Enable Currency Checkbox
    enable_currency_var = tk.BooleanVar()
    enable_currency_checkbox = tk.Checkbutton(root, text="Enable Currency", variable=enable_currency_var)
    enable_currency_checkbox.pack(padx=10, pady=5)

    # Export Option Dropdown
    export_label = tk.Label(root, text="Select Export Option:")
    export_label.pack(padx=10, pady=5)
    export_options = ["Pdf", "Excel", "Print"]
    export_var = tk.StringVar(root)
    export_var.set(export_options[0])  # Set default value
    export_dropdown = tk.OptionMenu(root, export_var, *export_options)
    export_dropdown.pack(padx=10, pady=5)

    # Submit Button
    submit_btn = tk.Button(root, text="Run Script", command=on_button_click)
    submit_btn.pack(pady=10)

    # Run the main loop
    root.mainloop()

# Call the function to open the popup
open_popup()
