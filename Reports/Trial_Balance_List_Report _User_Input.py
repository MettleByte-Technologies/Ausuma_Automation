import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time

# Function to run the script with user input
def run_script(account_class, account_type, account_detail_type, account_group, export_option,
               exclude_zero_balance, from_date, to_date, hide_class, hide_type, hide_detail_type, enable_currency):
    driver = webdriver.Edge()  # Setup Edge WebDriver
    driver.maximize_window()
    driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/TrialBalance")

    wait = WebDriverWait(driver, 20)

    # Login
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
    print("✅ Logged in successfully")

    # Helper function to select dropdowns
    def select_chosen_dropdown(dropdown_id, value):
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "preloader-it")))
        dropdown = wait.until(EC.element_to_be_clickable((By.ID, dropdown_id)))
        dropdown.click()
        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "chosen-results")))
        options = driver.find_elements(By.XPATH, f"//ul[@class='chosen-results']//li[text()='{value}']")
        if options:
            ActionChains(driver).move_to_element(options[0]).click().perform()
            print(f"✅ {dropdown_id} selected: {value}")
        else:
            print(f"❌ Option '{value}' not found in {dropdown_id}.")
        driver.find_element(By.TAG_NAME, 'body').click()

    # Select dropdowns
    select_chosen_dropdown("ddlAccountClass_chosen", account_class)
    select_chosen_dropdown("ddlAccountType_chosen", account_type)
    select_chosen_dropdown("ddlAccountDetailType_chosen", account_detail_type)
    select_chosen_dropdown("ddlAccountGroup_chosen", account_group)

    # Set checkboxes
    if exclude_zero_balance:
        zero_balance_checkbox = wait.until(EC.element_to_be_clickable((By.ID, "ZeroBalanceAccount")))
        if not zero_balance_checkbox.is_selected():
            zero_balance_checkbox.click()
            print("✅ Exclude Zero Balance checked.")
    
    if hide_class:
        hide_class_checkbox = wait.until(EC.element_to_be_clickable((By.ID, "classGrouping")))
        if not hide_class_checkbox.is_selected():
            hide_class_checkbox.click()
            print("✅ Hide Class checked.")

    if hide_type:
        hide_type_checkbox = wait.until(EC.element_to_be_clickable((By.ID, "typeGrouping")))
        if not hide_type_checkbox.is_selected():
            hide_type_checkbox.click()
            print("✅ Hide Type checked.")

    if hide_detail_type:
        hide_detail_type_checkbox = wait.until(EC.element_to_be_clickable((By.ID, "dtypeGrouping")))
        if not hide_detail_type_checkbox.is_selected():
            hide_detail_type_checkbox.click()
            print("✅ Hide Detail Type checked.")

    if enable_currency:
        enable_currency_checkbox = wait.until(EC.element_to_be_clickable((By.ID, "EnableCurrency")))
        if not enable_currency_checkbox.is_selected():
            enable_currency_checkbox.click()
            print("✅ Enable Currency checked.")

    # Fill From Date and To Date
    if from_date:
        from_date_input = wait.until(EC.presence_of_element_located((By.ID, "FromDate")))
        from_date_input.clear()
        from_date_input.send_keys(from_date)
        print(f"📅 From Date set to {from_date}")

    if to_date:
        to_date_input = wait.until(EC.presence_of_element_located((By.ID, "ToDate")))
        to_date_input.clear()
        to_date_input.send_keys(to_date)
        print(f"📅 To Date set to {to_date}")

    # Click Generate Report
    generate_btn = wait.until(EC.element_to_be_clickable((By.ID, "submitbtn")))
    driver.execute_script("arguments[0].scrollIntoView(true);", generate_btn)
    time.sleep(2)
    driver.execute_script("arguments[0].click();", generate_btn)
    print("📄 Generate Report button clicked.")
    time.sleep(15)

    # Export report
    export_btn = wait.until(EC.element_to_be_clickable((By.ID, "exportbtn")))
    driver.execute_script("arguments[0].scrollIntoView(true);", export_btn)
    export_btn.click()
    print("✅ Export dropdown opened.")

    try:
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
    except Exception as e:
        print(f"❌ Error with export option: {e}")

    time.sleep(10)
    print("✅ Report export process completed.")

# Function to open the popup and get user input
def open_popup():
    def on_button_click():
        selected_class = class_var.get()
        selected_type = type_var.get()
        selected_detail_type = detail_type_var.get()
        selected_group = group_var.get()
        export_option = export_var.get()
        exclude_zero_balance = exclude_zero_var.get()
        from_date = from_date_entry.get()
        to_date = to_date_entry.get()
        hide_class = hide_class_var.get()
        hide_type = hide_type_var.get()
        hide_detail_type = hide_detail_type_var.get()
        enable_currency = enable_currency_var.get()

        if selected_class and selected_type and selected_detail_type and selected_group:
            run_script(selected_class, selected_type, selected_detail_type, selected_group, export_option,
                       exclude_zero_balance, from_date, to_date, hide_class, hide_type, hide_detail_type, enable_currency)
            root.quit()
        else:
            messagebox.showerror("Error", "Please select Account Class, Account Type, Account Detail Type, and Account Group.")

    root = tk.Tk()
    root.title("Select Options")

    # Account Class
    tk.Label(root, text="Select Account Class:").pack(padx=10, pady=5)
    class_options = ["All", "Asset", "Expense", "Liability", "Equity", "Revenue"]
    class_var = tk.StringVar(root)
    class_var.set(class_options[0])
    tk.OptionMenu(root, class_var, *class_options).pack(padx=10, pady=5)

    # Account Type
    tk.Label(root, text="Select Account Type:").pack(padx=10, pady=5)
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
    type_var.set(type_options[0])
    tk.OptionMenu(root, type_var, *type_options).pack(padx=10, pady=5)

    # Account Detail Type
    tk.Label(root, text="Select Account Detail Type:").pack(padx=10, pady=5)
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
    detail_type_var.set(detail_type_options[0])
    tk.OptionMenu(root, detail_type_var, *detail_type_options).pack(padx=10, pady=5)

    # Account Group
    tk.Label(root, text="Select Account Group:").pack(padx=10, pady=5)
    group_options = [
        "All", "Expenses", "Purchase Expense", "Liability", "current", "saving", "Show", "Cost of Goods Sold", "Financial Groups",
        "Financial Groups1", "Acc Group", "Cost of Goods Sold Group", "Test Edit Groups", "TEST GROUP AUSUMA", "GST GROUP 03-02-2025"
    ]
    group_var = tk.StringVar(root)
    group_var.set(group_options[0])
    tk.OptionMenu(root, group_var, *group_options).pack(padx=10, pady=5)

    # New Fields after Account Group
    exclude_zero_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Exclude Zero Balance", variable=exclude_zero_var).pack(padx=10, pady=5)

    tk.Label(root, text="From Date (DD-MM-YYYY):").pack(padx=10, pady=5)
    from_date_entry = tk.Entry(root)
    from_date_entry.pack(padx=10, pady=5)

    tk.Label(root, text="To Date (DD-MM-YYYY):").pack(padx=10, pady=5)
    to_date_entry = tk.Entry(root)
    to_date_entry.pack(padx=10, pady=5)

    hide_class_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Hide Class", variable=hide_class_var).pack(padx=10, pady=5)

    hide_type_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Hide Type", variable=hide_type_var).pack(padx=10, pady=5)

    hide_detail_type_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Hide Detail Type", variable=hide_detail_type_var).pack(padx=10, pady=5)

    enable_currency_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Enable Currency", variable=enable_currency_var).pack(padx=10, pady=5)

    # Export Option
    tk.Label(root, text="Select Export Option:").pack(padx=10, pady=5)
    export_options = ["Pdf", "Excel", "Print"]
    export_var = tk.StringVar(root)
    export_var.set(export_options[0])
    tk.OptionMenu(root, export_var, *export_options).pack(padx=10, pady=5)

    tk.Button(root, text="Run Script", command=on_button_click).pack(pady=10)

    root.mainloop()

# Run the popup
open_popup()
