import time
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Dropdown options
customer_vendor_options = [
    "All",
    "Robinson EV Products Company | Liam White | 9685471325 | liam@yopmail.com",
    "Robinson EV Products Company | Jems Turnner | 9658745698 | jems@yopmail.com",
    "Robinson EV Products Company | William Turner | 96558458869 | william@yopmail.com",
    "Robinson EV Products Company | Theodore Adams | 4045040771 | theodore@yopmail.com"
]

payment_terms_options = [
    "All", "' DROP TABLE users --", "' OR '1''1", "' OR 11 --", "' OR 'a''a", "' UNION SELECT NULL --", "=1+2", "1+2",
    "15 Days of Credit", "20 Days Payment", "45 Days", "90 Days", "A TESR", "Apple Digital Production",
    "Cash Payment", "Credit Card Payment", "Electronic Transfer", "FIrst Payment on First Day", "Letter of Credit",
    "Mango", "Mango1", "Net 30", "Net Payment Terms", "Partial Payments", "Payment after 15 Days of Delivery",
    "Payment in Advance", "test date format", "Test Edit Payment Terms", "Test Payment Term", "TEST PAYMENT TERMS",
    "TEST PAYMENT TERMS 05-03-2025", "TEST PAYMENT TERMS 2"
]

discount_level_options = [
    "All", "Percentage Discounts/7", "Customer Loyalty Discount/8", "Promotional Discounts/5", "Regular Customer/12",
    "COD/30", "Customer Discount/17", "Test Discount Level/7", "Advance Payment Discount /25.99", "Test Date Format/10",
    "Test Edit Discount.........../25.36", "TEST DISCOUNT LEVEL 0001/13", "TEST DISCOUNT LEVEL 21-01-2025/28",
    "TEST DISCOUNT LEVEL 24-01-2025/12", "TEST DL 05-03-2025/10", "Apple Production Airppods/9", "Pay Cash Discount/25",
    "DL 16042025/20"
]

position_options = [
    "All", "AA", "ADMIN", "ADMIN CUSTOMER", "ADMIN MAIL", "CUSTOME?VENDOR", "CUSTOMER", "CUSTOMER 2",
    "CUSTOMER ADMIN", "CUSTOMER MAIL", "CV", "DDD", "EDIT ADMIN NAME", "EDIT ADMIN.....", "EDIT ADMIN..........",
    "EDIT CV NEW", "EDIT VENDOR........", "EIDI CUSTOMER............", "EXECUTIVE CHAIRPERSON", "HRM", "INDIVIDUAL",
    "MANAGER", "QWE", "S", "SALESPERSON", "TEST", "TS", "VENDOR"
]

export_options = ["None", "PDF", "Excel", "Print"]

# Tkinter input
def get_user_input():
    def submit():
        result['both'] = both_var.get()
        result['customer'] = customer_var.get()
        result['vendor'] = vendor_var.get()
        result['cust_vendor_value'] = cust_vendor_var.get()
        result['include_contact'] = include_contact_var.get()
        result['payment_term'] = payment_term_var.get()
        result['discount_level'] = discount_level_var.get()
        result['position'] = position_var.get()
        result['export'] = export_var.get()
        root.destroy()

    result = {}
    root = tk.Tk()
    root.title("Customer/Vendor Details Filter")

    both_var = tk.BooleanVar()
    customer_var = tk.BooleanVar()
    vendor_var = tk.BooleanVar()
    cust_vendor_var = tk.StringVar(value="All")
    include_contact_var = tk.BooleanVar()
    payment_term_var = tk.StringVar(value="All")
    discount_level_var = tk.StringVar(value="All")
    position_var = tk.StringVar(value="All")
    export_var = tk.StringVar(value="None")

    tk.Checkbutton(root, text="Both", variable=both_var).pack(anchor='w')
    tk.Checkbutton(root, text="Customer", variable=customer_var).pack(anchor='w')
    tk.Checkbutton(root, text="Vendor", variable=vendor_var).pack(anchor='w')

    tk.Label(root, text="Customer/Vendor Dropdown").pack(anchor='w')
    ttk.Combobox(root, textvariable=cust_vendor_var, values=customer_vendor_options, width=80).pack()

    tk.Checkbutton(root, text="Include contact", variable=include_contact_var).pack(anchor='w')

    tk.Label(root, text="Payment Terms").pack(anchor='w')
    ttk.Combobox(root, textvariable=payment_term_var, values=payment_terms_options, width=80, state="readonly").pack()

    tk.Label(root, text="Discount Level").pack(anchor='w')
    ttk.Combobox(root, textvariable=discount_level_var, values=discount_level_options, width=80, state="readonly").pack()

    tk.Label(root, text="Position").pack(anchor='w')
    ttk.Combobox(root, textvariable=position_var, values=position_options, width=80, state="readonly").pack()

    tk.Label(root, text="Export Format").pack(anchor='w')
    ttk.Combobox(root, textvariable=export_var, values=export_options, width=40, state="readonly").pack()

    tk.Button(root, text="Submit", command=submit).pack(pady=10)
    root.mainloop()
    return result

user_input = get_user_input()

# Selenium setup
driver = webdriver.Edge()
wait = WebDriverWait(driver, 20)
driver.maximize_window()
driver.get("https://softwaredevelopmentsolution.com/Contacts/Reports/CustomerVendorDetails")

# Login
wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
print("✅ Logged in")

# Wait for loader
try:
    wait.until(EC.invisibility_of_element_located((By.ID, "processing-scroll-container")))
except:
    pass

# Checkboxes
if user_input['both']:
    wait.until(EC.element_to_be_clickable((By.ID, "CustVendBoth"))).click()
if user_input['customer']:
    wait.until(EC.element_to_be_clickable((By.ID, "CustVendCustomer"))).click()
if user_input['vendor']:
    wait.until(EC.element_to_be_clickable((By.ID, "CustVendVendor"))).click()

# Customer/Vendor Dropdown
if user_input['cust_vendor_value'] and user_input['cust_vendor_value'] != "All":
    try:
        wait.until(EC.element_to_be_clickable((By.ID, "sl_chosen"))).click()
        search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#sl_chosen .chosen-search input")))
        search_term = user_input['cust_vendor_value'].split("|")[1].strip()
        search_input.send_keys(search_term)
        time.sleep(1)
        options = driver.find_elements(By.CSS_SELECTOR, "#sl_chosen .chosen-results li.active-result")
        for option in options:
            if search_term in option.text:
                driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", option)
                print(f"✔️ Selected: {option.text}")
                break
    except Exception as e:
        print(f"❌ Dropdown error: {e}")

# Include Contact
if user_input['include_contact']:
    try:
        checkbox = wait.until(EC.presence_of_element_located((By.ID, "Maincontact")))
        if not checkbox.is_selected():
            driver.execute_script("arguments[0].click();", checkbox)
    except Exception as e:
        print(f"❌ Include Contact checkbox error: {e}")

# Payment Term
if user_input['payment_term'] != "All":
    select = wait.until(EC.presence_of_element_located((By.ID, "ddlpaymentTerm")))
    for option in select.find_elements(By.TAG_NAME, "option"):
        if option.text.strip() == user_input['payment_term']:
            driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'))",
                                  select, option.get_attribute("value"))
            break

# Discount Level
if user_input['discount_level'] != "All":
    select = wait.until(EC.presence_of_element_located((By.ID, "ddldiscountLevel")))
    for option in select.find_elements(By.TAG_NAME, "option"):
        if option.text.strip() == user_input['discount_level']:
            driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'))",
                                  select, option.get_attribute("value"))
            break

# Position
if user_input['position'] != "All":
    select = wait.until(EC.presence_of_element_located((By.ID, "ddlposition")))
    for option in select.find_elements(By.TAG_NAME, "option"):
        if option.text.strip() == user_input['position']:
            driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'))",
                                  select, option.get_attribute("value"))
            break

# Generate Report
try:
    generate_btn = wait.until(EC.element_to_be_clickable((By.ID, "submitbtn")))
    driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", generate_btn)
    print("📄 Report generated.")
    time.sleep(5)
except Exception as e:
    print(f"❌ Failed to click Generate Report: {e}")

# Export
export_choice = user_input['export']
if export_choice != "None":
    try:
        export_btn = wait.until(EC.element_to_be_clickable((By.ID, "exportbtn")))
        driver.execute_script("arguments[0].click();", export_btn)
        time.sleep(1)

        if export_choice in ["PDF", "Excel"]:
            driver.execute_script(f"GenerateReport('{export_choice}')")
            time.sleep(5)  # Wait for export to complete
            print(f"📁 Exported to {export_choice}")
        elif export_choice == "Print":
            driver.execute_script("GenerateReport('Print')")
            print("🖨️ Print preview triggered.")
            time.sleep(5)
            # Attempt to close print window if it opened
            main_window = driver.current_window_handle
            all_windows = driver.window_handles
            for handle in all_windows:
                if handle != main_window:
                    driver.switch_to.window(handle)
                    driver.close()
                    driver.switch_to.window(main_window)
                    print("🪟 Closed print window.")
    except Exception as e:
        print(f"❌ Error in exporting report: {e}")
else:
    print("⏭️ No export selected.")

time.sleep(5)
driver.quit()
