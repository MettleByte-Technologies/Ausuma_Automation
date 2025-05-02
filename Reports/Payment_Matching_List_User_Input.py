import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import time

# --- Tkinter GUI popup ---
class FilterPopup:
    def __init__(self, root):
        self.root = root
        self.root.title("Payment Matching List Filter")

        # Role checkboxes
        self.is_customer = tk.BooleanVar()
        self.is_vendor = tk.BooleanVar()
        self.is_employee = tk.BooleanVar()

        tk.Checkbutton(root, text="Customer", variable=self.is_customer, command=self.update_dropdowns).grid(row=0, column=0, sticky='w')
        tk.Checkbutton(root, text="Vendor", variable=self.is_vendor, command=self.update_dropdowns).grid(row=0, column=1, sticky='w')
        tk.Checkbutton(root, text="Employee", variable=self.is_employee, command=self.update_dropdowns).grid(row=0, column=2, sticky='w')

        # Customer dropdown
        self.selected_customer = tk.StringVar()
        self.customer_menu = ttk.Combobox(root, textvariable=self.selected_customer, values=[
            "Ola Shop001 | js10855@gmail.com",
            "Jimmy Turner",
            "John White | zainfast503@gmail.com",
            "Harry Smith | zainfast503@gmail.com",
            "Lucas Adams | quick@yopmail.com"
        ], state="readonly")
        tk.Label(root, text="Customer:").grid(row=1, column=0, sticky='w')
        self.customer_menu.grid(row=1, column=1, columnspan=2, sticky='ew')

        # Vendor dropdown
        self.selected_vendor = tk.StringVar()
        self.vendor_menu = ttk.Combobox(root, textvariable=self.selected_vendor, values=[
            "Theodore Adams | krunal@ausuma.com",
            "Liam  White | krunal@ausuma.com",
            "William Turner | quick@yopmail.com",
            "Jimmy Turner",
            "OLA Vendor | adam@gmail.com",
            "BroToMotive | prajapatibalwant12345@gmail.com",
            "OLA-M | OLAM@yopmail.com",
            "OLA MS HUB | liam@yopmail.com",
            "OLA MFR Vendor | william@yopmail.com",
        ], state="readonly")
        tk.Label(root, text="Vendor:").grid(row=2, column=0, sticky='w')
        self.vendor_menu.grid(row=2, column=1, columnspan=2, sticky='ew')

        # Employee dropdown
        self.selected_employee = tk.StringVar()
        self.employee_menu = ttk.Combobox(root, textvariable=self.selected_employee, values=[
            "Christopher Smith | ola1234@yopmail.com",
            "Noha Scott | noha@gmail.com",
            "James Green | james@gmail.com",
            "Henry  Brown | henry@gmail.com",
            "Jems Turner",
            "Adam Adams | adam@yopmail.com",
            "Lucas White | lucas@yopmail.com",
            "test emp for Ola test",
            "Adam1 Adams",
            "Kiran Suthar | kiran@yopmail.com",
            "Drashti Patel | drashti@yopmail.com",
            "Test Employee",
            "Test User Access T | testuseraccess@yopmail.com",
            "Jigar  Shah | js_test@yopmail.com",
            "Piyush  Suthar | piyush@yopmail.com | Sales Manager",
            "Test Date Format | testdate@yopmail.com",
            "Test Password  Employee | prajapatibalwant12345@gmail.com",
            "Terry Smith | terry@yopmail.com | Service Manager",
            "Test TestEdit | testemployeeeee@yopmail.com | Test Edit",
            "Employee   of Gusto | empgusto@yopmail.com",
            "TEST USER  ACCESS | prajapatibalwant555@gmail.com",
            "TEST ADD  EMPLOYEE | testaddemployee@yopmail.com | Sales Manager",
            "TEST ADD EMP  ON 20-12 | testemp@yopmail.com",
            "J5 S6 | empnewtestbug@yopmail.com | TEST EMP",
            "TEST EMP  ON 23-01-2025 | testemp23012025@yopmail.com",
            "123 456 | jigs@yopmail.com",
            "TEST uubed",
            "TEST als0000",
            "test uyy",
            "ESTER wATES | teestemp@yopmail.com | TEST",
            "jerry address | addressemp@yopmail.com",
            "kiran kumar | kiran123@yopmail.com",
            "Mango Sudama | sudama1234@email.com",
            "Patel Parshad | patelparsad@yopmail.com | Sales Manager",
            "employee araice | addempnew@yopmail.com | Employee",
            "ps ps | pspsp@yopmail.com | Employee"
        ], state="readonly")
        tk.Label(root, text="Employee:").grid(row=3, column=0, sticky='w')
        self.employee_menu.grid(row=3, column=1, columnspan=2, sticky='ew')

        # Initially hide dropdowns
        self.vendor_menu.grid_remove()
        self.employee_menu.grid_remove()

        # Submit button
        tk.Button(root, text="Run Report", command=self.submit).grid(row=4, column=0, columnspan=3)

    def update_dropdowns(self):
        # Show customer dropdown by default, if no checkboxes are selected
        if not (self.is_customer.get() or self.is_vendor.get() or self.is_employee.get()):
            self.customer_menu.grid()
        else:
            if self.is_customer.get():
                self.customer_menu.grid()
            else:
                self.customer_menu.grid_remove()

            if self.is_vendor.get():
                self.vendor_menu.grid()
            else:
                self.vendor_menu.grid_remove()

            if self.is_employee.get():
                self.employee_menu.grid()
            else:
                self.employee_menu.grid_remove()

    def submit(self):
        self.root.quit()

def run_tkinter_popup():
    root = tk.Tk()
    popup = FilterPopup(root)
    root.mainloop()
    return popup

# --- Selenium script ---
def run_selenium_script(is_customer, is_vendor, is_employee, customer_value, vendor_value, employee_value):
    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 20)
    driver.maximize_window()
    driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/PaymentMatchingList")

    # Login
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()

    # Wait for page load
    wait.until(EC.invisibility_of_element_located((By.ID, "processing-scroll-container")))

    # Handle checkboxes
    checkbox_ids = []
    if is_customer:
        checkbox_ids.append("CustVendCustomer")
    if is_vendor:
        checkbox_ids.append("CustVendVendor")
    if is_employee:
        checkbox_ids.append("Employee")

    for cid in checkbox_ids:
        checkbox = wait.until(EC.presence_of_element_located((By.ID, cid)))
        wait.until(EC.invisibility_of_element_located((By.ID, "processing-scroll-container")))
        driver.execute_script("arguments[0].click();", checkbox)

    # Wait before dropdown
    wait.until(EC.invisibility_of_element_located((By.ID, "processing-scroll-container")))

    # Handle customer dropdown interaction
    if is_customer and customer_value:
        customer_dropdown = wait.until(EC.presence_of_element_located((By.ID, "sl_chosen")))
        customer_dropdown.click()
        search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#sl_chosen .chosen-search input")))
        search_input.send_keys(customer_value)
        time.sleep(1)
        result = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#sl_chosen .chosen-results li.active-result")))
        result.click()

    if is_vendor and vendor_value:
        vendor_dropdown = wait.until(EC.presence_of_element_located((By.ID, "sl_chosen")))
        vendor_dropdown.click()
        search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#sl_chosen .chosen-search input")))
        search_input.send_keys(vendor_value)
        time.sleep(1)
        result = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#sl_chosen .chosen-results li.active-result")))
        result.click()

    if is_employee and employee_value:
        employee_dropdown = wait.until(EC.presence_of_element_located((By.ID, "sl_chosen")))
        employee_dropdown.click()
        search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#sl_chosen .chosen-search input")))
        search_input.send_keys(employee_value)
        time.sleep(1)
        result = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#sl_chosen .chosen-results li.active-result")))
        result.click()

    # Generate Report
    wait.until(EC.element_to_be_clickable((By.ID, "submitbtn"))).click()
    print("📄 Report generated")

# --- Run ---
popup = run_tkinter_popup()

run_selenium_script(
    is_customer=popup.is_customer.get(),
    is_vendor=popup.is_vendor.get(),
    is_employee=popup.is_employee.get(),
    customer_value=popup.selected_customer.get(),
    vendor_value=popup.selected_vendor.get(),
    employee_value=popup.selected_employee.get()
)
