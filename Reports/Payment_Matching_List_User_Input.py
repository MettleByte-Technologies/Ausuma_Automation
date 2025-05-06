import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# --- Tkinter GUI popup ---
class FilterPopup:
    def __init__(self, root):
        self.root = root
        self.root.title("Payment Matching List Filter")

        self.is_customer = tk.BooleanVar()
        self.is_vendor = tk.BooleanVar()
        self.is_employee = tk.BooleanVar()

        tk.Checkbutton(root, text="Customer", variable=self.is_customer, command=self.update_dropdowns).grid(row=0, column=0, sticky='w')
        tk.Checkbutton(root, text="Vendor", variable=self.is_vendor, command=self.update_dropdowns).grid(row=0, column=1, sticky='w')
        tk.Checkbutton(root, text="Employee", variable=self.is_employee, command=self.update_dropdowns).grid(row=0, column=2, sticky='w')

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

        self.selected_employee = tk.StringVar()
        self.employee_menu = ttk.Combobox(root, textvariable=self.selected_employee, values=[
            "Christopher Smith | ola1234@yopmail.com",
            "Noha Scott | noha@gmail.com",
            "James Green | james@gmail.com",
            "Henry  Brown | henry@gmail.com",
            "Jems Turner",
            "Adam Adams | adam@yopmail.com",
            "Lucas White | lucas@yopmail.com"
        ], state="readonly")
        tk.Label(root, text="Employee:").grid(row=3, column=0, sticky='w')
        self.employee_menu.grid(row=3, column=1, columnspan=2, sticky='ew')

        self.vendor_menu.grid_remove()
        self.employee_menu.grid_remove()

        tk.Label(root, text="From Date (MM/DD/YYYY):").grid(row=4, column=0, sticky='w')
        self.from_date = tk.Entry(root)
        self.from_date.grid(row=4, column=1, columnspan=2, sticky='ew')

        tk.Label(root, text="To Date (MM/DD/YYYY):").grid(row=5, column=0, sticky='w')
        self.to_date = tk.Entry(root)
        self.to_date.grid(row=5, column=1, columnspan=2, sticky='ew')

        tk.Label(root, text="Matching No:").grid(row=6, column=0, sticky='w')
        self.matching_no = tk.StringVar()
        self.matching_dropdown = ttk.Combobox(root, textvariable=self.matching_no, values=[
            "PIYUSH000216", "PIYUSH000217", "PIYUSH000219", "PIYUSH000220",
            "PM000121", "PM000125", "PM000131", "PM000141", "PM000142",
            "PMN000145", "PMN000146", "PMN000147", "PMN000148", "PMN000150"
        ], state="readonly")
        self.matching_dropdown.grid(row=6, column=1, columnspan=2, sticky='ew')

        self.enable_currency = tk.BooleanVar()
        tk.Checkbutton(root, text="Enable Currency", variable=self.enable_currency).grid(row=7, column=0, sticky='w')

        tk.Button(root, text="Run Report", command=self.submit).grid(row=8, column=0, columnspan=3)

    def update_dropdowns(self):
        self.customer_menu.grid() if self.is_customer.get() else self.customer_menu.grid_remove()
        self.vendor_menu.grid() if self.is_vendor.get() else self.vendor_menu.grid_remove()
        self.employee_menu.grid() if self.is_employee.get() else self.employee_menu.grid_remove()

    def submit(self):
        self.root.quit()

def run_tkinter_popup():
    root = tk.Tk()
    popup = FilterPopup(root)
    root.mainloop()
    return popup

def select_chosen_option(driver, wait, chosen_id, value):
    try:
        dropdown = wait.until(EC.element_to_be_clickable((By.ID, f"{chosen_id}_chosen")))
        dropdown.click()
        search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, f"#{chosen_id}_chosen .chosen-search input")))
        search_input.clear()
        search_input.send_keys(value)
        time.sleep(1.5)
        result = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"#{chosen_id}_chosen .chosen-results li.active-result")))
        result.click()
    except Exception as e:
        print(f"❌ Failed to select {value} from {chosen_id}: {e}")

def run_selenium_script(popup):
    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 20)
    driver.maximize_window()
    driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/PaymentMatchingList")

    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()

    wait.until(EC.invisibility_of_element_located((By.ID, "processing-scroll-container")))

    if popup.is_customer.get():
        wait.until(EC.element_to_be_clickable((By.ID, "CustVendCustomer"))).click()
    if popup.is_vendor.get():
        wait.until(EC.element_to_be_clickable((By.ID, "CustVendVendor"))).click()
    if popup.is_employee.get():
        wait.until(EC.element_to_be_clickable((By.ID, "Employee"))).click()

    wait.until(EC.invisibility_of_element_located((By.ID, "processing-scroll-container")))

    if popup.is_customer.get():
        select_chosen_option(driver, wait, "sl", popup.selected_customer.get())
    if popup.is_vendor.get():
        select_chosen_option(driver, wait, "sl", popup.selected_vendor.get())
    if popup.is_employee.get():
        select_chosen_option(driver, wait, "sl", popup.selected_employee.get())

    wait.until(EC.presence_of_element_located((By.ID, "FromDate"))).send_keys(popup.from_date.get())
    wait.until(EC.presence_of_element_located((By.ID, "ToDate"))).send_keys(popup.to_date.get())

    if popup.matching_no.get():
        select_chosen_option(driver, wait, "ddlDocNo", popup.matching_no.get())

    currency_checkbox = wait.until(EC.presence_of_element_located((By.ID, "EnableCurrency")))
    if popup.enable_currency.get() != currency_checkbox.is_selected():
        driver.execute_script("arguments[0].click();", currency_checkbox)

    wait.until(EC.element_to_be_clickable((By.ID, "submitbtn"))).click()
    time.sleep(15)
    print("✅ Report generated successfully.")

# --- Main ---
popup = run_tkinter_popup()
run_selenium_script(popup)
