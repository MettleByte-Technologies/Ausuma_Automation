from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import ttk
import threading
import time

# Dropdown values
product_list = [
    "All", "A B C Product", "A R T Product", "ABC 28-03-2025",
    "ABC TEST  28-03-2025", "Apple Production Product", "Arice Priduct",
    "avalara test product", "Battery", "Battery Management Systems", "Bike Tire 100"
]

location_list = [
    "All", "OLA Electric Showroom", "OLA Electric Service Station", "OLA Electric Showroom (TW)"
]

location_bin_list = [
    "All", "EV Warehouse 001", "OLA Products", "No Location", "EV Servise 1", "EV Parts"
]

export_options = ["Pdf", "Excel", "Print"]

# Global selections
selected_product = None
selected_location = None
selected_location_bin = None
selected_export = None

# Tkinter popup
def show_popup():
    def submit():
        global selected_product, selected_location, selected_location_bin, selected_export
        selected_product = product_combo.get()
        selected_location = location_combo.get()
        selected_location_bin = bin_combo.get()
        selected_export = export_combo.get()
        root.destroy()

    root = tk.Tk()
    root.title("Select Filters")

    tk.Label(root, text="Choose a product:").pack(pady=5)
    product_combo = ttk.Combobox(root, values=product_list, width=40)
    product_combo.pack()
    product_combo.current(0)

    tk.Label(root, text="Choose a location:").pack(pady=5)
    location_combo = ttk.Combobox(root, values=location_list, width=40)
    location_combo.pack()
    location_combo.current(0)

    tk.Label(root, text="Choose a location bin:").pack(pady=5)
    bin_combo = ttk.Combobox(root, values=location_bin_list, width=40)
    bin_combo.pack()
    bin_combo.current(0)

    tk.Label(root, text="Export Format:").pack(pady=5)
    export_combo = ttk.Combobox(root, values=export_options, width=40)
    export_combo.pack()
    export_combo.current(1)  # Default to Excel

    tk.Button(root, text="Run", command=submit).pack(pady=10)
    root.mainloop()

# Selenium automation
def run_automation():
    global selected_product, selected_location, selected_location_bin, selected_export

    options = webdriver.EdgeOptions()
    options.add_argument("start-maximized")
    driver = webdriver.Edge(options=options)
    wait = WebDriverWait(driver, 20)

    driver.get("https://softwaredevelopmentsolution.com/Inventory/Reports/CurrentProductStock")

    try:
        # Login
        wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
        wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
        wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
        print("✅ Logged in")

        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "preloader-it")))
        print("✅ Page loaded")

        time.sleep(1)  # Ensure dropdowns are ready

        # --- Select Product ---
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@data-id, 'ddlProduct')]"))).click()
        time.sleep(1)
        product_xpath = f"//ul[@class='dropdown-menu inner']//span[contains(text(), \"{selected_product}\")]"
        wait.until(EC.element_to_be_clickable((By.XPATH, product_xpath))).click()
        print(f"✅ Product selected: {selected_product}")

        # --- Select Location ---
        wait.until(EC.element_to_be_clickable((By.ID, "ddlWarehouse_chosen"))).click()
        time.sleep(1)
        location_xpath = f"//div[@id='ddlWarehouse_chosen']//li[contains(text(), \"{selected_location}\")]"
        wait.until(EC.element_to_be_clickable((By.XPATH, location_xpath))).click()
        print(f"✅ Location selected: {selected_location}")

        # --- Select Location Bin ---
        wait.until(EC.element_to_be_clickable((By.ID, "ddlLocation_chosen"))).click()
        time.sleep(1)
        bin_xpath = f"//div[@id='ddlLocation_chosen']//li[contains(text(), \"{selected_location_bin}\")]"
        wait.until(EC.element_to_be_clickable((By.XPATH, bin_xpath))).click()
        print(f"✅ Location Bin selected: {selected_location_bin}")

        # --- Generate Report ---
        generate_button = wait.until(EC.presence_of_element_located((By.ID, "submitbtn")))
        time.sleep(1)
        driver.execute_script("arguments[0].click();", generate_button)
        print("✅ Generate Report clicked")

        time.sleep(5)  # Allow report to load

        # --- Export Report ---
        export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Export')]")))
        export_btn.click()
        print("✅ Export menu opened")

        export_xpath = f"//ul[@class='dropdown-menu show']//a[contains(text(), '{selected_export}')]"
        wait.until(EC.element_to_be_clickable((By.XPATH, export_xpath))).click()
        print(f"✅ Exported as: {selected_export}")

        time.sleep(3)

    except Exception as e:
        print(f"⚠️ Error: {e}")

    # Optional: driver.quit()

# Show popup
threading.Thread(target=show_popup).start()

# Wait until user makes selections
while selected_product is None or selected_location is None or selected_location_bin is None or selected_export is None:
    time.sleep(1)

# Start automation
run_automation()
