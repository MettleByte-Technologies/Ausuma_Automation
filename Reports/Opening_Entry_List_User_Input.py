import time
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Step 1: Tkinter Popup for Inputs
def get_user_input():
    root = tk.Tk()
    root.title("Opening Entry List Inputs")

    # From Date
    tk.Label(root, text="From Date (MM/DD/YYYY):").grid(row=0, column=0, padx=10, pady=5)
    from_date_entry = tk.Entry(root)
    from_date_entry.grid(row=0, column=1, padx=10, pady=5)

    # To Date
    tk.Label(root, text="To Date (MM/DD/YYYY):").grid(row=1, column=0, padx=10, pady=5)
    to_date_entry = tk.Entry(root)
    to_date_entry.grid(row=1, column=1, padx=10, pady=5)

    # Document No Dropdown
    tk.Label(root, text="Document No:").grid(row=2, column=0, padx=10, pady=5)
    doc_no_dropdown = ttk.Combobox(root, values=[
        "OPN000007", "OPN000008", "OPN000009", "OPN000011", "OPN000012",
        "OPN000013", "OPN000014", "OPN000015", "OPN000016"
    ])
    doc_no_dropdown.grid(row=2, column=1, padx=10, pady=5)
    doc_no_dropdown.set("OPN000007")

    # Export Option
    tk.Label(root, text="Export Option:").grid(row=3, column=0, padx=10, pady=5)
    export_option = ttk.Combobox(root, values=["Pdf", "Excel", "Print"])
    export_option.grid(row=3, column=1, padx=10, pady=5)
    export_option.set("Pdf")

    def submit():
        root.quit()

    tk.Button(root, text="Submit", command=submit).grid(row=4, column=0, columnspan=2, pady=10)
    root.mainloop()

    return from_date_entry.get(), to_date_entry.get(), doc_no_dropdown.get(), export_option.get()

# Step 2: Selenium Automation
def automate_report(from_date, to_date, doc_no, export_option):
    driver = webdriver.Edge()
    wait = WebDriverWait(driver, 20)
    driver.maximize_window()
    driver.get("https://softwaredevelopmentsolution.com/Accounting/Reports/OpeningEntryList")

    # Login
    wait.until(EC.presence_of_element_located((By.ID, "Email"))).send_keys("ola123@yopmail.com")
    wait.until(EC.presence_of_element_located((By.ID, "Password"))).send_keys("1")
    wait.until(EC.element_to_be_clickable((By.ID, "LoginSubmit"))).click()
    print("✅ Logged in")

    # Fill From Date
    try:
        from_input = wait.until(EC.presence_of_element_located((By.ID, "FromDate")))
        from_input.clear()
        from_input.send_keys(from_date)
    except Exception as e:
        print("⚠️ Skipping From Date:", e)

    # Fill To Date
    try:
        to_input = wait.until(EC.presence_of_element_located((By.ID, "ToDate")))
        to_input.clear()
        to_input.send_keys(to_date)
    except Exception as e:
        print("⚠️ Skipping To Date:", e)

    # Call SearchDocument directly
    try:
        driver.execute_script("SearchDocument('OP')")
        print("📥 SearchDocument('OP') called")
        time.sleep(2)
    except Exception as e:
        print(f"❌ Error in calling SearchDocument: {e}")
        driver.quit()
        return

    # Select Document No
    try:
        dropdown_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-id='ddlDocNo']")))
        dropdown_btn.click()
        time.sleep(1)

        option_xpath = f"//span[@class='title' and text()='{doc_no}']"
        document_option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
        document_option.click()
        print(f"✅ Document No '{doc_no}' selected")
    except Exception as e:
        print(f"❌ Error in selecting document: {e}")
        driver.quit()
        return

    # Click Generate Report
    try:
        generate_btn = wait.until(EC.element_to_be_clickable((By.ID, "submitbtn")))
        generate_btn.click()
        print("📊 Report generated")
    except Exception as e:
        print(f"❌ Error in clicking generate report button: {e}")
        driver.quit()
        return

    time.sleep(5)
    wait.until(EC.invisibility_of_element_located((By.ID, "processing-scroll")))

    # Export
    try:
        export_btn = wait.until(EC.element_to_be_clickable((By.ID, "exportbtn")))
        driver.execute_script("arguments[0].click();", export_btn)
        time.sleep(1)

        driver.execute_script(f"GenerateReport('{export_option}')")
        print(f"📁 Exported to {export_option}")
    except Exception as e:
        print(f"❌ Error in exporting report: {e}")
        driver.quit()
        return

    time.sleep(5)
    driver.quit()

# Main
if __name__ == "__main__":
    from_date, to_date, doc_no, export_option = get_user_input()
    automate_report(from_date, to_date, doc_no, export_option)
