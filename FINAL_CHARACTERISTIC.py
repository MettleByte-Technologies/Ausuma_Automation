import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
import threading
import logging
import os

# Try importing screen recording dependencies
try:
    import pyautogui
    import cv2
    import numpy as np
    RECORDING_AVAILABLE = True
except ImportError:
    print("⚠️ Screen recording dependencies not found. Skipping screen recording.")
    RECORDING_AVAILABLE = False

def record_screen(filename, stop_event):
    screen_size = pyautogui.size()
    fourcc = cv2.VideoWriter_fourcc(*"XVID")
    out = cv2.VideoWriter(filename, fourcc, 10.0, screen_size)
    print("🎥 Screen recording started...")

    while not stop_event.is_set():
        img = pyautogui.screenshot()
        frame = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2RGB)
        out.write(frame)

    out.release()
    print("🛑 Screen recording stopped and saved.")

def run_selenium(name, description, value_type, value_range, unit_required):
    stop_recording = threading.Event()
    if RECORDING_AVAILABLE:
        threading.Thread(target=record_screen, args=("screen_record.avi", stop_recording)).start()

    try:
        edge_options = EdgeOptions()
        edge_options.add_argument("--start-maximized")
        driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=edge_options)
    except Exception as e:
        print(f"❌ Failed to start Edge WebDriver: {e}")
        stop_recording.set()
        return

    try:
        driver.get("https://softwaredevelopmentsolution.com/Inventory/Characteristics/Create")
        wait = WebDriverWait(driver, 30)
        wait.until(EC.presence_of_element_located((By.ID, "Email")))

        username_field = wait.until(EC.visibility_of_element_located((By.ID, "Email")))
        password_field = driver.find_element(By.ID, "Password")
        login_button = driver.find_element(By.ID, "LoginSubmit")

        username_field.send_keys("ola123@yopmail.com")
        password_field.send_keys("ola@1234")
        login_button.click()
        print("✅ Logged in successfully")

        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            driver.switch_to.frame(iframes[0])

        name_field = wait.until(EC.element_to_be_clickable((By.ID, "cHARACTERISTIC_NAME")))
        name_field.send_keys(name)

        description_field = wait.until(EC.element_to_be_clickable((By.ID, "cHARACTERISTIC_DESCRIPTION")))
        description_field.send_keys(description)

        def select_div_dropdown(dropdown_id, option_text):
            try:
                dropdown = wait.until(EC.element_to_be_clickable((By.ID, dropdown_id)))
                dropdown.click()
                time.sleep(1)
                option_xpath = f"//div[@id='{dropdown_id}']//li[contains(text(), '{option_text}')]"
                option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
                option.click()
                driver.find_element(By.TAG_NAME, "body").click()
                time.sleep(1)
            except Exception as e:
                print(f"❌ Dropdown '{dropdown_id}' error: {e}")

        select_div_dropdown("cHARACTERISTIC_VALUETYPE_chosen", value_type)
        if value_type != "Boolean":
            select_div_dropdown("cHARACTERISTIC_VALUERANGE_chosen", value_range)
            select_div_dropdown("cHARACTERISTIC_ISUNITREQUIRED_chosen", unit_required)

        try:
            submit_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@class, 'btn-success') and text()='Submit']")))
            driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
            time.sleep(1)
            try:
                ActionChains(driver).move_to_element(submit_button).click().perform()
            except:
                try:
                    submit_button.click()
                except:
                    driver.execute_script("arguments[0].click();", submit_button)
            print("✅ Submit button clicked")
        except Exception as e:
            print(f"❌ Submit error: {e}")
            driver.save_screenshot("submit_error.png")

    except Exception as e:
        print(f"❌ Main error: {e}")
    finally:
        time.sleep(3)
        driver.quit()
        stop_recording.set()

def start_gui():
    def on_submit():
        name = name_entry.get()
        description = desc_entry.get()
        value_type = type_var.get()
        value_range = range_var.get()
        unit_required = unit_var.get()

        threading.Thread(target=run_selenium, args=(name, description, value_type, value_range, unit_required)).start()

    def on_type_change(event):
        selected = type_var.get()
        if selected == "Boolean":
            range_dropdown.config(state="disabled")
            unit_dropdown.config(state="disabled")
        else:
            range_dropdown.config(state="readonly")
            unit_dropdown.config(state="readonly")

    root = tk.Tk()
    root.title("Automation Form Input")
    root.geometry("400x400")

    tk.Label(root, text="Characteristic Name").pack(pady=5)
    name_entry = tk.Entry(root, width=40)
    name_entry.pack()

    tk.Label(root, text="Description").pack(pady=5)
    desc_entry = tk.Entry(root, width=40)
    desc_entry.pack()

    tk.Label(root, text="Value Type").pack(pady=5)
    type_var = tk.StringVar()
    type_dropdown = ttk.Combobox(root, textvariable=type_var, values=["Alpha", "Numeric", "Boolean"], state="readonly")
    type_dropdown.current(1)
    type_dropdown.pack()
    type_dropdown.bind("<<ComboboxSelected>>", on_type_change)

    tk.Label(root, text="Value Range").pack(pady=5)
    range_var = tk.StringVar()
    range_dropdown = ttk.Combobox(root, textvariable=range_var, values=["Close ended", "Open ended"], state="readonly")
    range_dropdown.current(0)
    range_dropdown.pack()

    tk.Label(root, text="Unit Required").pack(pady=5)
    unit_var = tk.StringVar()
    unit_dropdown = ttk.Combobox(root, textvariable=unit_var, values=["Yes", "No"], state="readonly")
    unit_dropdown.current(1)
    unit_dropdown.pack()

    tk.Button(root, text="Submit", command=on_submit, bg="green", fg="white").pack(pady=20)

    root.mainloop()

# Run GUI
start_gui()
    