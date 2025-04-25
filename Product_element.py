
# ========== Step 4: Select Weight ==========
try:
    # Locate the weight input field
    weight_input = wait.until(EC.presence_of_element_located((By.ID, "pRODUCT_UNITWEIGHT")))

    # Scroll to make sure the input is visible
    driver.execute_script("arguments[0].scrollIntoView(true);", weight_input)

    # Ensure the field is enabled before interacting
    wait.until(EC.element_to_be_clickable((By.ID, "pRODUCT_UNITWEIGHT")))

    # Clear any existing value and enter the new weight
    weight_input.clear()
    weight_input.send_keys("10")

    print("✔️ Weight value '10' added successfully")

except Exception as e:
    print(f"❌ Failed at Step 5 (Weight Value): {e}")

# ========== Step 5: Enter Weight ==========
try:
    wait.until(EC.presence_of_element_located((By.ID, "ddlWeightunit"))).send_keys("10")
    print("✔️ Weight value '10' added successfully")
except Exception as e:
    print(f"❌ Failed at Step 5 (Weight Value): {e}")


    # ========== Step 8: Carrier ==========
try:
    carrier_dropdown = Select(wait.until(EC.presence_of_element_located((By.ID, "ddlCarrier"))))
    carrier_dropdown.select_by_visible_text("Air Carriers")
    print("✔️ Carrier 'Air Carriers' selected successfully")
except Exception as e:
    print(f"❌ Failed at Step 8 (Carrier): {e}")

# ========== Step 9: Package ==========
try:
    package_dropdown = Select(wait.until(EC.presence_of_element_located((By.ID, "ddlPackage"))))
    package_dropdown.select_by_visible_text("envelopes and mailers1")
    print("✔️ Package 'envelopes and mailers1' selected successfully")
except Exception as e:
    print(f"❌ Failed at Step 9 (Package): {e}")

# ========== Step 10: Lead Time ==========
try:
    wait.until(EC.presence_of_element_located((By.ID, "pRODUCT_LEADTIME"))).send_keys("10")
    print("✔️ Lead Time '10' added successfully")
except Exception as e:
    print(f"❌ Failed at Step 10 (Lead Time): {e}")

# ========== Step 11: Taxable ==========
try:
    wait.until(EC.element_to_be_clickable((By.ID, "pRODUCT_TAXABLE"))).click()
    print("✔️ Taxable checkbox selected successfully")
except Exception as e:
    print(f"❌ Failed at Step 11 (Taxable): {e}")

# ========== Step 12: Supplier Identifier ==========
try:
    wait.until(EC.presence_of_element_located((By.ID, "pRODUCT_SUPPLIERIDENTIFIER"))).send_keys("XYZ0001")
    print("✔️ Supplier Identifier added successfully")
except Exception as e:
    print(f"❌ Failed at Step 12 (Supplier Identifier): {e}")

# ========== Step 13: Avalara Taxcode ==========
try:
    taxcode_dropdown = Select(wait.until(EC.presence_of_element_located((By.ID, "avaTaxCode"))))
    taxcode_dropdown.select_by_visible_text("PB100000")
    print("✔️ Avalara Taxcode 'PB100000' selected successfully")
except Exception as e:
    print(f"❌ Failed at Step 13 (Taxcode): {e}")

    