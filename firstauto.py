from selenium import webdriver

# Set up Chrome options (optional)
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")  # Opens Chrome in maximized mode

# Initialize the Chrome WebDriver
driver = webdriver.Chrome(options=options)

# Open a website
driver.get("https://www.google.com")

# Keep the browser open for a few seconds
input("Press Enter to close the browser...")

# Close the browser
driver.quit()
