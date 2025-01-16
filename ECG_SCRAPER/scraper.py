from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from datetime import datetime
import pandas as pd
import time

# Path to your existing Chrome profile
chrome_profile_path = r"C:\\Users\\OMEN 16\\AppData\\Local\\Google\\Chrome\\User Data\\"

# Set up Chrome options to use the profile
chrome_options = webdriver.ChromeOptions()
# chrome_options.add_argument(f"user-data-dir={chrome_profile_path}")  # Use your profile directory
# chrome_options.add_argument("profile-directory=Profile 7")  # Replace with your specific profile
chrome_options.add_argument("--no-sandbox")  # Prevent sandbox issues
chrome_options.add_argument("--disable-dev-shm-usage")  # Use shared memory
chrome_options.add_argument("--remote-debugging-port=9222")  # Avoid debugging conflicts
chrome_options.add_argument("--disable-gpu")  # Disable GPU (optional)
chrome_options.add_argument("--start-maximized") 
# Initialize the WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)




date_found =False
username = "techops_recons_1@kowri.app"
password = "b0Ew5\"YgI6K3jb"

driver.get("https://oltpv.ecg.com.gh/")
driver.implicitly_wait(5)

email_signin_btn = driver.find_element(By.XPATH, "/html/body/main/div[2]/div/div/div/div/div/div/div[2]/a")
email_signin_btn.click()
email_input = driver.find_element(By.XPATH, "/html/body/main/div[2]/div/main/div/div/div/div/div/form/div/div[2]/input")

email_input.send_keys(username)
email_input.send_keys(Keys.ENTER)

# WebDriverWait(driver, 20).until(EC.visibility_of((By.XPATH, "/html/body/main/div[2]/div/div/div/div[2]/div[1]/div/table")))
time.sleep(15)

filter_date = "2025-01-15"
filter_date_obj = datetime.strptime(filter_date, "%Y-%m-%d")

table_head = driver.find_element(By.CSS_SELECTOR, ".table thead")
table_heads = table_head.find_elements(By.TAG_NAME, "th")

table_headers = [head.text for head in table_heads]


# Locate the table body
table_body = driver.find_element(By.CSS_SELECTOR, ".table tbody")
# # Locate the table header
# table_header = table_body.find_element(By.TAG_NAME, "thead")

# Extract rows
rows = table_body.find_elements(By.TAG_NAME, "tr")
headers = table_body.find_elements(By.TAG_NAME, "th")

# Initialize list to store table data
table_data = []
trans_ids = []

# Iterate over each row
while not date_found:
    # Re-fetch table body and rows on each loop
    table_body = driver.find_element(By.CSS_SELECTOR, ".table tbody")
    rows = table_body.find_elements(By.TAG_NAME, "tr")
    counter = 0

    for row in rows:
        counter += 1
        # Extract cells in the row
        cells = row.find_elements(By.TAG_NAME, "td")

        raw_date_text = cells[7].text.strip()
        row_menu = cells[9]

        try:
            row_date_obj = datetime.strptime(raw_date_text, "%a, %b %d, %Y, %I:%M %p")
            print("Row date:", row_date_obj)
            print("Filter date:", filter_date_obj)
        except ValueError:
            print("Invalid date format:", raw_date_text)
            continue

        if row_date_obj.date() == filter_date_obj.date():
            date_found = True
            row_data = [cell.text.strip() for cell in cells]
            table_data.append(row_data)

            # Click the menu and navigate to transaction details
            row_menu.click()
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    f"body > main > div:nth-child(2) > div > div > div > div:nth-child(3) > div.table-responsive-parent > div > table > tbody > tr:nth-child({counter}) > td.show > div > button"
                ))
            ).click()
            time.sleep(5)

            # Get transaction ID from the URL
            trans_id = driver.current_url.split("/")[-1]
            trans_ids.append(trans_id)

            # Navigate back to the original page
            driver.back()
            time.sleep(5)

            # Break the inner loop and restart row processing after navigating back
            break

        else:
            print("Skipping row with date:", raw_date_text)

    # WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > main > div:nth-child(2) > div > div > div > div:nth-child(3) > div.mt-5.mb-4.d-md-flex.justify-content-between.align-items-center.px-4 > nav > ul > li:nth-child(8) > span"))).click()
      
print(table_data)
print(trans_ids)

try:
    results_df = pd.DataFrame(data=table_data,columns=table_headers).to_excel("results.xlsx", index=False)
except:
    print("Error")

time.sleep(300)

# driver.quit()