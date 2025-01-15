from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime
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


username = "techops_recons_1@kowri.app"
password = "b0Ew5\"YgI6K3jb"

driver.get("https://oltpv.ecg.com.gh/")
driver.implicitly_wait(5)

email_signin_btn = driver.find_element(By.XPATH, "/html/body/main/div[2]/div/div/div/div/div/div/div[2]/a")
email_signin_btn.click()
email_input = driver.find_element(By.XPATH, "/html/body/main/div[2]/div/main/div/div/div/div/div/form/div/div[2]/input")

email_input.send_keys(username)
email_input.send_keys(Keys.ENTER)

time.sleep(20)

# # Specify your target date (make sure it matches the format you want to compare)
# target_date = "2025-01-14"  # Example: Compare only the date portion

# # Locate all rows in the table
# rows = driver.find_elements(By.XPATH, "/html/body/main/div[2]/div/div/div/div[2]/div[1]/div/table")

# # Iterate through each row to find the matching date
# for row in rows:
#     # Locate the date cell (adjust column index as per the table structure)
#     date_cell = row.find_element(By.XPATH, ".//td[8]")  # Assuming the date is in the first column
    
#     # Get the date text
#     date_text = date_cell.text.strip()  # Example: "Tue, Jan 14, 2025, 5:02 PM"
    
#     # Parse the date to a datetime object
#     parsed_date = datetime.strptime(date_text, "%a, %b %d, %Y, %I:%M %p")  # Adjust format as needed
    
#     # Extract only the date portion for comparison
#     if parsed_date.strftime("%Y-%m-%d") == target_date:
#         # Print or process the entire row data
#         cells = row.find_elements(By.TAG_NAME, "td")
#         row_data = [cell.text.strip() for cell in cells]
#         print(row_data)

rows = driver.find_elements(By.XPATH, "/html/body/main/div[2]/div/div/div/div[2]/div[1]/div/table")
for row in rows:
    cells = row.find_elements(By.TAG_NAME, "td")
    for cell in cells:
        print(cell.text)

time.sleep(300)

# driver.quit()
