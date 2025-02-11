import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from selenium.common.exceptions import TimeoutException

# Path to your Chromedriver executable
webdriver_path = '/Users/Sashank/OneDrive/Desktop/chromedriver-win64/chromedriver-win64/chromedriver.exe'
excel_file_path = r"C:\Users\Sashank\OneDrive\Desktop\ReceptionGuestFinal.xlsx"
video_file_path = r"C:\Users\Sashank\OneDrive\Desktop\Greeting.jpg"
user_data_dir = r'C:\\Users\\Sashank\\AppData\\Local\\Google\\Chrome\\User Data'

# Set Chrome options
chrome_options = Options()
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument(f'--user-data-dir={user_data_dir}')
#chrome_options.add_argument('--profile-directory=Default')

# Initialize Chrome WebDriver with service and options
try:
    driver = webdriver.Chrome(service=Service(webdriver_path), options=chrome_options)
    driver.get("https://web.whatsapp.com")
    print("Successfully opened Google with WebDriver.")
    
    # Example: Print title of the page
    print(driver.title)

    # Close the browser window
    #driver.quit()

except Exception as e:
    print(f"Failed to open Google with WebDriver: {str(e)}")
input("Press Enter after scanning QR code") 

def send_video_via_whatsapp(phone_number, video_path):
    try:
        # Open chat
        print(phone_number);
        time.sleep(2);
        driver.get(f"https://web.whatsapp.com/send?phone={phone_number}")
        time.sleep(3);
        print("Chat window should be loaded");

        # Attach video
        print(1)
        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '//div[@title="Attach"]')))
        attach_button = driver.find_element(By.XPATH, '//div[@title="Attach"]')
        attach_button.click()
        time.sleep(10);
        # Attach video file
        video_input = driver.find_element(By.XPATH, '//input[@accept="*"]')
        video_input.send_keys(video_path)

        # Wait for the video to upload
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//span[@data-icon="send"]')))

        # Send the video
        send_button = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
        send_button.click()

        print(f"Video sent to {phone_number}")
        time.sleep(8);

    except TimeoutException:
        print(f"Timeout waiting for element in WhatsApp window for {phone_number}")
    except Exception as e:
        print(f"Failed to send video to {phone_number}: {e}")

# Load the Excel file
df = pd.read_excel(excel_file_path, engine='openpyxl')

# Extract phone numbers
phone_numbers = df['Mobile Phone'].astype(str).tolist()
print("List of ph number",len(phone_numbers))
print("Phone numbers - ", phone_numbers)

# Iterate over each phone number and send the video
for phone in phone_numbers:
    send_video_via_whatsapp(f"+{phone}", video_file_path)

