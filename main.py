from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import time
from dotenv import load_dotenv
import os
from datetime import datetime, timedelta

# Load environment variables from .env file
load_dotenv()

def extract_otp_from_email_content(email_content):
    # Use regex to find 6-digit OTP
    otp_match = re.search(r'\b(\d{6})\b', email_content)
    if otp_match:
        return otp_match.group(1)
    return None

def login_clic(driver, clic_id, clic_password):
    driver.get("https://clic.mmu.edu.my/psc/csprd_320/EMPLOYEE/SA/s/WEBLIB_PTBR.ISCRIPT1.FieldFormula.IScript_StartPage")
    
    # Enter User ID
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "userid"))
    ).send_keys(clic_id)
    
    # Enter Password
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "pwd"))
    ).send_keys(clic_password)
    
    # Click 'Sign In' button
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, "Submit"))
    ).click()
    
    # Wait for OTP input to appear
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "otp"))
    )
    print("CLIC login successful, waiting for OTP!")

def validate_otp(driver, otp):
    # Enter OTP
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "otp"))
    ).send_keys(otp)
    
    # Click 'Validate OTP' button
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ps_submit_button"))
    ).click()
    
    # Wait for validation to complete
    WebDriverWait(driver, 10).until(
        EC.url_contains("EMPLOYEE/SA/c/")
    )
    print("OTP validated successfully!")
    
    # Wait for 10 seconds before closing
    time.sleep(10)

def wait_for_new_email(driver, sender_email):
    while True:
        # Refresh the inbox
        driver.refresh()
        
        # Wait for the inbox to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "ms-FocusZone"))
        )
        
        # Check for new email from the sender
        try:
            first_email = driver.find_element(By.XPATH, f"//div[contains(@class, 'jGG6V') and .//span[text()='{sender_email}']]")
            return first_email
        except Exception:
            print("No new email found, waiting for 30 seconds...")
            time.sleep(30)

def login_outlook(driver, email, password):
    # Open a new tab for Outlook
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    
    # Go to Outlook
    driver.get("https://outlook.office.com/mail/")
    
    # Enter email
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "i0116"))
    ).send_keys(email, Keys.RETURN)
    
    # Enter password
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "i0118"))
    ).send_keys(password, Keys.RETURN)
    
    # Click 'Sign In' button
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Sign in']"))
    ).click()
    
    # Handle "Stay signed in?" prompt if it appears
    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "idBtn_Back"))
        ).click()
    except Exception:
        pass
    
    # Wait for the inbox to load
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CLASS_NAME, "ms-FocusZone"))
    )
    print("Login successful!")
    
    # Wait for a new email from 'otp-noreply@mmu.edu.my'
    first_email = wait_for_new_email(driver, "otp-noreply@mmu.edu.my")
    first_email.click()
    
    # Extract email content
    email_content_element = driver.find_element(By.XPATH, "//div[contains(@class, 'Zgp3k')]//span")
    email_content = email_content_element.text
    
    # Extract OTP
    otp = extract_otp_from_email_content(email_content)
    
    if otp:
        print(f"OTP found: {otp}")
        # Switch back to the CLIC tab
        driver.switch_to.window(driver.window_handles[0])
        # Validate OTP on CLIC
        validate_otp(driver, otp)
    else:
        print("No OTP found in the email")
    
    return otp

def main():
    service = Service("//Users/leeweixuan/Desktop/SassyNic-selenium/driver/chromedriver")
    driver = webdriver.Chrome(service=service)
    try:
        # Login to CLIC first
        login_clic(driver, os.getenv("CLIC_ID"), os.getenv("CLIC_PASSWORD"))
        
        # Login to Outlook and get OTP
        otp = login_outlook(driver, os.getenv("EMAIL"), os.getenv("PASSWORD"))
        
        print(f"Result: {otp}")
    except Exception as e:
        print(f"Login failed: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()