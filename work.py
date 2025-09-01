from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Initialize the browser
driver = webdriver.Chrome()
driver.get("https://www.tiktok.com/")

# Wait and manually log in if needed
time.sleep(30)  # Adjust for manual login

# Go to a specific user's profile
driver.get("https://www.tiktok.com/@xyz")  # Replace with target username

time.sleep(5)

# Click on the latest video
videos = driver.find_elements(By.CSS_SELECTOR, 'a[href*="https://vt.tiktok.com/ZSM1k9Sva/"]')
if videos:
    videos[0].click()

time.sleep(5)

# Find the comment box and type a comment
comment_box = driver.find_element(By.CSS_SELECTOR, 'div[contenteditable="true"]')
comment_box.click()
comment_box.send_keys("Fraud User! Fraud video !" + Keys.RETURN)

time.sleep(5)  # Allow time for comment to post
driver.quit()

