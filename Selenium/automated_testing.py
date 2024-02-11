from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

# Set up the Selenium driver
driver = webdriver.Chrome()

# Navigate to the login page
driver.get("https://www.stg-lanstad.com/")

# Find the username and password input fields
username_input = driver.find_element(By.ID, "username")
password_input = driver.find_element(By.ID, "password")

# Enter your login credentials
username_input.send_keys("techsupport@deantaglobal.com")
password_input.send_keys("123456")

# Submit the login form
password_input.send_keys(Keys.RETURN)

search = driver.find_element(By.XPATH, "//div[@id='root']//[@class='search']")
search.send_keys("AI-NAPH")

# Interact with the dashboard or perform any necessary automated tests
# ...

# Close the Selenium driver
# driver.quit()
