# importing webdriver from selenium
from selenium import webdriver
import time
from PIL import Image

# Here Chrome  will be used

driver = webdriver.Chrome()

# URL of website
url = "https://www.youtube.com/"

# Opening the website
driver.get(url)
driver.execute_script("document.body.style.zoom= '50%' ")
time.sleep(2)
driver.save_screenshot("image.png")

# Loading the image
image = Image.open("image.png")

# Showing the image
image.show()