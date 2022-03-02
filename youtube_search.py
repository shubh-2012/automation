import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome( )
driver.get('http://www.youtube.com/')
#searchbox = driver.find_element_by_name("search_query")
searchbox = driver.find_element(By.NAME,"search_query")
searchbox.send_keys('selenium')
searchbox.submit()
time.sleep(15)
driver.quit()