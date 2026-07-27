from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep



site_url = 'https://www.saucedemo.com/'
login = 'standard_user'
password = 'secret_sauce'

options = Options()
#options.add_argument('--headless=new')

driver = webdriver.Chrome(options=options)
driver.get(site_url)

#!--------task1-----------!
driver.find_element(By.ID, 'user-name').send_keys(login)
driver.find_element(By.ID, 'password').send_keys(password)
driver.find_element(By.ID, 'login-button').click()
sleep(5)

#print(main_page)