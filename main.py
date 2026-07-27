from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from random import sample
from time import sleep



site_url = 'https://www.saucedemo.com/'
login = 'standard_user'
password = 'secret_sauce'

options = Options()
#options.add_argument('--headless=new')

driver = webdriver.Chrome(options=options)
driver.get(site_url)

#!--------task1--------!
driver.find_element(By.ID, 'user-name').send_keys(login)
driver.find_element(By.ID, 'password').send_keys(password)
driver.find_element(By.ID, 'login-button').click()

#!--------task2-3--------!
WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'root')))

try:
    alert_obj = driver.switch_to.alert 
    alert_obj.accept()

except:
    pass

items_count = 3
items = driver.find_element(By.CLASS_NAME, 'inventory_list').find_elements(By.CLASS_NAME, 'inventory_item')
buy_items = sorted(sample(range(len(items)), items_count))
k = 0
database = []
for index, item in enumerate(items):
    if index == buy_items[k]:
        title = item.find_element(By.CLASS_NAME, 'inventory_item_name').text 
        description = item.find_element(By.CLASS_NAME, 'inventory_item_desc').text 
        price = float(item.find_element(By.CLASS_NAME, 'inventory_item_price').text[1:])
        item.find_element(By.CLASS_NAME, 'btn').click()
        database.append((title, description, price))
        k += 1
        if k == items_count:
            break

sleep(30)

