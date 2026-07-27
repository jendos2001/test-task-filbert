from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from random import sample, randint
from time import sleep
import os



site_url = 'https://www.saucedemo.com/'
login = 'standard_user'
password = 'secret_sauce'
download_dir = os.path.abspath(os.getcwd()) + '/Downloads'

options = Options()
prefs = {
        "profile.password_manager_leak_detection": False,
        "download.default_directory": download_dir,
        "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)
#options.add_argument('--headless=new')

driver = webdriver.Chrome(options=options)
driver.get(site_url)

#!--------task1--------!
driver.find_element(By.ID, 'user-name').send_keys(login)
driver.find_element(By.ID, 'password').send_keys(password)
driver.find_element(By.ID, 'login-button').click()

#!--------task2-3--------!
WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'root')))

items_count = 3
items = driver.find_element(By.CLASS_NAME, 'inventory_list').find_elements(By.CLASS_NAME, 'inventory_item')
buy_items_indecies = sorted(sample(range(len(items)), items_count))
k = 0

database = []
for index, item in enumerate(items):
    if index == buy_items_indecies[k]:
        title = item.find_element(By.CLASS_NAME, 'inventory_item_name').text 
        description = item.find_element(By.CLASS_NAME, 'inventory_item_desc').text 
        price = float(item.find_element(By.CLASS_NAME, 'inventory_item_price').text[1:])
        item.find_element(By.CLASS_NAME, 'btn').click()
        database.append((title, description, price))
        k += 1
        if k == items_count:
            break

#!--------task4--------!
delete_item_index = randint(0, items_count - 1)
delete_item = database.pop(delete_item_index)

#!--------task5--------!
driver.find_element(By.CLASS_NAME, 'shopping_cart_link').click()

#!--------task6--------!
WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'root')))
cart_items = driver.find_element(By.CLASS_NAME, 'cart_list').find_elements(By.CLASS_NAME, 'cart_item')
for item in cart_items:
    if item.find_element(By.CLASS_NAME, 'inventory_item_name').text == delete_item[0]:
        item.find_element(By.CLASS_NAME, 'btn').click()
        break

#!--------task7--------!
driver.find_element(By.ID, 'checkout').click()

#!--------task8--------!
WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'root')))
first_name = 'Evgeny'
last_name = 'Kurshev'
postal_code = '187550'
driver.find_element(By.ID, 'first-name').send_keys(first_name)
driver.find_element(By.ID, 'last-name').send_keys(last_name)
driver.find_element(By.ID, 'postal-code').send_keys(postal_code)
driver.find_element(By.ID, 'continue').click()

#!--------task9--------!
WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'root')))
info_item_block = driver.find_element(By.CLASS_NAME, 'summary_info')
info_items = info_item_block.find_elements(By.CLASS_NAME, 'summary_value_label')
payment_information = info_items[0].text
shipping_information = info_items[1].text
tax = float(driver.find_element(By.CLASS_NAME, 'summary_tax_label').text.split()[1][1:])
total = float(driver.find_element(By.CLASS_NAME, 'summary_total_label').text.split()[1][1:])
info_item_block.find_element(By.ID, 'finish').click()

#!--------task10--------!
WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'root')))
driver.find_element(By.ID, 'generate-pdf-order').click()


sleep(30)

