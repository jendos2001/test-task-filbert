from parcer.parsers import Worker

site_url = 'https://www.saucedemo.com/'
login = 'standard_user'
password = 'secret_sauce'
download_dir = '/Downloads'
items_count = 3
first_name = 'Evgeny'
last_name = 'Kurshev'
postal_code = '187550'


if __name__ == '__main__':
    worker = Worker(site_url, login, password, download_dir, 
                    items_count, first_name, last_name, postal_code)
    worker.start()

