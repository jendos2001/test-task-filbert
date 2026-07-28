from parcer.parsers import Parser
import os

site_url = os.environ.get('SITE_URL', 'https://www.saucedemo.com/')
db_name = os.environ.get('DB_NAME', 'orders.db')
log_filename = os.environ.get('LOG_FILENAME', 'program.log')
xlsx_filename = os.environ.get('XLSX_FILENAME', 'Report.xlsx')
login = os.environ.get('LOGIN', 'standard_user')
password = os.environ.get('PASSWORD', 'secret_sauce')
download_dir = os.environ.get('DOWNLOAD_DIR', '/Downloads')
items_count = int(os.environ.get('ITEMS_COUNT', 3))
first_name = os.environ.get('FIRST_NAME', 'Evgeny')
last_name = os.environ.get('LAST_NAME', 'Kurshev')
postal_code = os.environ.get('POSTAL_CODE', '187550')


if __name__ == '__main__':
    parser = Parser(site_url, login, password, download_dir, items_count,
                    first_name, last_name, postal_code,
                    db_name, log_filename, xlsx_filename)
    parser.start()
