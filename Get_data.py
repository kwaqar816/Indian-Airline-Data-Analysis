from traceback import print_tb
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
import sys
import random

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
path = r'C:\\chromedriver.exe'
driver = webdriver.Chrome(chrome_options=options, executable_path=path)
driver.get('https://www.dgca.gov.in/digigov-portal/?page=4268/4207/servicename')
time.sleep(random.randint(3, 7))
current_url = driver.current_url
quarter_links = []
tried_times = 1


while tried_times <= 4:
    try:
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "a[data-url*='quat']")))
        break
    except Exception as e:
        this_exception = 'Error on line {} {} {}'.format(
            sys.exc_info()[-1].tb_lineno, type(e).__name__, e)
        print(this_exception)
        driver.refresh()
        tried_times += 1


quarters = driver.find_elements_by_css_selector("a[data-url*='quat']")
for i in quarters:
    quarter_links.append(i.get_attribute('data-url'))
print('All quarter links extracted')
print()
for i in range(0, 4):
    try:
        temp_url = current_url.split('=')
        main_url = temp_url[0] + '=' + quarter_links[i] + '&main'+temp_url[1]
        time.sleep(1)
        print()
        print('Opening new tab')
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[1])
        driver.get(main_url)
        print()
        time.sleep(random.randint(2, 4))
        tried_times = 1
        while tried_times <= 4:
            try:
                WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, 'td>a[data-url*=".xlsx"]')))
                break
            except Exception as e:
                this_exception = 'Error on line {} {} {}'.format(
                    sys.exc_info()[-1].tb_lineno, type(e).__name__, e)
                print(this_exception)
                driver.refresh()
                tried_times += 1
        Excel_buttons = driver.find_elements_by_css_selector(
            'td>a[data-url*=".xlsx"]')
        for each_excel in Excel_buttons:
            each_excel.click()
            time.sleep(2)
            print('File downloaded successfully')
        print()
        print('Closing current tab')
        time.sleep(2)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
    except Exception as e:
        this_exception = 'Error on line {} {} {}'.format(
            sys.exc_info()[-1].tb_lineno, type(e).__name__, e)
        print(this_exception)


print('Program ended.......................................')
time.sleep(0.5)
print('Program ended.......................................')
time.sleep(0.5)
print('Program ended.......................................')
time.sleep(0.5)
print('Program ended.......................................')
time.sleep(0.5)
print('Program ended.......................................')
print()
print()
print()
time.sleep(2)
print('CLOSING CHROME DRIVER')
driver.close()
print('*****THANK YOU GOOD BYE*****')
print()
print()
print()
