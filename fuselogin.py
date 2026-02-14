import json
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import traceback
def log_message(log_file, message):
    with open(log_file, 'a', encoding='utf-8') as log:
        log.write(message + '\n')
    print(message)
def fuse_login(driver, credentials, log_file):
    FuseLogin = False
    try:
        driver.get(credentials['FuseURL'])
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(3)
        
        try:
            WebDriverWait(driver, 5).until(EC.alert_is_present())
            time.sleep(1)
            alert = driver.switch_to.alert
            alert.accept()
            print("Popup clicked")
        except TimeoutException:
            print("No alert appeared.")
        driver.refresh()
        try:
            driver.refresh()
            time.sleep(3)
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="myNavbar"]/ul/li[2]/a')))
            log_message(log_file, "Fuse login was successful (session already active).")
            driver.find_element(By.XPATH, '//*[@id="myNavbar"]/ul/li[2]/a')
            FuseLogin = True
        except:
            print("Fuse Login Page is visibled")
            driver.refresh()
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="form"]/div[1]/input')))
            driver.find_element(By.XPATH, '//*[@id="form"]/div[1]/input').send_keys(credentials['Fusername'])
            time.sleep(1)
            driver.find_element(By.ID, 'multi_user_timeout_pin').send_keys(credentials['Fpassword'])
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="form"]/div[3]/a').click()
            time.sleep(3)

            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@href='/mobile/mobile-login.php']")))
            log_message(log_file, "Fuse login was successful.")
            FuseLogin = True

    except Exception as e:
        log_message(log_file, f"Fuse login was Unsuccessful: {str(e)}")
        log_message(log_file, traceback.format_exc())

    return FuseLogin
    