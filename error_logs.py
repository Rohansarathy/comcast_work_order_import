import os
import time

from openpyxl import Workbook, load_workbook
from Sendmail import Sendmail
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
 
 
def log_message(log_file, message):
    with open(log_file, 'a', encoding='utf-8') as log:
        log.write(message + '\n')
    print(message)
    
def check_error_logs(driver, date_folder, credentials, department,  log_file):
    log_message(log_file, "check for Total job attempt and import")
    time.sleep(8)
    
    ####Import Type Dropdown####
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="Import_type"]')))
    dropdown_element = driver.find_element(By.XPATH, '//*[@id="Import_type"]')
    select = Select(dropdown_element)
    select.select_by_value("2")
    time.sleep(2)
    ####Import Template Dropdown####
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="templateid"]')))
    dropdown_element = driver.find_element(By.XPATH, '//*[@id="templateid"]')
    select = Select(dropdown_element)
    select.select_by_value("14")
    time.sleep(6)

    ###Import Status Table####
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "(//tr[td/input[@id='wobr1' and contains(@value, '.xlsx')]])[1]")))
    schedule_date_row = driver.find_element(By.XPATH, "//table[@id='exampleImport']//tr[td[8][normalize-space()='done']][1]")
    schedule_date = schedule_date_row.find_element(By.XPATH, "(//table[@id='exampleImport']//tbody/tr/td[4])[1]")
    date_str_1 = schedule_date.text.strip()
    print(f"Upload Date:{date_str_1}")
    upload_date = datetime.strptime(date_str_1, "%Y-%m-%d")
    print("Parsed date:", upload_date)
    formatted_date = upload_date.strftime("%d/%m/%Y")
    print("Formatted date:", formatted_date)
    
    today_date = datetime.now().strftime("%d/%m/%Y")
    print("Today date:", today_date)
    print(f"Schedule Date:{date_str_1}")
    log_message(log_file, f"Schedule Date:{date_str_1}")
    if today_date == formatted_date:
        ####Status####
        # status = driver.find_element(By.XPATH, "(//table[@id='exampleImport']//tbody/tr/td[8])[1]")
        # driver.execute_script("arguments[0].scrollIntoView({block:'center'});", status)
        # status_result = status.text.strip()
        # print(f"Status Result:{status_result}")
        # time.sleep(900)
        # driver.execute_script("location.reload(true);")
        MAX_WAIT_TIME = 900   
        POLL_INTERVAL = 10    
        start_time = time.time()
        while True:
            time.sleep(7)
            WebDriverWait(driver, 50).until(EC.visibility_of_element_located((By.XPATH,  "(//table[@id='exampleImport']//tbody/tr/td[8])[1]")))
            status = driver.find_element(By.XPATH, "(//table[@id='exampleImport']//tbody/tr/td[8])[1]")
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", status)
            status_result = status.text.strip().lower()
            print(f"Current Status: {status_result}")
            if status_result == "done":
                print("Status is DONE. Proceeding to next step...")
                break
            if time.time() - start_time > MAX_WAIT_TIME:
                raise TimeoutError("Status did not become 'done' within the expected time.")
            time.sleep(POLL_INTERVAL)
            driver.execute_script("location.reload(true);")
        if status_result == "done":
            print("Status is DONE. Proceeding to next step...")
            ####Import Type Dropdown####
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="Import_type"]')))
            dropdown_element = driver.find_element(By.XPATH, '//*[@id="Import_type"]')
            select = Select(dropdown_element)
            select.select_by_value("2")
            time.sleep(2)
            ####Import Template Dropdown####
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="templateid"]')))
            dropdown_element = driver.find_element(By.XPATH, '//*[@id="templateid"]')
            select = Select(dropdown_element)
            select.select_by_value("14")
            time.sleep(6)
            print(f"Status Result:{status_result}")
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "(//tr[td/input[@id='wobr1' and contains(@value, '.xlsx')]])[1]")))
            upload_date_row = driver.find_element(By.XPATH, "//table[.//th[normalize-space(.)='Upload Date and Time']]//tbody/tr[1]")
            upload_date_input = upload_date_row.find_element(By.XPATH, "(//table[.//th[normalize-space(.)='Upload Date and Time']]//tbody/tr[1]//input[@disabled])[3]")
            date_str_2 = upload_date_input.get_attribute("value")
            print(f"Upload Date String:{date_str_2}")
            upload_date = datetime.strptime(date_str_2, "%Y-%m-%d %H:%M:%S")
            print("Parsed date:", upload_date)
            formatted_date = upload_date.strftime("%d/%m/%Y")
            print("Formatted date:", formatted_date)
            print("Today date:", today_date)
            today_date = datetime.now().strftime("%d/%m/%Y")
            print(f"Upload_date:{date_str_2}")
            log_message(log_file, f"Upload_date:{date_str_2}")
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", upload_date_row)
            if today_date == formatted_date:
                ####Total Jobs Attempted####
                total_jobs_attempt = driver.find_element(By.XPATH, "(//table[.//th[normalize-space(.)='Upload Date and Time']]//tbody/tr[1]//input[@disabled and not(@id)])[5]")
                get_value_jobs_attempt = total_jobs_attempt.get_attribute("value")
                print(f"Total Jobs attemp:{get_value_jobs_attempt}")
                ####Import Status####
                import_status_xpath = driver.find_element(By.XPATH, "(//table[.//th[normalize-space(.)='Upload Date and Time']]//tbody/tr[1]//input[@disabled and not(@id)])[6]")
                import_status = import_status_xpath.get_attribute("value")
                print(f"import Status:{import_status}")
                if import_status == "Done":
                    ####Total Imported####
                    total_imported = driver.find_element(By.XPATH, "(//table[.//th[normalize-space(.)='Upload Date and Time']]//tbody/tr[1]//input[@disabled and not(@id)])[7]")
                    total_imported  = total_imported.get_attribute("value")
                    print(f"Total Jobs import:{total_imported }")
                    if total_imported  == '0':
                        total_updated_xpath = driver.find_element(By.XPATH,"(//table[.//th[normalize-space(.)='Upload Date and Time']]//tbody/tr[1]//input[@disabled and not(@id)])[8]")
                        total_updated = total_updated_xpath.get_attribute("value")
                        print(f"Total Updated: {total_updated}")
                        if get_value_jobs_attempt == total_updated:
                            log_message(log_file, "Total Jobs Attempted matches Total Updated")
                        else:
                            log_message(log_file, "Total Jobs Attempted not matched Total Updated")
                            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//table[.//th[normalize-space(.)='Upload Date and Time']]//tbody/tr[1]//a[normalize-space(text())='Error Log']")))
                            driver.find_element(By.XPATH, "//table[.//th[normalize-space(.)='Upload Date and Time']]//tbody/tr[1]//a[normalize-space(text())='Error Log']").click()
                            time.sleep(5)
                            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//h2[normalize-space(text())='Failed Import Results']")))
                            print("Error Log page now visible.")
                            dropdown_element = driver.find_element(By.XPATH, "//select[@name='example3_length']")
                            select = Select(dropdown_element)
                            select.select_by_value("-1")
                            time.sleep(6)
                            screenshot_path = os.path.join(date_folder, "Error_log.png")
                            driver.save_screenshot(screenshot_path)
                            time.sleep(6)
                            driver.find_element(By.XPATH, "//a[@class='btn btn-default' and normalize-space(text())='X']").click()
                            time.sleep(3)
                            #### Mail Logic ####
                            recipient_emails = credentials['ybotID']
                            cc_emails = credentials['ybotID']
                            subject = f"An Exception occurred during Work Order Import for {department} Department"
                            body_message = f"An Exception occurred during Work Order Import for {department} Department trying to import the Work Order file uploaded. Please find the attached Error Log screenshot for more details."
                            body_message1 = ""
                            attachment_path = date_folder + "\\Error_log.png"
                            try:
                                Sendmail(recipient_emails, cc_emails, subject, body_message, body_message1, attachment_path)
                                log_message(log_file, "Mail sent Successfully for Error Logs.")
                            except Exception:
                                log_message(log_file, "Mail not sent for Error Logs.")
                    else:
                        total_imported_xpath = driver.find_element(By.XPATH,"(//table[.//th[normalize-space(.)='Upload Date and Time']]//tbody/tr[1]//input[@disabled and not(@id)])[7]")
                        total_import = total_imported_xpath.get_attribute("value")
                        print(f"Total Updated: {total_import}")
                        if get_value_jobs_attempt == total_import:
                            log_message(log_file, "Total Jobs Attempted matches Total Imported")
                        else:
                            log_message(log_file, "Total Jobs Attempted not matched Total Imported")
                            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//table[.//th[normalize-space(.)='Upload Date and Time']]//tbody/tr[1]//a[normalize-space(text())='Error Log']")))
                            driver.find_element(By.XPATH, "//table[.//th[normalize-space(.)='Upload Date and Time']]//tbody/tr[1]//a[normalize-space(text())='Error Log']").click()
                            time.sleep(5)
                            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//h2[normalize-space(text())='Failed Import Results']")))
                            print("Error Log page now visible.")
                            dropdown_element = driver.find_element(By.XPATH, "//select[@name='example3_length']")
                            select = Select(dropdown_element)
                            select.select_by_value("-1")
                            time.sleep(6)
                            screenshot_path = os.path.join(date_folder, "Error_log.png")
                            driver.save_screenshot(screenshot_path)
                            time.sleep(4)
                            driver.find_element(By.XPATH, "//a[@class='btn btn-default' and normalize-space(text())='X']").click()
                            time.sleep(3)
                            #### Mail Logic ####
                            recipient_emails = credentials['ybotID']
                            cc_emails = credentials['ybotID']
                            subject = f"An Exception occurred during Work Order Import for {department} Department"
                            body_message = f"An Exception occurred during Work Order Import for {department} Department trying to import the Work Order file uploaded. Please find the attached Error Log screenshot for more details."
                            body_message1 = ""
                            attachment_path = date_folder + "\\Error_log.png"
                            try:
                                Sendmail(recipient_emails, cc_emails, subject, body_message, body_message1, attachment_path)
                                log_message(log_file, "Mail sent Successfully for Error Logs.")
                            except Exception:
                                log_message(log_file, "Mail not sent for Error Logs.")
                else:
                    log_message(log_file, "Import Status not matched Done")
            else:
                log_message(log_file, "dates are not matched")
        else:
            log_message(log_file, "Status not done yet.")
    else:
        log_message(log_file, "today dates are not matched")