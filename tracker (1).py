from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
import time

excel_file_path = "Input.xlsx"  
workbook = openpyxl.load_workbook(excel_file_path)  
add_note = "Current CST-  7:30 AM- no driver found due to capacity constraints "  

sheet = workbook["Sheet1"]  

executed_count = 0  
exempted_count = 0  

def wait_and_click(driver, locator):
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(locator)).click()

def wait_and_send_keys(driver, locator, keys):
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located(locator)).send_keys(keys)

def check_tooltip_value_and_print(driver):
    tooltip_elements = driver.find_elements(By.XPATH, '//span[@class="inline-tooltip-wrapper"]')
    for tooltip_element in tooltip_elements:
        span_text = tooltip_element.text.strip()  
        if span_text in ["AZNG", "RLB1"]:  
            print(f"Found tooltip with text: {span_text}. Skipping to the next element in Excel.")
            return True  
    print("Tooltip element not found or no relevant text found")  
    return False

def execute_workflow(driver, value):
    global executed_count
    executed_count += 1  

    wait_and_click(driver, (By.XPATH, '//span[@class="fa fa-comments "]'))  
    time.sleep(2)
    wait_and_click(driver, (By.XPATH, '//input[@name="addCommentFake"]'))  
    time.sleep(2)
    wait_and_send_keys(driver, (By.XPATH, '//textarea[@name="addComment"]'), add_note)  
    time.sleep(2)
    wait_and_click(driver, (By.XPATH, '//button[@name="saveComment"]'))  
    time.sleep(2)
    wait_and_click(driver, (By.XPATH, '//button[@title="Close"]'))  
    time.sleep(2)
    wait_and_click(driver, (By.XPATH, '//span[@class="fa fa-pencil"]'))  
    time.sleep(2)
    wait_and_click(driver, (By.XPATH, '//span[@class="icon-holder fa fa-times"]'))  
    time.sleep(2)
    wait_and_click(driver, (By.XPATH, '//span[text()=" -- Select a cancellation reason -- "]'))  
    time.sleep(2)
    wait_and_click(driver, (By.XPATH, '//li[text()="TPC shipper closed"]'))  
    wait_and_click(driver, (By.XPATH, '//button[text()="Submit"]'))  
    time.sleep(2)

    rows = driver.find_elements(By.XPATH, '//table[@class="dataTable"]/tbody/tr')
    for row in rows:
        row_data = [col.text for col in row.find_elements(By.TAG_NAME, 'td')]  
        #r_sheet.append([value] + row_data)  
        workbook.save(excel_file_path)  

def main():
    global exempted_count

    for cell in sheet['A']:
        value = cell.value  
        if not value:
            continue  

        print(f"Processing Value: {value}")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))  
        driver.get("https://trans-logistics.amazon.com/fmc/execution")  
        driver.maximize_window()  

        wait_and_click(driver, (By.XPATH, '/html/body/div[2]/div[2]/div/div/div[2]/div/div[3]/div/div/div[1]/div/form/div[2]/div[1]/div/div[1]/a/h5'))
        wait_and_send_keys(driver, (By.XPATH, '/html/body/div[2]/div[2]/div/div/div[2]/div/div[3]/div/div/div[1]/div/form/div[2]/div[1]/div/div[2]/div/div/textarea'), value + Keys.CONTROL + Keys.RETURN)
        time.sleep(2)

        if check_tooltip_value_and_print(driver):
            exempted_count += 1  
            driver.quit()  
            continue

        execute_workflow(driver, value)
        driver.quit()  

    print(f"Executed count: {executed_count}")
    print(f"Exempted count: {exempted_count}")

if __name__ == "__main__":
    main()  
    print("Browser closed. Press Enter to exit...")
    input()  
