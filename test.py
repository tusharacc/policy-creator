import CH_data as CH
import time,traceback,names
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from openpyxl import load_workbook

#the list will contain all the values from excel
input_test_policies = []
print (CH.states['AZ'])
wb = load_workbook('Test-Case.xlsm')
ws = wb['Sheet1']

for i in range(2,3):
    CH.value_read['state'] = ws['A'+str(i)].value
    CH.value_read['business_segment'] = ws['B'+str(i)].value
    CH.value_read['business_type'] = ws['C'+str(i)].value
    CH.value_read['curr_coverage'] = ws['D'+str(i)].value
    CH.value_read['business_ownership'] = ws['G'+str(i)].value
    CH.value_read['business_start_date'] = ws['H'+str(i)].value
    CH.value_read['employees_count'] = ws['J'+str(i)].value
    CH.value_read['annual_payroll'] = ws['K'+str(i)].value
    CH.value_read['annual_gross_sales'] = ws['L'+str(i)].value
    CH.value_read['footage'] = ws['M'+str(i)].value
    CH.value_read['address_line'] = ws['S'+str(i)].value
    CH.value_read['city'] = ws['U'+str(i)].value
    CH.value_read['zip_code'] = ws['W'+str(i)].value
    input_test_policies.append(CH.value_read)
    
#Constants used in the program
COVERAGE_ID = ['product_codes__general_liability',
               'product_codes__professional_liability',
               'product_codes__workers_compensation',
               'product_codes__commercial_auto',
               'product_codes__bop']
COVERAGE_XPATH = {
    'gl':'//*[@id="field_for_product_codes__general_liability"]/label',
    'pl':'//*[@id="field_for_product_codes__professional_liability"]/label',
    'wc':'//*[@id="field_for_product_codes__workers_compensation"]/label',
    'ca':'//*[@id="field_for_product_codes__commercial_auto"]/label',
    'bp':'//*[@id="field_for_product_codes__bop"]/label'
}

EMAIL = 'bill.clinton@whitehouse.com'
ADDITIONAL_QUESTIONS = ['Personal Training (Health And Fitness)']

for policy in input_test_policies:
    COMPANY_NAME = names.get_full_name()
    FIRST_NAME = names.get_first_name()
    LAST_NAME = names.get_last_name()
    
    curr_url = ''
    first_pass = True
    
    driver = webdriver.PhantomJS()
    driver.delete_all_cookies()
    driver.start_session(DesiredCapabilities.PHANTOMJS)
    driver.implicitly_wait(10)
    driver.set_window_size(1120, 550)
    try:
        #HOME PAGE
        driver.get("https://chubb:ChubbCH34@psc-chubb-sit.coverhound.us/")
        select_state = Select(driver.find_element_by_id('state_abbrev'))
        select_state.select_by_visible_text(states[policy['state']])
        time.sleep(5)
        select_business_segment = Select(driver.find_element_by_id('business_segment_id'))
        select_business_segment.select_by_visible_text(policy['business_segment'])
        time.sleep(5)
        select_business_type = Select(driver.find_element_by_id('business_type_id'))
        select_business_type.select_by_visible_text(policy['business_type'])
        
        driver.save_screenshot(COMPANY_NAME+ '_home_page_screenshot.png')
        driver.find_element_by_xpath('//*[@id="chubb_commercial_entry_form"]/div/button').click()
        time.sleep(10)
        #wait = WebDriverWait(driver, 10).until( EC.element_to_be_clickable((By.ID, 'product_codes__bop')))
        print (driver.current_url)
        
        driver.save_screenshot(COMPANY_NAME+ '_home_page_screenshot_1.png')

    except Exception as e:
        print (e)
        print (driver.current_url)
        driver.save_screenshot('error_screenshot.png')
        traceback.print_exc()
    finally:
        driver.quit()