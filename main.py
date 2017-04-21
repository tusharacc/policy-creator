import time,traceback,names,CH_data,user_detail as user
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from openpyxl import load_workbook


#Link of Pages - 
#https://psc-chubb-sit.coverhound.us/business-info
#https://psc-chubb-sit.coverhound.us/business-operations
#https://psc-chubb-sit.coverhound.us/contact
#https://psc-chubb-sit.coverhound.us/coverage-detail/bop

#XPATH for Coverages
#gl = //*[@id="field_for_product_codes__general_liability"]/label
#pl = //*[@id="field_for_product_codes__professional_liability"]/label
#wc = //*[@id="field_for_product_codes__workers_compensation"]/label
#ca = //*[@id="field_for_product_codes__commercial_auto"]/label

def checkPageTransition(old,new,msg):
    if old == new:
        raise ValueError(msg)

wb = load_workbook('Test-Case.xlsx')
ws = wb['Sheet1']

coverage_xpath = {
    'gl':'//*[@id="field_for_product_codes__general_liability"]/label',
    'pl':'//*[@id="field_for_product_codes__professional_liability"]/label',
    'wc':'//*[@id="field_for_product_codes__workers_compensation"]/label',
    'ca':'//*[@id="field_for_product_codes__commercial_auto"]/label',
    'bp':'//*[@id="field_for_product_codes__bop"]/label'
}
ADDITIONAL_QUESTIONS = ['Personal Training (Health And Fitness)']
COMPANY_NAME = names.get_full_name()
FIRST_NAME = names.get_first_name()
LAST_NAME = names.get_last_name()
EMAIL = 'bill.clinton@whitehouse.com'
business_segment = 'Healthcare, Therapy & Fitness'
business_type = 'Personal Training (Health And Fitness)'
employees_count = 15
annual_payroll = 15222
annual_gross_sales = 785963
footage = 5000
address_line = '4041 North Central Avenue'
city = 'Phoenix'
state = 'AZ'
zip_code = 85012

prev_coverage = 'bp'
curr_coverage = 'bp'
curr_url = ''
first_pass = True

driver = webdriver.PhantomJS()
driver.delete_all_cookies()
driver.start_session(DesiredCapabilities.PHANTOMJS)
driver.implicitly_wait(10)
driver.set_window_size(1120, 550)
try:
    #HOME PAGE
    driver.get('https://'+user.USERNAME+':'+user.PASSWORD+'@psc-chubb-sit.coverhound.us/')
    select_state = Select(driver.find_element_by_id('state_abbrev'))
    select_state.select_by_visible_text('Arizona')
    time.sleep(5)
    select_business_segment = Select(driver.find_element_by_id('business_segment_id'))
    select_business_segment.select_by_visible_text(business_segment)
    time.sleep(5)
    select_business_type = Select(driver.find_element_by_id('business_type_id'))
    select_business_type.select_by_visible_text(business_type)
    driver.save_screenshot(COMPANY_NAME+ '_home_page_screenshot.png')
    driver.find_element_by_xpath('//*[@id="chubb_commercial_entry_form"]/div/button').click()
    time.sleep(10)
    #wait = WebDriverWait(driver, 10).until( EC.element_to_be_clickable((By.ID, 'product_codes__bop')))
    print (driver.current_url)
    #BUSINESS INFO
    if driver.current_url == 'https://'+user.USERNAME+':'+user.PASSWORD+'@psc-chubb-sit.coverhound.us/business-info':
        curr_url = driver.current_url
        if prev_coverage != curr_coverage:
            driver.find_element_by_xpath(coverage_xpath[curr_coverage]).click()
            prev_coverage = curr_coverage
        comp_name = driver.find_element_by_id('business_name')
        comp_name.clear()
        comp_name.send_keys(COMPANY_NAME)
        email = driver.find_element_by_id('email')
        email.clear()
        email.send_keys((Keys.CONTROL, "a"))
        email.send_keys(EMAIL)
        driver.save_screenshot(COMPANY_NAME+ '_business_info_screenshot.png')
        driver.find_element_by_xpath('//*[@id="commercial-app"]/div/div[2]/div[2]/div/div[2]/form/div[1]/div/div/button').click()
        print (driver.current_url)
        time.sleep(10)
        checkPageTransition(curr_url,driver.current_url,'Error in BusinessInfo')
            
    
    #BUSINESS OPERATION
    if driver.current_url == 'https://'+user.USERNAME+':'+user.PASSWORD+'@psc-chubb-sit.coverhound.us/business-operations':
        curr_url = driver.current_url
        select_business_structure = Select(driver.find_element_by_xpath('//*[@id="CH_024"]'))
        select_business_structure.select_by_visible_text('Individual/Sole Proprietor')
        driver.find_element_by_id('CH_105').send_keys('04/04/2016')
        if first_pass:
            driver.find_element_by_xpath('//*[@id="field_for_CH_026"]/div[1]/div[2]/label').click()
        number_of_employees = driver.find_element_by_id('CH_029')
        number_of_employees.clear()
        number_of_employees.send_keys(employees_count)
        payroll = driver.find_element_by_id('CH_030')
        payroll.clear()
        payroll.send_keys(annual_payroll)
        projected_sales = driver.find_element_by_id('CH_031')
        projected_sales.clear()
        projected_sales.send_keys(annual_gross_sales)
        square_footage = driver.find_element_by_id('CH_032')
        square_footage.clear()
        square_footage.send_keys(footage)
        driver.save_screenshot(COMPANY_NAME+ '_business_operation_screenshot.png')
        driver.find_element_by_xpath('//*[@id="commercial-app"]/div/div[2]/div[2]/div/div[2]/form/div[1]/div/div/button').click()
        print (driver.current_url)
        time.sleep(10)
        checkPageTransition(curr_url,driver.current_url,'Error in BusinessOp')
    
    if driver.current_url == 'https://'+user.USERNAME+':'+user.PASSWORD+'@psc-chubb-sit.coverhound.us/contact':
        curr_url = driver.current_url
        select_title = Select(driver.find_element_by_id('title'))
        select_title.select_by_visible_text('Dr.')
        f_name = driver.find_element_by_id('first_name')
        f_name.clear()
        f_name.send_keys(FIRST_NAME)
        l_name = driver.find_element_by_id('last_name')
        l_name.clear()
        l_name.send_keys(LAST_NAME)
        ph_num = driver.find_element_by_id('telephone')
        ph_num.clear()
        ph_num.send_keys('8888888888')
        time.sleep(1)
        confirm_email = driver.find_element_by_id('email').get_attribute('value').encode('utf-8')
        address = driver.find_element_by_xpath('//*[@id="field_for_CH_020"]/div[1]/div[1]/input')
        address.clear()
        address.send_keys(address_line)
        time.sleep(1)
        city_text = driver.find_element_by_id('CH_037')
        city_text.clear()
        city_text.send_keys(city)
        time.sleep(1)
        select_insured_state = Select(driver.find_element_by_id('CH_018'))
        select_insured_state.select_by_visible_text(state)
        zip_text = driver.find_element_by_id('CH_038')
        zip_text.clear()
        zip_text.send_keys(zip_code)
        select_other_address = Select(driver.find_element_by_id('CH_027'))
        select_other_address.select_by_visible_text('No Additional Locations')
        driver.save_screenshot(COMPANY_NAME+ '_contact_info_screenshot.png')
        driver.find_element_by_xpath('//*[@id="commercial-app"]/div/div[2]/div[2]/div/form/div/div[1]/div/div/button').click()
        print (driver.current_url)
        time.sleep(10)
        checkPageTransition(curr_url,driver.current_url,'Error in ContactInfo')
        
    if driver.current_url == 'https://'+user.USERNAME+':'+user.PASSWORD+'@psc-chubb-sit.coverhound.us/coverage-detail/bop':   
        curr_url = driver.current_url
        if first_pass:
            driver.find_element_by_xpath('//*[@id="field_for_CH_327"]/div[1]/div[2]/label').click()
            driver.find_element_by_xpath('//*[@id="field_for_CH_300"]/div[1]/div[2]/label').click()
            driver.find_element_by_xpath('//*[@id="field_for_CH_301"]/div[1]/div[2]/label').click()
            driver.find_element_by_xpath('//*[@id="field_for_CH_302"]/div[1]/div[2]/label').click()
            driver.find_element_by_xpath('//*[@id="field_for_CH_303"]/div[1]/div[1]/label').click()
            driver.find_element_by_xpath('//*[@id="field_for_CH_304"]/div[1]/div[2]/label').click()
            if ADDITIONAL_QUESTIONS.count(business_type) > 0:
                driver.find_element_by_xpath('//*[@id="field_for_CH_322"]/div[1]/div[1]/label').click()
                driver.find_element_by_xpath('//*[@id="field_for_CH_323__1122"]/label').click()
            driver.save_screenshot(COMPANY_NAME+'_coverage_detail_screenshot.png')
            driver.find_element_by_xpath('//*[@id="commercial-app"]/div/div[2]/div[2]/div/form/div/div[1]/div/div/button').click()
        print (driver.current_url)
        time.sleep(60)
        checkPageTransition(curr_url,driver.current_url,'Error in Coverage')
        
    driver.save_screenshot('success_screenshot.png')
except Exception as e:
    print (e)
    print (driver.current_url)
    driver.save_screenshot('error_screenshot.png')
    traceback.print_exc()
finally:
    driver.quit()
    print ("Hi")
