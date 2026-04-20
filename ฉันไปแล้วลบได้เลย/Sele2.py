from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv, find_dotenv
from selenium.webdriver.common.action_chains import ActionChains
import os;
import shutil;
import glob;
import time;
import datetime;
from datetime import date, timedelta;


env_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'PP.env')
load_dotenv(env_path)
password = os.getenv('pass')
username = os.getenv('user')

download_dir = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\logfile_formorning\From_load'

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--force-device-scale-factor=0.35")  # ซูมออก 50%
chrome_options.add_argument("--high-dpi-support=0.35")  # รองรับ DPI ปรับสเกล
chrome_options.add_experimental_option("prefs", {
    'download.default_directory': download_dir,
    'profile.default_content_setting_values.automatic_downloads': 1
})

web = 'https://www.ibond.thaibma.or.th/zrr-index'
web2 = 'https://www.ibond.thaibma.or.th/shortterm-index'
web3 = 'https://www.ibond.thaibma.or.th/bond-index'
web4 = 'https://www.ibond.thaibma.or.th/mtm-corp-index'
web5 = 'https://www.ibond.thaibma.or.th/composite-index'
web6 = 'https://www.ibond.thaibma.or.th/corp-zrr-index'


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
actions = ActionChains(driver)
driver.get(web)
driver.execute_script("document.body.style.zoom = '0.3'")
driver.execute_script("document.body.style['-webkit-transform'] = 'scale(0.5)';")

time.sleep(2)
for i in range(360):
    try:
        driver.find_element(By.XPATH,"/html/body/div/div/div/main/form/div[1]/input").send_keys(f"{username}")
        driver.find_element(By.XPATH,"/html/body/div/div/div/main/form/div[2]/input").send_keys(f"{password}")
        driver.find_element(By.XPATH,"/html/body/div/div/div/main/form/button").click()
        break
    except:
        time.sleep(0.5)
print('loginสำเร็จ')

WebDriverWait(driver, 10).until(
    EC.invisibility_of_element_located((By.XPATH, "//button[contains(text(), 'Download all TTM and Date')]"))
)

for i in range(5):
    try:
        accept_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Download all TTM and Date')]"))
        )
        accept_button.click()
        print('กดปุ่ม Accept สำเร็จ')
    except:
        print('ไม่มีปุ่ม Accept หรือกดไม่สำเร็จ')


#WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/span[2]/div/button'))).click()

#driver.find_element(By.XPATH,'//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/span[2]/div/button').click()

time.sleep(30)
#---------------------------------------
driver.get(web2)

time.sleep(3)


time.sleep(2)

icon = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 
        "table > tbody > tr > td:nth-child(13) > div > button > i"))
)
icon.click()
time.sleep(7)


#------------------------------------------
driver.get(web3)

time.sleep(2)

button_1 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 
        "table > tbody > tr:nth-child(1) > td.text-center > div > button > i"))
)
button_1.click()
time.sleep(7)


button_2 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 
        "table > tbody > tr:nth-child(2) > td.text-center > div > button"))
)
button_2.click()
time.sleep(7)

button_3 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 
        "table > tbody > tr:nth-child(3) > td.text-center > div > button > i"))
)
button_3.click()
time.sleep(7)

#-----------------------------------
driver.get(web4)

time.sleep(2)

button_bbb = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 
        "table > tbody > tr:nth-child(7) > td:nth-child(14) > div > button > i"))
)
button_bbb.click()
time.sleep(7)


button_bbb2 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 
        "table > tbody > tr:nth-child(8) > td:nth-child(14) > div > button > i"))
)
button_bbb2.click()
time.sleep(7)

button_bbb3 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 
        "table > tbody > tr:nth-child(9) > td:nth-child(14) > div > button > i"))
)
button_bbb3.click()
time.sleep(7)
#----------------------------------------

driver.get(web5)

time.sleep(2)

button_com = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 
        "table > tbody > tr > td:nth-child(11) > div > button > i"))
)
button_com.click()
time.sleep(7)

#------------------------------------------
driver.get(web6)

time.sleep(2)

button_com = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 
        "div.container > div.main-content > div.essential-div > div > div:nth-child(7) > div > div > div > table > tbody > tr:nth-child(1) > td:nth-child(8) > div > button > i"))
)
button_com.click()
time.sleep(7)


#driver.quit()



#----------------------------------------------------------


#-----------------------------------------------------------


def copy_latest_file(dl_path, des_path, _from, _to):
    latest_file = None
    for i in range(360):
        try:
            list_of_files = glob.glob(f'{dl_path}\\{_from}*.xlsx')
            if not list_of_files:
                raise FileNotFoundError(f"No files found with prefix '{_from}' in directory '{dl_path}'")
            latest_file = max(list_of_files, key=os.path.getctime)
            break
        except Exception as e:
            print(f"Attempt {i+1}: Error occurred - {e}")
            time.sleep(0.5)
    
    if latest_file:
        destination_path = f'{des_path}\\{_to}.xlsx'
        try:
            shutil.copy(latest_file, destination_path)
            print(f"Copied {latest_file} to {destination_path}")
        except Exception as e:
            print(f"Failed to copy {latest_file} to {destination_path}: {e}")
    else:
        print(f"Failed to find or copy the latest file with prefix '{_from}' after 180 seconds")


dl_path = r"\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\logfile_formorning\From_load"
des_path = r"\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\logfile_formorning"


file_mappings = [
    ('ZRRIndexAll', 'ZRR_All'),
    ('STGov_Index_', 'ST_Gov'),
    ('GOV_Index_20', 'Gov'),  
    ('GOV_Index_G1_','Gov_G1'),
    ('GOV_Index_G2_','Gov_G2'),
    ('MTMCorp_BBBplusup_Index_','MTMCorp_BBBplus'),
    ('MTMCorp_BBBplusup_G1_Index_','MTMCorp_BBBplus_G1'),
    ('MTMCorp_BBBplusup_G2_Index_','MTMCorp_BBBplus_G2'),
    ('Composite_Index_','Composite'),
    ('CorpZRR_A_1Y_Index_','CorpZRR_A_1Y')


    
    
]

for _from, _to in file_mappings:
    copy_latest_file(dl_path, des_path, _from, _to)
    
    
    
    
import win32com.client

def run_excel_macro(file_path, macro_name):
    try:

        excel_app = win32com.client.Dispatch("Excel.Application")
        

        excel_app.Visible = True


        workbook = excel_app.Workbooks.Open(file_path)


        excel_app.Run(macro_name)


        workbook.Close(SaveChanges=False)
        excel_app.Quit()
        
        print("Macro executed successfully.")

    except Exception as e:
        print("An error occurred:", e)


file_path = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\power_automate.xlsm'
macro_name2 = 'UpdateData2'
run_excel_macro(file_path, macro_name2)
