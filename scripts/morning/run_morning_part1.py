from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
from risk_logger import get_logger

import os
import shutil
import glob
import time
import datetime
from datetime import date, timedelta

log = get_logger("Run_morning_ThaiBMA")
log.info("=== Morning ThaiBMA Part 1 started ===")

# A2: allow override via RISK_ENV_PATH env var; fall back to Desktop path for backwards compat
env_path = os.environ.get('RISK_ENV_PATH') or os.path.join(os.path.expanduser('~'), 'Desktop', 'PP.env')
load_dotenv(env_path)
log.info(f"Loaded credentials from: {env_path}")
password = os.getenv('pass')
username = os.getenv('user')

#1 Check 

dl_path = r'\\172.16.21.100\Risk$\Morningstar Benchmark';
_from = 'Morningstar Benchmark';
des_path = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\98.Data File';
_to = 'Morningstar Benchmark';

download_dir = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\logfile_formorning\From_load'

def Morningstar_Benchmark(_from, _to):
    for i in range(360):
        try:
            list_of_files = glob.glob(f'{dl_path}\\{_from}*.xls')
            latest_file = max(list_of_files, key=os.path.getctime)
            break
        except:
            time.sleep(0.5)
    destination_path = f'{des_path}\\{_to}.xls'
    shutil.copy(latest_file, destination_path)

try:
    Morningstar_Benchmark(_from, _to)
    log.info("Morningstar Benchmark copied successfully")
except Exception as e:
    log.error(f"Morningstar Benchmark copy failed: {e}", exc_info=True)


def wait_for_download(download_path, timeout=10):
    start_time = time.time()  
    
    while any([filename.endswith(".crdownload") for filename in os.listdir(download_path)]):
        # ตรวจสอบว่าเวลาที่ผ่านไป เกิน timeout ที่ตั้งไว้หรือยัง
        if time.time() - start_time > timeout:
            print("หมดเวลารอ (10 วินาที) จะเริ่มทำงานขั้นตอนถัดไป...")
            return  # ออกจากฟังก์ชันทันที
            
        time.sleep(1)
        
    print("Download finished!")

#3 + 7


chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("prefs", {
    'download.default_directory': download_dir,
    'profile.default_content_setting_values.automatic_downloads': 1
})


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

web_login = 'https://www.ibond.thaibma.or.th/login?page=/bondsearch/bondsearchpage'
web = 'https://www.ibond.thaibma.or.th/yield-curve/government'
web2 = 'https://www.ibond.thaibma.or.th/mtm-gov-index'
web3 = 'https://www.ibond.thaibma.or.th/cp-index'
web4 = 'https://www.ibond.thaibma.or.th/mtm-corp-index'
web5 = 'https://www.ibond.thaibma.or.th/esg-index'

dl_path = 'C:\\Users\\Amornsiris\\Downloads'
_from = 'YieldTTM_202'
des_path = r"C:\Users\Amornsiris\Desktop\Daily_for_Daily"
_to = 'D_YieldTTM'



driver.get(web_login)

time.sleep(4)

for i in range(360):
    try:
        driver.find_element(By.XPATH,"/html/body/div/div/div/main/form/div[1]/input").send_keys(f"{username}")
        driver.find_element(By.XPATH,"/html/body/div/div/div/main/form/div[2]/input").send_keys(f"{password}")
        driver.find_element(By.XPATH,"/html/body/div/div/div/main/form/button").click()
        time.sleep(7.5)
        break
    except:
        time.sleep(0.5)
log.info("ThaiBMA login successful")



driver.get(web)
driver.execute_script("document.body.style.transform = 'scale(0.6)'")
time.sleep(4)



for i in range(360):
    try:
        driver.find_element(By.XPATH,"//*[@id='root']/div[3]/div[1]/div[3]/div/div[2]/div/div/span[3]/div/button").click()
        break
    except:
        time.sleep(6.5)

time.sleep(10)        
        
for i in range(360):
    try:
        driver.find_element(By.XPATH,"//*[@id='root']/div[3]/div[1]/div[3]/div/div[2]/div/div/span[3]/div/button").click()
        break
    except:
        time.sleep(3.5)        
                                     
print('ไฟล์Paperสำเร็จ')


time.sleep(5)

driver.get(web2)

time.sleep(7) 

try:
    accept_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "rcc-confirm-button"))
    )
    accept_button.click()
    
    driver.find_element(By.XPATH,"//*[@id='root']/div[3]/div[1]/div[3]/div/div[2]/div/div/span[3]/div/button").click()
    
    print('กดปุ่ม Accept สำเร็จ')
except:
    print('ไม่มีปุ่ม Accept หรือกดไม่สำเร็จ')


from selenium.common.exceptions import TimeoutException, NoSuchElementException

def LoadFile(xpath, file_number):
    try:
        # รอไม่เกิน 3 วินาที
        element = WebDriverWait(driver, 7).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", element)
        time.sleep(3)
        print(f'ไฟล์ {file_number} สำเร็จ')
    except (TimeoutException, NoSuchElementException):
        print(f'ไฟล์ {file_number} ข้าม (ไม่เจอ element)')
    except Exception as e:
        print(f'ไฟล์ {file_number} ไม่สำเร็จ: {e}')
        
    try:
        driver.find_element(By.XPATH,xpath).click()
        time.sleep(4)
        print(f'ไฟล์ {file_number} สำเร็จ')
    except (TimeoutException, NoSuchElementException):
        print(f'ไฟล์ {file_number} ข้าม (ไม่เจอ element)')
    except Exception as e:
        print(f'ไฟล์ {file_number} ไม่สำเร็จ: {e}')

LoadFile("//table//tr[1]/td[14]//button//i", "แรก")

wait_for_download(download_dir)
time.sleep(1)

LoadFile("//table//tr[2]/td[14]//button//i", "สอง")
wait_for_download(download_dir)
time.sleep(1)


LoadFile("//table//tr[3]/td[14]//button//i", "สาม")
wait_for_download(download_dir)
time.sleep(5)


driver.get(web3)

def LoadA():
    for i in range(360):
        try:
            driver.find_element(By.XPATH,"//table///tr[2]/td[13]//button//i").click()
            break
        except:
            time.sleep(0.5)
    print('ไฟล์Aสำเร็จ')
    
LoadFile("//table//tr[2]/td[13]//button//i", "สี่")  
time.sleep(7)


driver.get(web4)

def LoadCorBond():
    for i in range(360):
        try:
            driver.find_element(By.XPATH,'//table//tr[2]/td[14]//button//i').click()
            break
        except:
            time.sleep(0.5)
    print('ไฟล์CorBondสำเร็จ')

LoadFile('//table//tr[2]/td[14]//button//i',"ห้า")

time.sleep(10)

driver.get(web5)

def LoadESGGovBond():
    for i in range(360):
        try:
            driver.find_element(By.XPATH,'//table//tr[6]/td[14]//button//i').click()
            break
        except:
            time.sleep(0.5)
    print('ไฟล์ESGGovBondสำเร็จ')

LoadFile('//table//tr[6]/td[14]//button//i',"หก")

time.sleep(10)

driver.close()

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
    ('MTMGov_G1_Index_', 'MTMGov_G1'),
    ('MTMGov_G2_Index_', 'MTMGov_G2'),
    ('MTMGov_Index_202', 'Gov_Index'),  
    ('YieldTTM_202','D_YieldTTM'),
    ('CP_Aminusup_Index','Commercial Paper Index'),
    ('MTMCorp_Aminusup_G1_Index','CorBond_G1'),
    ('ESGGOV_Index','ESGGovBond')
    
    
]

for _from, _to in file_mappings:
    log.info(f"Copying file: {_from} → {_to}")
    copy_latest_file(dl_path, des_path, _from, _to)

log.info("All file copies complete")
time.sleep(2)

dl_path = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\logfile_formorning\From_load'
_from = 'YieldTTM_202'
des_path2 = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\98.Data File'
_to = 'D_YieldTTM'

def D_YieldTTM(_from, _to):
    for i in range(360):
        try:
            list_of_files = glob.glob(f'{dl_path}\\{_from}*.xlsx')
            latest_file = max(list_of_files, key=os.path.getctime)
            break
        except:
            time.sleep(0.5)
    destination_path = f'{des_path2}\\{_to}.xlsx'
    shutil.copy(latest_file, destination_path)


D_YieldTTM(_from, _to)




time.sleep(2)

import win32com.client

def run_excel_macro(file_path, macro_name):
    try:

        excel_app = win32com.client.Dispatch("Excel.Application")
        

        excel_app.Visible = False


        workbook = excel_app.Workbooks.Open(file_path)


        excel_app.Run(macro_name)


        workbook.Close(SaveChanges=False)
        excel_app.Quit()
        
        print("Macro executed successfully.")

    except Exception as e:
        print("An error occurred:", e)


file_path = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\power_automate.xlsm'
macro_name = 'UpdateData'
run_excel_macro(file_path, macro_name)


def get_previous_business_day(reference_date):
    previous_day = reference_date - timedelta(days=1)
    while previous_day.weekday() > 4:  
        previous_day -= timedelta(days=1)
    return previous_day

yesterday_business_day = get_previous_business_day(date.today())
folder = yesterday_business_day.strftime("%Y-%m-%d")

time.sleep(2)

# Exclude temporary files starting with ~$ in the search pattern
list_of_files = [file for file in glob.glob(r'P:\Fund_EQ\**\*PortVal*') if not os.path.basename(file).startswith('~$')]
latest_file = max(list_of_files, key=os.path.getctime)

print(latest_file)

destination = r"\\w2fsspho101.lhfund.net\FM-RI$\risk\98.Data File\D_LHReport.xlsx"

shutil.copy2(latest_file, destination)
log.info(f"Copied LHReport: {latest_file} → {destination}")
log.info("=== Morning ThaiBMA Part 1 completed ===")