from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.common.exceptions import NoSuchElementException
from risk_logger import get_logger
import time
import os
import glob
import shutil
import zipfile
import pandas as pd
import win32com.client
import subprocess
from pathlib import Path
from datetime import datetime, timedelta

log = get_logger("GPO")
log.info("=== GPO evening workflow started ===")


file_path = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\power_automate_for_Afternoon.xlsm'
#######*************************
macro_name = 'Create_Afternoon'
file_path_Bench = r"\\w2fsspho101.lhfund.net\FM-RI$\risk\1.Risk Report\1.Daily Risk Report\Benchmark.xlsm"
name2 =  r"Benchmark.xlsm"
macro_name_Bench = r'evening'


def Create_Afternoon(file_path, macro_name):
    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = True
        workbook = excel_app.Workbooks.Open(file_path)
        excel_app.Run(macro_name)
        workbook.Close(SaveChanges=False)
        excel_app.Quit()
        while not excel_app.Ready:
            time.sleep(1)  # รอ 1 วินาทีแล้วเช็คใหม่        
        print("Macro executed successfully.")

    except Exception as e:
        print("An error occurred:", e)

log.info("Running macro Create_Afternoon on power_automate_for_Afternoon.xlsm")
Create_Afternoon(file_path, macro_name)
log.info("Create_Afternoon macro complete")

def open_file_and_run_macro(file_path_Bench: str, name2: str):
    """
    เปิดไฟล์ Excel และรันมาโครในไฟล์
    :param file_path_Bench: เส้นทางของไฟล์ (รวมชื่อไฟล์)
    :param name2: ชื่อไฟล์ พร้อม .xlsm extension
    """
    excel_app = win32com.client.Dispatch('Excel.Application')
    excel_app.Visible = True
    wb = excel_app.Workbooks.Open(os.path.abspath(file_path_Bench))
    excel_app.Application.Run(f'{name2}!Module1.{macro_name_Bench}')

    while not excel_app.Application.Ready:
        time.sleep(1)  # รอ 1 วินาทีแล้วเช็คใหม่

    # wb.Close(SaveChanges=True)  # ปิดไฟล์และบันทึกการเปลี่ยนแปลงถ้าจำเป็น
    excel_app.Quit()
    del excel_app

# เรียกใช้งานฟังก์ชัน
#open_file_and_run_macro(file_path_Bench, name2)

# def open_file_and_run_macro(file_path_Bench: str, name2: str):
#         """
#         :param file_path: path to file, including file name
#         :param name: file name, with .xlsm extension
#         """
#         xl = win32com.client.Dispatch('Excel.Application')
#         xl.Visible = True
#         wb = xl.Workbooks.Open(os.path.abspath(file_path_Bench))
#         xl.Application.Run(f'{name2}!Module1.{macro_name_Bench}')
#         while not xl.Application.Ready:
#             time.sleep(1)  # รอ 1 วินาทีแล้วเช็คใหม่
#         #wb.Close(SaveChanges=True)
#         xl.Quit()
#         del xl
# open_file_and_run_macro(file_path_Bench, name2)


#########**********************

yesterday= (datetime.now() - timedelta(days=0)).strftime('%d/%m/%Y')
today = (datetime.now() - timedelta(days=0)).strftime('%d/%m/%Y')
#########*************************

xpath1 = '//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/table/tbody/tr/td[13]/div/button/i'
xpathdate1 = '//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/div[1]/div[1]'

xpath2 ='//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/table/tbody/tr[2]/td[13]/div/button/i'
xpathdate2 = '//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/div[1]/div[1]'

xpath3 = '//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/table/tbody/tr[8]/td[14]/div/button/i'
xpathdate3 = '//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/div[1]/div[1]'

web = 'https://www.ibond.thaibma.or.th/login?page=/shortterm-index'
web2 = 'https://www.ibond.thaibma.or.th/login?page=/bond-index'
web3 = 'https://www.ibond.thaibma.or.th/login?page=/mtm-corp-index'

dl_path = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\Logfile_forevening\From_load'
_from = 'GOV_Index_G1'
des_path = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\Logfile_forevening'
_to = 'GOV_Index_G1_after'

_from2 = 'MTMCorp_BBBplusup_G1'
_to2 = 'MTMCorp_BBBplusup_G1'

_from3 = 'STGov_Index'
_to3 = 'STGov_Index'


macro_name = 'finish_BMA'

fileGPO = r'P:\Bloomberg\Management Fee for PVD\Management Fee for PVD_GPO-FIXED - LHFUND_REVISED_RATE.xls'
fileGPO2 = r'P:\Bloomberg\Management Fee for PVD\Management Fee for PVD_GPO-EQ - LHFUND  REVISED_RATE.xls'
##########******************************


_proj_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
env_path = os.environ.get('RISK_ENV_PATH') or os.path.join(_proj_root, 'scripts', '.env')
load_dotenv(env_path)
log.info(f"Loaded credentials from: {env_path}")
password = os.getenv('pass')
username = os.getenv('user')

download_dir = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\Logfile_forevening\From_load'
options = webdriver.EdgeOptions()
prefs = {
    'download.default_directory': download_dir,
    'plugins.always_open_pdf_externally': True,
    'profile.default_content_setting_values.automatic_downloads': 1
}
options.add_argument("--start-maximized")
options.add_experimental_option('prefs', prefs)

service = Service(EdgeChromiumDriverManager().install())
driver = webdriver.Edge(options=options, service=service)

    
def LoadFile(driver, web, xpath, xpathdate, file_number, yesterday):
    driver.get(web)
    driver.execute_script("document.body.style.zoom = '0.7'")
    time.sleep(5)

    for i in range(360):
        try:
            date_value = driver.find_element(By.XPATH, xpathdate).text.split("as of ")[1]
            date_object = datetime.strptime(date_value, "%d %B %Y")
            formatted_date = date_object.strftime("%d/%m/%Y")
            print(formatted_date)
            print(yesterday)
            time.sleep(3)

            if formatted_date == yesterday:
                try:

                    driver.find_element(By.XPATH,xpath).click()

                    time.sleep(6.5)
                    print(f'ไฟล์ {file_number} สำเร็จ')
                    break  # ถ้าสำเร็จ ออกจากลูปเลย
                except Exception as e:
                    print(f'ไฟล์ {file_number} ไม่สำเร็จ: {e}')

            else:
                time.sleep(1.5)
                driver.refresh()
                time.sleep(1.5)

        except Exception as e:
            print(f'เกิดข้อผิดพลาดในการตรวจสอบวันที่: {e}')



def doallTBMA():
    driver = webdriver.Edge(options=options, service=service)

    # ล็อกอินแค่ครั้งเดียว
    driver.get(web)
    time.sleep(5)
    driver.find_element(By.XPATH, "/html/body/div/div/div/main/form/div[1]/input").send_keys(username)
    driver.find_element(By.XPATH, "/html/body/div/div/div/main/form/div[2]/input").send_keys(password)
    driver.find_element(By.XPATH, "/html/body/div/div/div/main/form/button").click()
    time.sleep(5)

    # โหลดไฟล์จากแต่ละเว็บ
    LoadFile(driver, web, xpath1, xpathdate1, 'ST_GBI', yesterday)
    time.sleep(1)
    LoadFile(driver, web2, xpath2, xpathdate2, 'GBI_G1', yesterday)
    time.sleep(1)
    LoadFile(driver, web3, xpath3, xpathdate3, 'BBBplus', yesterday)
    time.sleep(3)
    driver.quit()
    subprocess.call("TASKKILL /f  /IM  CHROMEDRIVER.EXE")
    time.sleep(4)


doallTBMA()


#---------------------------------------------------------------------------------


def Transfer(_from, _to):
    for i in range(360):
        try:
            list_of_files = glob.glob(f'{dl_path}\\{_from}*.xlsx')
            print(f'Attempt {i+1}: Found files: {list_of_files}')
            if not list_of_files:
                raise FileNotFoundError(f"No files matched pattern: {dl_path}\\{_from}*.xlsx")
            latest_file = max(list_of_files, key=os.path.getctime)
            print(f'Latest file: {latest_file}')
            break
        except Exception as e:
            print(f'Error: {e}')
            time.sleep(0.5)
    else:
        raise FileNotFoundError("No files found after multiple attempts.")

    destination_path = f'{des_path}\\{_to}.xlsx'
    print(f'Copying file to: {destination_path}')
    shutil.copy(latest_file, destination_path)
    print('File copied successfully.')



def run_excel_macro(file_path, macro_name):
    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = True
        workbook = excel_app.Workbooks.Open(file_path)
        excel_app.Run(macro_name)
        workbook.Close(SaveChanges=False)
        excel_app.Quit()
        
        print("Macro executed successfully.")
        import time
        while not excel.Application.Ready:
            time.sleep(1)  

    except Exception as e:
        print("An error occurred:", e)

#      """
# def run_excel_macro2(file_path2, macro_name2):
#     try:
#         excel_app = win32com.client.Dispatch("Excel.Application")
#         excel_app.Visible = False
#         workbook = excel_app.Workbooks.Open(file_path)
#         excel_app.Run(macro_name)
#         workbook.Close(SaveChanges=True)
#         excel_app.Quit()
        
#         print("Macro executed successfully.")

#     except Exception as e:
#          print("An error occurred:", e)

# file_path2 = r'C:\Users\Amornsiris\Desktop\Daily_for_Daily\afternoon\Benchmark.xlsm'
# macro_name2 = 'Evening'

#     """

def partonetransfer():
    time.sleep(3)
    Transfer(_from, _to)
    Transfer(_from2, _to2)
    Transfer(_from3, _to3)
    time.sleep(5)
    run_excel_macro(file_path, macro_name)
partonetransfer()


time.sleep(2)


def read_df():
    while True:
        df = pd.read_excel(fileGPO, sheet_name='Benchmark - PI')
        if df.iloc[-1, 0].strftime('%d/%m/%Y') == today and pd.notna(df.iloc[-1, [2, 4, 5]]).all().all() == True:
            #run_excel_macro2(file_path2, macro_name2)
            print('complete part one')
            break  # ออกจาก loop เมื่อเงื่อนไขสำเร็จ
        else:
            doallTBMA()
            partonetransfer()
            #run_excel_macro2(file_path2, macro_name2)
read_df()
subprocess.call("TASKKILL /f  /IM  CHROMEDRIVER.EXE")

#-----------------------------------------------------------------------------------------------------------
today_dash = datetime.now().strftime('%Y-%m-%d')
def new_set():
    today = datetime.now().strftime('%Y%m%d')
    today_dash = datetime.now().strftime('%Y-%m-%d')
    yesterday = (datetime.now()- timedelta(1)).strftime('%Y%m%d')
    yesterday_dash = (datetime.now()- timedelta(1)).strftime('%Y-%m-%d')
    global latest_file
    for i in range(360):    
        latest_file = max(glob.glob(r"P:\###RISK###\SET TRI\202*\*\*"), key=os.path.getmtime)
        print(latest_file)
        latest_filename = os.path.basename(latest_file)  # เอาเฉพาะชื่อไฟล์
        print(latest_filename)
        datefile = latest_filename[7:-4]
        print(datefile)
        if datefile == today:

            break
        else:
            print("ไม่มีไฟล์ในโฟลเดอร์นี้")
            time.sleep(2)
new_set()           

def Check_set():
    df_set = pd.read_csv(latest_file)
    if df_set.iloc[0, 0][:10] == today_dash:
        print('ผ่าน')
        print(df_set.iloc[0, 0][:10])

    else:
        new_set()
Check_set()           


des_path = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\Logfile_forevening'
_toset = 'set'
destination_path = f'{des_path}\\{_toset}.csv'


shutil.copy(latest_file, destination_path)



macro_name99 = 'new_finish_set'

run_excel_macro(file_path, macro_name99)



#--------------------------------------------------------------------------------------------------------------


df1 = pd.read_excel(fileGPO,sheet_name = 'Benchmark - PI',header=4 )
df2 = pd.read_excel(fileGPO2,sheet_name = 'Benchmark - PI',header=4)


table_str1 = df1.iloc[:,[0,1,2,4,5]].tail(1).to_html(header=True, index=False)
table_str2 = df2.iloc[:, [0, 2]].tail(1).to_html(header=True, index=False)

#table_str1 = df1.iloc[:, :7].tail(3).to_html(header=True, index=False)
#table_str2 = df2.iloc[:, [0, 2, 4]].tail(3).to_html(header=True, index=False)

ol = win32com.client.Dispatch('Outlook.Application')
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)
newmail.Subject = 'อัพเดท GPO'
newmail.To = 'risk@lhfund.co.th ; operation@lhfund.co.th'
#newmail.To = 'Panisarap@lhfund.co.th ; kornwipad@lhfund.co.th ; Amornsiris@lhfund.co.th'
#newmail.To = 'Amornsiris@lhfund.co.th'
############*********************************
newmail.HTMLBody = f'''
<html>
<head>
<style>
    body {{ font-size: 10px; }}
    table {{ width: 100%; border-collapse: collapse; }}
    table, th, td {{ border: 1px solid black; }}
    th, td {{ padding: 8px; text-align: left; }}
    tr:nth-child(even) {{ background-color: #f2f2f2; }}
    tr:hover {{ background-color: #ddd; }}
    th {{ background-color: #4CAF50; color: white; }}
</style>
</head>
<body>
<p style="font-size:18px; color: navy;">GPO เรียบร้อยครับ</p>
<p></p>
<p style="font-size:13px;">GPO-FIXED </p>
<p style="font-size:10px;">{table_str1}</p>
<p></p>
<p style="font-size:13px;">GPO-EQ</p>
<p style="font-size:10px;">{table_str2}</p>
</body>
</html>
'''

# แนบไฟล์ (ถ้ามี)
# attach = 'C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
# newmail.Attachments.Add(attach)

# ส่งอีเมล
newmail.Send()
log.info("GPO update email sent successfully")
log.info("=== GPO evening workflow completed ===")
