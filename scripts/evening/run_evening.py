from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from risk_logger import get_logger, send_failure_alert
import sys
import time
import os
import glob
import shutil
import pandas as pd
import win32com.client
import subprocess
from datetime import datetime, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from config import (
    EVENING_DL_DIR, EVENING_DATA_DIR,
    POWER_AUTOMATE_PM,
    THAIBMA_LOGIN_ST, THAIBMA_LOGIN_BOND, THAIBMA_LOGIN_CORP,
    GPO_FIXED_FILE, GPO_EQ_FILE, SET_TRI_GLOB,
    MACRO_NEW_FINISH_SET, MACRO_CREATE_PM, MACRO_EVENING_BENCH,
    EVENING_FILE_MAPPINGS, EMAIL_RECIPIENTS,
)

log = get_logger("GPO")

_proj_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
env_path = os.environ.get('RISK_ENV_PATH') or os.path.join(_proj_root, 'scripts', '.env')
load_dotenv(env_path)
password = os.getenv('pass')
username = os.getenv('user')

download_dir = EVENING_DL_DIR

options = webdriver.EdgeOptions()
prefs = {
    'download.default_directory': download_dir,
    'plugins.always_open_pdf_externally': True,
    'profile.default_content_setting_values.automatic_downloads': 1
}
options.add_argument("--start-maximized")
options.add_experimental_option('prefs', prefs)
service = Service(EdgeChromiumDriverManager().install())

xpath1 = '//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/table/tbody/tr/td[13]/div/button/i'
xpathdate1 = '//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/div[1]/div[1]'
xpath2 = '//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/table/tbody/tr[2]/td[13]/div/button/i'
xpathdate2 = '//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/div[1]/div[1]'
xpath3 = '//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/table/tbody/tr[8]/td[14]/div/button/i'
xpathdate3 = '//*[@id="root"]/div[3]/div[1]/div[3]/div/div[2]/div/div/div/div[1]/div[1]'

macro_name_Bench = MACRO_EVENING_BENCH
fileGPO = GPO_FIXED_FILE
fileGPO2 = GPO_EQ_FILE


def Create_Afternoon(file_path, macro_name):
    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = True
        workbook = excel_app.Workbooks.Open(file_path)
        excel_app.Run(macro_name)
        workbook.Close(SaveChanges=False)
        excel_app.Quit()
        while not excel_app.Ready:
            time.sleep(1)
        print("Macro executed successfully.")
    except Exception as e:
        print("An error occurred:", e)


def open_file_and_run_macro(file_path_Bench: str, name2: str):
    excel_app = win32com.client.Dispatch('Excel.Application')
    excel_app.Visible = True
    excel_app.Workbooks.Open(os.path.abspath(file_path_Bench))
    excel_app.Application.Run(f'{name2}!Module1.{macro_name_Bench}')
    while not excel_app.Application.Ready:
        time.sleep(1)
    excel_app.Quit()
    del excel_app


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
                    driver.find_element(By.XPATH, xpath).click()
                    time.sleep(6.5)
                    print(f'ไฟล์ {file_number} สำเร็จ')
                    break
                except Exception as e:
                    print(f'ไฟล์ {file_number} ไม่สำเร็จ: {e}')
            else:
                time.sleep(1.5)
                driver.refresh()
                time.sleep(1.5)
        except Exception as e:
            print(f'เกิดข้อผิดพลาดในการตรวจสอบวันที่: {e}')


def doallTBMA(yesterday):
    driver = webdriver.Edge(options=options, service=service)
    try:
        driver.get(THAIBMA_LOGIN_ST)
        time.sleep(5)
        driver.find_element(By.XPATH, "/html/body/div/div/div/main/form/div[1]/input").send_keys(username)
        driver.find_element(By.XPATH, "/html/body/div/div/div/main/form/div[2]/input").send_keys(password)
        driver.find_element(By.XPATH, "/html/body/div/div/div/main/form/button").click()
        time.sleep(5)

        LoadFile(driver, THAIBMA_LOGIN_ST, xpath1, xpathdate1, 'ST_GBI', yesterday)
        time.sleep(1)
        LoadFile(driver, THAIBMA_LOGIN_BOND, xpath2, xpathdate2, 'GBI_G1', yesterday)
        time.sleep(1)
        LoadFile(driver, THAIBMA_LOGIN_CORP, xpath3, xpathdate3, 'BBBplus', yesterday)
        time.sleep(3)
    finally:
        driver.quit()
        subprocess.call("TASKKILL /f /IM msedgedriver.exe")
        time.sleep(4)


def Transfer(_from, _to):
    for i in range(360):
        try:
            list_of_files = glob.glob(f'{EVENING_DL_DIR}\\{_from}*.xlsx')
            print(f'Attempt {i+1}: Found files: {list_of_files}')
            if not list_of_files:
                raise FileNotFoundError(f"No files matched pattern: {EVENING_DL_DIR}\\{_from}*.xlsx")
            latest_file = max(list_of_files, key=os.path.getctime)
            print(f'Latest file: {latest_file}')
            break
        except Exception as e:
            print(f'Error: {e}')
            time.sleep(0.5)
    else:
        raise FileNotFoundError("No files found after multiple attempts.")

    destination_path = f'{EVENING_DATA_DIR}\\{_to}.xlsx'
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
    except Exception as e:
        print("An error occurred:", e)


def partonetransfer():
    time.sleep(3)
    for _from, _to in EVENING_FILE_MAPPINGS:
        Transfer(_from, _to)
    time.sleep(5)
    run_excel_macro(POWER_AUTOMATE_PM, MACRO_CREATE_PM)


def new_set():
    today = datetime.now().strftime('%Y%m%d')
    global latest_set_file
    for i in range(360):
        latest_set_file = max(glob.glob(SET_TRI_GLOB), key=os.path.getmtime)
        print(latest_set_file)
        latest_filename = os.path.basename(latest_set_file)
        datefile = latest_filename[7:-4]
        print(datefile)
        if datefile == today:
            break
        else:
            print("ไม่มีไฟล์ในโฟลเดอร์นี้")
            time.sleep(2)


def main():
    log.info("=== GPO evening workflow started ===")
    log.info(f"Loaded credentials from: {env_path}")

    yesterday = (datetime.now() - timedelta(days=0)).strftime('%d/%m/%Y')
    today = (datetime.now() - timedelta(days=0)).strftime('%d/%m/%Y')
    today_dash = datetime.now().strftime('%Y-%m-%d')

    log.info("Running macro Create_Afternoon on power_automate_for_Afternoon.xlsm")
    Create_Afternoon(POWER_AUTOMATE_PM, MACRO_CREATE_PM)
    log.info("Create_Afternoon macro complete")

    doallTBMA(yesterday)

    partonetransfer()

    def read_df():
        while True:
            df = pd.read_excel(fileGPO, sheet_name='Benchmark - PI')
            if df.iloc[-1, 0].strftime('%d/%m/%Y') == today and pd.notna(df.iloc[-1, [2, 4, 5]]).all().all():
                print('complete part one')
                break
            else:
                doallTBMA(yesterday)
                partonetransfer()

    read_df()
    subprocess.call("TASKKILL /f /IM msedgedriver.exe")

    global latest_set_file
    latest_set_file = None
    new_set()

    def Check_set():
        df_set = pd.read_csv(latest_set_file)
        if df_set.iloc[0, 0][:10] == today_dash:
            print('ผ่าน')
            print(df_set.iloc[0, 0][:10])
        else:
            new_set()

    Check_set()

    shutil.copy(latest_set_file, f'{EVENING_DATA_DIR}\\set.csv')
    run_excel_macro(POWER_AUTOMATE_PM, MACRO_NEW_FINISH_SET)

    df1 = pd.read_excel(fileGPO, sheet_name='Benchmark - PI', header=4)
    df2 = pd.read_excel(fileGPO2, sheet_name='Benchmark - PI', header=4)

    table_str1 = df1.iloc[:, [0, 1, 2, 4, 5]].tail(1).to_html(header=True, index=False)
    table_str2 = df2.iloc[:, [0, 2]].tail(1).to_html(header=True, index=False)

    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'อัพเดท GPO'
    newmail.To = EMAIL_RECIPIENTS
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
    newmail.Send()
    log.info("GPO update email sent successfully")
    log.info("=== GPO evening workflow completed ===")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log.critical(f"Script failed: {e}", exc_info=True)
        send_failure_alert("GPO", str(e))
        sys.exit(1)
