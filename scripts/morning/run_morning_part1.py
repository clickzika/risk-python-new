from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from dotenv import load_dotenv
from risk_logger import get_logger, send_failure_alert
import win32com.client
import sys
import os
import shutil
import glob
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from config import (
    MORNINGSTAR_SRC, DATA_FILE_DIR, MORNING_DL_DIR, MORNING_DATA_DIR,
    LH_REPORT_GLOB, LH_REPORT_DEST, POWER_AUTOMATE,
    THAIBMA_LOGIN, THAIBMA_YIELD_GOV, THAIBMA_MTM_GOV,
    THAIBMA_CP, THAIBMA_MTM_CORP, THAIBMA_ESG,
    MACRO_UPDATE_DATA, MORNING_PART1_FILE_MAPPINGS,
)

log = get_logger("Run_morning_ThaiBMA")

_proj_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
env_path = os.environ.get('RISK_ENV_PATH') or os.path.join(_proj_root, 'scripts', '.env')
load_dotenv(env_path)
password = os.getenv('pass')
username = os.getenv('user')

download_dir = MORNING_DL_DIR


def Morningstar_Benchmark():
    for i in range(360):
        try:
            list_of_files = glob.glob(f'{MORNINGSTAR_SRC}\\Morningstar Benchmark*.xls')
            latest_file = max(list_of_files, key=os.path.getctime)
            break
        except Exception:
            time.sleep(0.5)
    shutil.copy(latest_file, f'{DATA_FILE_DIR}\\Morningstar Benchmark.xls')


def wait_for_download(download_path, timeout=10):
    start_time = time.time()
    while any([filename.endswith(".tmp") for filename in os.listdir(download_path)]):
        if time.time() - start_time > timeout:
            print("หมดเวลารอ (10 วินาที) จะเริ่มทำงานขั้นตอนถัดไป...")
            return
        time.sleep(1)
    print("Download finished!")


def LoadFile(driver, xpath, file_number):
    try:
        element = WebDriverWait(driver, 7).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", element)
        time.sleep(3)
        log.info(f"Downloaded file {file_number}")
    except (TimeoutException, NoSuchElementException):
        log.warning(f"File {file_number} skipped — element not found")
    except Exception as e:
        log.error(f"File {file_number} failed: {e}")


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


def D_YieldTTM():
    for i in range(360):
        try:
            list_of_files = glob.glob(f'{MORNING_DL_DIR}\\YieldTTM_202*.xlsx')
            latest_file = max(list_of_files, key=os.path.getctime)
            break
        except Exception:
            time.sleep(0.5)
    shutil.copy(latest_file, f'{DATA_FILE_DIR}\\D_YieldTTM.xlsx')


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


def main():
    log.info("=== Morning ThaiBMA Part 1 started ===")
    log.info(f"Loaded credentials from: {env_path}")

    try:
        Morningstar_Benchmark()
        log.info("Morningstar Benchmark copied successfully")
    except Exception as e:
        log.error(f"Morningstar Benchmark copy failed: {e}", exc_info=True)

    edge_options = Options()
    edge_options.add_argument("--start-maximized")
    edge_options.add_experimental_option("prefs", {
        'download.default_directory': download_dir,
        'profile.default_content_setting_values.automatic_downloads': 1
    })

    driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=edge_options)
    try:
        driver.get(THAIBMA_LOGIN)
        time.sleep(4)

        for i in range(360):
            try:
                driver.find_element(By.XPATH, "/html/body/div/div/div/main/form/div[1]/input").send_keys(f"{username}")
                driver.find_element(By.XPATH, "/html/body/div/div/div/main/form/div[2]/input").send_keys(f"{password}")
                driver.find_element(By.XPATH, "/html/body/div/div/div/main/form/button").click()
                time.sleep(7.5)
                break
            except Exception:
                time.sleep(0.5)
        log.info("ThaiBMA login successful")

        driver.get(THAIBMA_YIELD_GOV)
        driver.execute_script("document.body.style.transform = 'scale(0.6)'")
        time.sleep(4)

        for i in range(360):
            try:
                driver.find_element(By.XPATH, "//*[@id='root']/div[3]/div[1]/div[3]/div/div[2]/div/div/span[3]/div/button").click()
                break
            except Exception:
                time.sleep(6.5)

        time.sleep(10)

        for i in range(360):
            try:
                driver.find_element(By.XPATH, "//*[@id='root']/div[3]/div[1]/div[3]/div/div[2]/div/div/span[3]/div/button").click()
                break
            except Exception:
                time.sleep(3.5)

        print('ไฟล์Paperสำเร็จ')
        time.sleep(5)

        driver.get(THAIBMA_MTM_GOV)
        time.sleep(7)

        try:
            accept_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "rcc-confirm-button"))
            )
            accept_button.click()
            driver.find_element(By.XPATH, "//*[@id='root']/div[3]/div[1]/div[3]/div/div[2]/div/div/span[3]/div/button").click()
            print('กดปุ่ม Accept สำเร็จ')
        except Exception:
            print('ไม่มีปุ่ม Accept หรือกดไม่สำเร็จ')

        LoadFile(driver, "//table//tr[1]/td[14]//button//i", "แรก")
        wait_for_download(download_dir)
        time.sleep(1)

        LoadFile(driver, "//table//tr[2]/td[14]//button//i", "สอง")
        wait_for_download(download_dir)
        time.sleep(1)

        LoadFile(driver, "//table//tr[3]/td[14]//button//i", "สาม")
        wait_for_download(download_dir)
        time.sleep(5)

        driver.get(THAIBMA_CP)
        LoadFile(driver, "//table//tr[2]/td[13]//button//i", "สี่")
        time.sleep(7)

        driver.get(THAIBMA_MTM_CORP)
        LoadFile(driver, '//table//tr[2]/td[14]//button//i', "ห้า")
        time.sleep(10)

        driver.get(THAIBMA_ESG)
        LoadFile(driver, '//table//tr[6]/td[14]//button//i', "หก")
        time.sleep(10)
    finally:
        driver.quit()

    for _from, _to in MORNING_PART1_FILE_MAPPINGS:
        log.info(f"Copying file: {_from} → {_to}")
        copy_latest_file(MORNING_DL_DIR, MORNING_DATA_DIR, _from, _to)
    log.info("All file copies complete")

    time.sleep(2)
    D_YieldTTM()

    time.sleep(2)
    run_excel_macro(POWER_AUTOMATE, MACRO_UPDATE_DATA)

    time.sleep(2)

    list_of_files = [f for f in glob.glob(LH_REPORT_GLOB) if not os.path.basename(f).startswith('~$')]
    latest_file = max(list_of_files, key=os.path.getctime)
    shutil.copy2(latest_file, LH_REPORT_DEST)
    log.info(f"Copied LHReport: {latest_file} → {LH_REPORT_DEST}")
    log.info("=== Morning ThaiBMA Part 1 completed ===")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log.critical(f"Script failed: {e}", exc_info=True)
        send_failure_alert("Run_morning_ThaiBMA", str(e))
        sys.exit(1)
