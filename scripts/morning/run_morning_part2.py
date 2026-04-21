from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
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
    MORNING_DL_DIR, MORNING_DATA_DIR, POWER_AUTOMATE,
    THAIBMA_LOGIN, THAIBMA_ZRR, THAIBMA_SHORTTERM, THAIBMA_BOND,
    THAIBMA_MTM_CORP, THAIBMA_COMPOSITE, THAIBMA_CORP_ZRR,
    MACRO_UPDATE_DATA2, MORNING_PART2_FILE_MAPPINGS,
)

log = get_logger("Run_morning_ThaiBMA_part2")

_proj_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
env_path = os.environ.get('RISK_ENV_PATH') or os.path.join(_proj_root, 'scripts', '.env')
load_dotenv(env_path)
password = os.getenv('pass')
username = os.getenv('user')

download_dir = MORNING_DL_DIR


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


def main():
    log.info("=== Morning ThaiBMA Part 2 started ===")
    log.info(f"Loaded credentials from: {env_path}")

    edge_options = Options()
    edge_options.add_argument("--start-maximized")
    edge_options.add_argument("--force-device-scale-factor=0.35")
    edge_options.add_argument("--high-dpi-support=0.35")
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
                time.sleep(1.5)
        log.info("ThaiBMA login successful")

        driver.get(THAIBMA_ZRR)
        driver.execute_script("document.body.style.zoom = '0.3'")
        driver.execute_script("document.body.style['-webkit-transform'] = 'scale(0.5)';")
        time.sleep(2)

        for i in range(5):
            try:
                btn = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(text(),'Download all TTM and Date')]")
                    )
                )
                driver.execute_script("arguments[0].click();", btn)
                print("กดปุ่ม Download สำเร็จ")
                break
            except Exception:
                print("ยังไม่เจอปุ่ม ลองใหม่...", i)

        time.sleep(35)

        driver.get(THAIBMA_SHORTTERM)
        time.sleep(3)
        time.sleep(2)

        icon = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR,
                "table > tbody > tr > td:nth-child(13) > div > button > i"))
        )
        icon.click()
        time.sleep(7)

        driver.get(THAIBMA_BOND)
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

        driver.get(THAIBMA_MTM_CORP)
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

        driver.get(THAIBMA_COMPOSITE)
        time.sleep(2)

        button_com = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR,
                "table > tbody > tr > td:nth-child(11) > div > button > i"))
        )
        button_com.click()
        time.sleep(7)

        driver.get(THAIBMA_CORP_ZRR)
        time.sleep(2)

        button_corp_zrr = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR,
                "div.container > div.main-content > div.essential-div > div > div:nth-child(7) > div > div > div > table > tbody > tr:nth-child(1) > td:nth-child(8) > div > button > i"))
        )
        button_corp_zrr.click()
        time.sleep(7)
    finally:
        driver.quit()

    for _from, _to in MORNING_PART2_FILE_MAPPINGS:
        log.info(f"Copying file: {_from} → {_to}")
        copy_latest_file(MORNING_DL_DIR, MORNING_DATA_DIR, _from, _to)
    log.info("All file copies complete")

    run_excel_macro(POWER_AUTOMATE, MACRO_UPDATE_DATA2)
    log.info("=== Morning ThaiBMA Part 2 completed ===")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log.critical(f"Script failed: {e}", exc_info=True)
        send_failure_alert("Run_morning_ThaiBMA_part2", str(e))
        sys.exit(1)
