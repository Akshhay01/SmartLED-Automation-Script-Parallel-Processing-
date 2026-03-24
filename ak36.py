import os
import time
import logging
import pandas as pd
import numpy as np
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, ElementClickInterceptedException, WebDriverException
)
import concurrent.futures
from tqdm import tqdm
import threading

SCRIPT_VERSION = "1.0"
RUN_TIMESTAMP = time.strftime("%Y-%m-%d %H:%M:%S")

logging.basicConfig(
    filename="fault_log.txt",
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

def create_driver():
    options = webdriver.EdgeOptions()
    options.add_argument("--headless")  # Headless mode active
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    driver = webdriver.Edge(options=options)
    driver.set_page_load_timeout(60)
    return driver

def safe_click(driver, by, value):
    wait = WebDriverWait(driver, 20)
    element = wait.until(EC.element_to_be_clickable((by, value)))
    driver.execute_script("arguments[0].scrollIntoView(true);", element)
    try:
        element.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", element)

def safe_input(element, value):
    element.clear()
    element.send_keys(value)

def get_excel_value(row, col):
    val = row.get(col, "")
    if pd.isna(val):
        return ""
    return str(val).strip()

def login(driver):
    wait = WebDriverWait(driver, 30)
    driver.get("https://smartled.in/Installation/index.php")
    wait.until(EC.element_to_be_clickable((By.ID, "input-1"))).send_keys("jaiauto.******")
    driver.find_element(By.ID, "input-2").send_keys("*****")
    safe_click(driver, By.XPATH, "//button[text()='Login']")
    wait.until(EC.url_contains("Dashboard.php"))

def process_row(idx, row, driver):
    wait = WebDriverWait(driver, 30)
    try:
        driver.get("https://smartled.in/Installation/Dashboard.php")
        wait.until(EC.presence_of_element_located((By.ID, "IsAction")))

        Select(driver.find_element(By.ID, "IsAction")).select_by_visible_text("Luminary Pole Id")

        pole_id = get_excel_value(row, "Pole ID")
        pole_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter Luminary Pole Id']")))
        pole_input.clear()
        pole_input.send_keys(pole_id)
        safe_click(driver, By.XPATH, "//button[text()='Next']")

        time.sleep(2)  # Let page update

        page_source = driver.page_source
        if "Luminary Already Installed!" in page_source:
            return idx, "Luminary Already Installed!"

        serial_xpath = "//span[contains(text(), 'Luminary Serial No:')]/parent::td"
        td_elem = wait.until(EC.presence_of_element_located((By.XPATH, serial_xpath)))
        serial_text = td_elem.text.strip()

        import re
        match = re.search(r"Luminary Serial No:\s*(.+)", serial_text)
        if not match:
            return idx, "Luminary Serial No not found on page"
        serial_on_site = match.group(1).strip().strip('"')
        serial_excel = get_excel_value(row, "Luminary Serial No")
        if serial_on_site != serial_excel:
            return idx, "Luminary Serial No not match"

        district_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "DistrictId")))
        Select(district_dropdown).select_by_visible_text(get_excel_value(row, "District"))
        block_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "BlockId")))
        time.sleep(1)
        Select(block_dropdown).select_by_visible_text(get_excel_value(row, "Block"))
        panchayat_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "PanchayatId")))
        time.sleep(1)
        Select(panchayat_dropdown).select_by_visible_text(get_excel_value(row, "Panchayat"))
        network_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "NetworkId")))
        time.sleep(1)
        Select(network_dropdown).select_by_visible_text(get_excel_value(row, "Network"))
        safe_click(driver, By.ID, "btn")

        wait.until(EC.presence_of_element_located((By.NAME, "BatSerialNo")))
        safe_input(driver.find_element(By.NAME, "BatSerialNo"), get_excel_value(row, "Battery Serial No"))
        safe_input(driver.find_element(By.NAME, "PanelSerialNo"), get_excel_value(row, "Solar Panel Serial No"))
        safe_input(driver.find_element(By.NAME, "Village"), get_excel_value(row, "Ward No"))
        safe_input(driver.find_element(By.NAME, "PRDPoleId"), get_excel_value(row, "PRD Pole ID"))
        safe_click(driver, By.ID, "btn")

        time.sleep(3)

        page_source = driver.page_source
        success_msg = "Luminary Installation Details Added Successfully !"
        if success_msg in page_source:
            return idx, "Successfully"
        else:
            try:
                msg_elem = driver.find_element(By.CSS_SELECTOR, ".alert, .error, #message, p")
                msg_text = msg_elem.text.strip()
            except Exception:
                msg_text = "Unknown response"
            return idx, msg_text

    except Exception as e:
        logging.error(f"Error processing row {idx}: {e}")
        raise

def robust_worker(chunk_df, progress_bar, progress_lock):
    driver = create_driver()
    login(driver)
    results = []
    for idx, row in chunk_df.iterrows():
        retry_count = 0
        while retry_count < 3:
            try:
                idx_actual, remark = process_row(idx, row, driver)
                results.append((idx_actual, remark))
                with progress_lock:
                    progress_bar.update(1)
                break
            except (WebDriverException, TimeoutException) as e:
                logging.error(f"Restarting browser at row {idx}: {e}")
                try:
                    driver.quit()
                except Exception:
                    pass
                time.sleep(3)
                driver = create_driver()
                login(driver)
                retry_count += 1
            except Exception as e:
                results.append((idx, f"Error: {e}"))
                with progress_lock:
                    progress_bar.update(1)
                break
    try:
        driver.quit()
    except Exception:
        pass
    return results

def main_parallel():
    print(f"Script version: {SCRIPT_VERSION}")
    print(f"Run started at: {RUN_TIMESTAMP}")

    Tk().withdraw()
    excel_path = askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
    if not excel_path:
        print("No file selected. Exiting.")
        return

    df = pd.read_excel(excel_path)
    df.dropna(subset=["Pole ID"], inplace=True)
    if "Remarks" not in df.columns:
        df.insert(len(df.columns), "Remarks", "")
    df["Script Version"] = SCRIPT_VERSION
    df["Processed Timestamp"] = RUN_TIMESTAMP

    chunks = np.array_split(df, 3)
    progress_bar = tqdm(total=len(df), desc="Processing rows")
    progress_lock = threading.Lock()

    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
        futures = [
            executor.submit(robust_worker, chunk, progress_bar, progress_lock)
            for chunk in chunks
        ]
        for future in concurrent.futures.as_completed(futures):
            results = future.result()
            for idx_actual, remark in results:
                df.at[idx_actual, "Remarks"] = remark

    progress_bar.close()
    output_path = os.path.join(os.path.dirname(excel_path), "final.xlsx")
    df.to_excel(output_path, index=False)
    print(f"\nResults saved to {output_path}")

if __name__ == "__main__":
    main_parallel()
