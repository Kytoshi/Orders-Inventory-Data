import os
import sys
import glob
import time
import threading
from helpers import get_current_dir, subtract_one_business_day, today_date, wait_for_element
from file_utils import remove_old_files, wait_for_download
from send2trash import send2trash
from excel_manager import excel_manager

from logger import logger
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from webdriver_manager.chrome import ChromeDriverManager

# Lock to prevent concurrent chromedriver initialization
_chromedriver_lock = threading.Lock()

def create_Driver(download_dir):
    chrome_options = Options()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--headless=new")
    chrome_options.page_load_strategy = 'eager'
    
    chrome_option_args = [
        '--start-maximized',
        '--window-size=1920,1080',
        '--disable-popup-blocking',
        '--disable-gpu',
        '--disable-dev-shm-usage',
        '--disable-notifications',
        '--disable-infobars',
        '--no-sandbox',
        '--disable-blink-features=AutomationControlled',
        '--disable-extensions',
        '--disable-plugins',
    ]
    for arg in chrome_option_args:
        chrome_options.add_argument(arg)

    driver = None
    
    # If running as PyInstaller executable, try bundled chromedriver first
    if getattr(sys, 'frozen', False):
        bundled_chromedriver = os.path.join(sys._MEIPASS, 'chromedriver.exe')
        
        if os.path.exists(bundled_chromedriver):
            try:
                logger.info("Using bundled chromedriver...")
                service = Service(bundled_chromedriver)
                driver = webdriver.Chrome(service=service, options=chrome_options)
                logger.info("✓ Bundled chromedriver working")
            except Exception as e:
                logger.warning(f"Bundled chromedriver incompatible: {e}")
                logger.info("Downloading compatible chromedriver (this may take a moment)...")
                driver = None  # Will fallback below
    
    # Fallback: use webdriver-manager (for script mode or if bundled failed)
    if driver is None:
        # Use lock to prevent concurrent chromedriver downloads/initialization
        with _chromedriver_lock:
            try:
                logger.info("Downloading/updating chromedriver...")
                chromedriver_path = ChromeDriverManager().install()
                service = Service(chromedriver_path)
                driver = webdriver.Chrome(service=service, options=chrome_options)
                logger.info("✓ Chromedriver ready")
            except Exception as e:
                logger.error(f"Failed to initialize chromedriver: {e}")
                raise
    
    # Configure driver
    driver.set_page_load_timeout(600)
    driver.set_script_timeout(600)
    driver.implicitly_wait(30)
    driver.execute_cdp_cmd("Network.enable", {})
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": download_dir
    })
    
    return driver

def open_PDBS_Homepage():
    from config import get_web_config
    web_cfg = get_web_config()
    driver = create_Driver(get_current_dir())
    driver.get(web_cfg["pdbs_url"])
    return driver

def login_credentials(username, password, driver):
    user_field = driver.find_element(By.ID, "txtUserName")
    pass_field = driver.find_element(By.ID, "xPWD")

    user_field.send_keys(username)
    pass_field.send_keys(password)
    driver.find_element(By.ID, "btnSubmit").click()

    logger.info("Waiting for login response...")
    
    try:
        # Wait for successful navigation (an element that appears ONLY on the next page)
        # e.g., look for an element that only exists after successful login
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.TAG_NAME, "a"))  # or a more specific element on the next page
        )
        logger.info("✓ Login successful")
    
    except TimeoutException:
        # If we get here, page didn't navigate—check for error message
        try:
            error_element = driver.find_element(By.CLASS_NAME, "text-danger")
            if error_element.is_displayed() and error_element.text.strip():
                logger.error(f"Login failed: {error_element.text.strip()}")
                raise ValueError("Invalid username or password. Please check your credentials.")
        except NoSuchElementException:
            logger.error("Login timeout - no success page and no error message found")
            raise ValueError("Login failed - server not responding")

def get_MatShortage_Data(username, password):
    driver = None
    try:
        driver = open_PDBS_Homepage()
        logger.info("Logging in for MatShortage Data...")
        login_credentials(username=username, password=password, driver=driver)

        wait_for_element(driver, By.TAG_NAME, 'a')

        # Navigate to the first link - wait until the specific anchor is clickable
        logger.info("Navigating to AODN Process Control...")
        try:
            xpath = "//a[@href='javascript:onClickTaskMenu(\"DNProcessRedirect.asp\", 351)']"
            link = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            try:
                link.click()
            except Exception:
                # Try scrolling into view and retrying the normal click
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
                time.sleep(0.4)
                try:
                    link.click()
                except Exception:
                    # Last resort: use JS click to avoid interception
                    driver.execute_script("arguments[0].click();", link)

            logger.info("AODN Process Control Page Loaded.")
        except TimeoutException:
            logger.error("AODN Process Control link not found or not clickable via XPath, falling back to scanning anchors.")
            # Fallback: scan anchors and attempt robust click per anchor
            links = driver.find_elements(By.TAG_NAME, 'a')
            for link in links:
                try:
                    if link.get_attribute('href') == 'javascript:onClickTaskMenu("DNProcessRedirect.asp", 351)':
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
                        time.sleep(0.4)
                        try:
                            link.click()
                        except Exception:
                            try:
                                driver.execute_script("arguments[0].click();", link)
                            except Exception as e:
                                logger.error(f"Could not click fallback link: {e}")
                        break
                except Exception:
                    continue

        driver.find_element(by=By.ID, value="Submit").click()

        driver.find_element(by=By.ID, value="pnlMartShortage").click()

        driver.find_element(by=By.ID, value="MainContent_btnExportExcel").click()
        time.sleep(2)  # Give the download a moment to start

        logger.info("Waiting for MatShortageRpt file download...")
        wait_for_download(file="MatShortageRpt", timeout=300)
        logger.info("MatShortageRpt Excel file downloaded.")

        mat_files = glob.glob(f"{get_current_dir()}/MatShortageRpt_*.xlsx")

        # Rename the most recent .xlsx file if necessary
        if mat_files:
            logger.info("MatShortageRpt file found, processing rename...")
            first_converted_path = os.path.splitext(mat_files[0])[0] + ".xlsx"
            new_file_path = os.path.join(get_current_dir(), "MatShortageRpt.xlsx")

            if not os.path.exists(new_file_path):
                os.rename(first_converted_path, new_file_path)
                logger.info(f"MatShortageRpt file renamed to {new_file_path}.")
            else:
                logger.error(f"The file {new_file_path} already exists.")
    except ValueError as ve:
        # Authentication error - log and re-raise
        logger.error(f"Failed to login for MatShortage Data: {ve}")
        if driver:
            driver.quit()
        raise  # Re-raise to stop the thread
    except Exception as e:
        logger.error(f"Error in MatShortage download: {e}")
        if driver:
            driver.quit()
        raise
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass

def navigate_DailyReport(username, password):
    driver = open_PDBS_Homepage()
    logger.info("Logging in for DailyReports...")
    login_credentials(username=username, password=password, driver=driver)

    links = driver.find_elements(by=By.TAG_NAME, value='a')

    for link in links:
        if link.get_attribute('href') == 'javascript:onClickTaskMenu("OrdReport.asp", 65)':
            link.click()
            break
    return driver

def get_DailyReport_Completed(prevDate, driver):
    DailyOrders_date_field = wait_for_element(driver, By.NAME, "Date")
    DailyOrders_date_field.clear()

    DailyOrders_date_field.send_keys(prevDate.strftime("%m/%d/%Y"))

    driver.execute_script("ChgDate()")

    download_start_time = time.time()  # Record time before download

    try:
        # Find the link by its visible text and click it
        link = driver.find_element(By.LINK_TEXT, "Order Fulfillment Report")
        link.click()

    except Exception as e:
        logger.error(f"Error: {e}")

    wait_for_download(file="DailyReport.xls", timeout=300, after_time=download_start_time)

    dailyRpt_Initial_File = os.path.join(get_current_dir(), "DailyReport.xls")
    if os.path.exists(dailyRpt_Initial_File):

        DailyRpt_xlsx_path = os.path.join(get_current_dir(), "DailyReport Completed.xlsx")

        try:
            success = excel_manager.convert_xls_to_xlsx(dailyRpt_Initial_File, DailyRpt_xlsx_path)
            if success:
                logger.info("DailyReport.xls has been converted to DailyReport Completed.xlsx")
                if os.path.exists(dailyRpt_Initial_File):
                    send2trash(dailyRpt_Initial_File)
                    logger.info("Original DailyReport.xls file has been deleted")
                else:
                    logger.error("DailyReport.xls does not exist")
            else:
                logger.error("Failed to convert DailyReport.xls to XLSX")
        except Exception as e:
            logger.error(f"Error processing Completed report: {e}")
    else:
        pass

def get_DailyReport_Incompletes(currDate, driver):
    DailyOrders_date_field = wait_for_element(driver, By.NAME, "Date")
    DailyOrders_date_field.clear()

    DailyOrders_date_field.send_keys(currDate.strftime("%m/%d/%Y"))

    driver.execute_script("ChgDate()")

    download_start_time = time.time()  # Record time before download

    try:
        # Find the link by its visible text and click it
        link = driver.find_element(By.LINK_TEXT, "Order Fulfillment Report")
        link.click()

    except Exception as e:
        logger.error(f"Error: {e}")

    wait_for_download(file="DailyReport.xls", timeout=300, after_time=download_start_time)

    dailyRpt_Initial_File = os.path.join(get_current_dir(), "DailyReport.xls")
    if os.path.exists(dailyRpt_Initial_File):

        DailyRpt_xlsx_path = os.path.join(get_current_dir(), "DailyReport Incompletes.xlsx")

        try:
            success = excel_manager.convert_xls_to_xlsx(dailyRpt_Initial_File, DailyRpt_xlsx_path)
            if success:
                logger.info("DailyReport.xls converted to DailyReport Incompletes.xlsx")
                if os.path.exists(dailyRpt_Initial_File):
                    send2trash(dailyRpt_Initial_File)
                    logger.info("Original DailyReport.xls file has been deleted")
                else:
                    logger.error("DailyReport.xls does not exist")
            else:
                logger.error("Failed to convert DailyReport.xls to XLSX")
        except Exception as e:
            logger.error(f"Error processing Incompletes report: {e}")
    else:
        pass

def get_DailyReport_Billing(currDate, driver):
    DailyOrders_date_field = wait_for_element(driver, By.NAME, "Date")
    DailyOrders_date_field.clear()

    DailyOrders_date_field.send_keys(currDate.strftime("%m/%d/%Y"))

    driver.execute_script("ChgDate()")

    download_start_time = time.time()  # Record time before download

    try:
        # Find the link by its visible text and click it
        link = driver.find_element(By.LINK_TEXT, "Report in Excel")
        link.click()

    except Exception as e:
        logger.error(f"Error: {e}")

    wait_for_download(file="DailyReport.xls", timeout=300, after_time=download_start_time)

    dailyRpt_Initial_File = os.path.join(get_current_dir(), "DailyReport.xls")
    if os.path.exists(dailyRpt_Initial_File):

        DailyRpt_xlsx_path = os.path.join(get_current_dir(), "Billing Only.xlsx")

        try:
            success = excel_manager.convert_xls_to_xlsx(dailyRpt_Initial_File, DailyRpt_xlsx_path)
            if success:
                logger.info("DailyReport.xls converted to Billing Only.xlsx")
                if os.path.exists(dailyRpt_Initial_File):
                    send2trash(dailyRpt_Initial_File)
                    logger.info("Original DailyReport.xls file has been deleted")
                else:
                    logger.error("DailyReport.xls does not exist")
            else:
                logger.error("Failed to convert DailyReport.xls to XLSX")
        except Exception as e:
            logger.error(f"Error processing Billing report: {e}")
    else:
        pass

def run_all_DailyReport_downloads(username, password):
    driver = None
    try:
        remove_old_files(folder_path=get_current_dir())
        driver = navigate_DailyReport(username=username, password=password)

        today = today_date()
        prev_date = subtract_one_business_day(today)

        get_DailyReport_Billing(currDate=today, driver=driver)
        get_DailyReport_Incompletes(currDate=today, driver=driver)
        get_DailyReport_Completed(prevDate=prev_date, driver=driver)

        logger.info("✓ All DailyReport downloads completed")
        
    except ValueError as ve:
        # Authentication error - log and re-raise
        logger.error(f"Failed to login for DailyReport: {ve}")
        if driver:
            driver.quit()
        raise  # Re-raise to stop the thread
    except Exception as e:
        logger.error(f"Error in DailyReport downloads: {e}")
        if driver:
            driver.quit()
        raise
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass
def main(username, password):
    errors = []
    
    def thread1_wrapper():
        try:
            get_MatShortage_Data(username, password)
        except Exception as e:
            errors.append(('MatShortage', e))
    
    def thread2_wrapper():
        try:
            run_all_DailyReport_downloads(username, password)
        except Exception as e:
            errors.append(('DailyReport', e))
    
    # Start both threads
    thread1 = threading.Thread(target=thread1_wrapper)
    thread2 = threading.Thread(target=thread2_wrapper)
    
    thread1.start()
    thread2.start()
    
    # Wait for both to complete
    thread1.join()
    thread2.join()
    
    # Check if any errors occurred
    if errors:
        error_messages = [f"{name}: {str(err)}" for name, err in errors]
        combined_error = "\n".join(error_messages)
        logger.error(f"Downloads failed:\n{combined_error}")
        raise ValueError(combined_error)
    
    logger.info("✓ All downloads completed successfully")


if __name__ == "__main__":
    import sys
    username = sys.argv[1]
    password = sys.argv[2]
    main(username=username, password=password)