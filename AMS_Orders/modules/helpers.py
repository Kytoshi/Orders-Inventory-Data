from logger import logger
from datetime import datetime, timedelta
from contextlib import contextmanager
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
import win32com.client
import subprocess
import os
import pythoncom


@contextmanager
def com_context():
    """
    Context manager for proper COM initialization and cleanup.
    Use this in any thread that needs to access COM objects (SAP, Excel, etc.)
    """
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            logger.debug(f"COM cleanup: {e}")

def get_current_dir():
    return os.getcwd()

def today_date():
    return datetime.today().date()


def get_company_holidays(year):
    """
    Calculate company holidays for a given year.
    Handles both fixed holidays and floating holidays (like MLK Day, Thanksgiving).
    """
    holidays = set()

    # Fixed holidays (same date every year)
    holidays.add(datetime(year, 1, 1))    # New Year's Day
    holidays.add(datetime(year, 6, 19))   # Juneteenth
    holidays.add(datetime(year, 7, 4))    # Independence Day
    holidays.add(datetime(year, 12, 24))  # Christmas Eve
    holidays.add(datetime(year, 12, 25))  # Christmas Day

    # Check if July 4th falls on weekend, add observed day
    july_4 = datetime(year, 7, 4)
    if july_4.weekday() == 5:  # Saturday -> Friday observed
        holidays.add(datetime(year, 7, 3))
    elif july_4.weekday() == 6:  # Sunday -> Monday observed
        holidays.add(datetime(year, 7, 5))

    # MLK Day: Third Monday of January
    jan_first = datetime(year, 1, 1)
    first_monday = jan_first + timedelta(days=(7 - jan_first.weekday()) % 7)
    if jan_first.weekday() == 0:
        first_monday = jan_first
    mlk_day = first_monday + timedelta(weeks=2)
    holidays.add(mlk_day)

    # Presidents' Day: Third Monday of February
    feb_first = datetime(year, 2, 1)
    first_monday = feb_first + timedelta(days=(7 - feb_first.weekday()) % 7)
    if feb_first.weekday() == 0:
        first_monday = feb_first
    presidents_day = first_monday + timedelta(weeks=2)
    holidays.add(presidents_day)

    # Memorial Day: Last Monday of May
    may_last = datetime(year, 5, 31)
    memorial_day = may_last - timedelta(days=(may_last.weekday() - 0) % 7)
    holidays.add(memorial_day)

    # Labor Day: First Monday of September
    sep_first = datetime(year, 9, 1)
    first_monday = sep_first + timedelta(days=(7 - sep_first.weekday()) % 7)
    if sep_first.weekday() == 0:
        first_monday = sep_first
    labor_day = first_monday
    holidays.add(labor_day)

    # Thanksgiving: Fourth Thursday of November
    nov_first = datetime(year, 11, 1)
    first_thursday = nov_first + timedelta(days=(3 - nov_first.weekday()) % 7)
    if nov_first.weekday() == 3:
        first_thursday = nov_first
    thanksgiving = first_thursday + timedelta(weeks=3)
    holidays.add(thanksgiving)

    # Day after Thanksgiving
    holidays.add(thanksgiving + timedelta(days=1))

    return holidays


def subtract_one_business_day(date):
    """
    Subtract one business day from the given date, considering weekends and holidays.
    Uses dynamically calculated holidays that work for any year.
    """
    date -= timedelta(days=1)
    logger.info(f"Previous Date is {date.strftime('%m/%d/%Y')}")
    logger.info(f"Checking if {date.strftime('%m/%d/%Y')} is a business day...")

    # Get holidays for this year and adjacent years (in case we cross year boundary)
    holidays = get_company_holidays(date.year)
    holidays.update(get_company_holidays(date.year - 1))

    while True:
        # Convert date to datetime for comparison with holidays set
        date_as_datetime = datetime(date.year, date.month, date.day)

        if date_as_datetime in holidays:
            logger.info(f"{date.strftime('%m/%d/%Y')} is a holiday, going back one more day.")
            date -= timedelta(days=1)
            continue

        if date.weekday() == 5:  # Saturday
            logger.info(f"{date.strftime('%m/%d/%Y')} is Saturday, going back one more day.")
            date -= timedelta(days=1)
            continue
        elif date.weekday() == 6:  # Sunday
            logger.info(f"{date.strftime('%m/%d/%Y')} is Sunday, going back two days.")
            date -= timedelta(days=2)
            continue

        break  # Found valid business day

    logger.info(f"Final business day found: {date.strftime('%m/%d/%Y')}")
    return date

def SAP_Init():
    """
    Initialize and return an SAP connection object.
    Waits for SAP to be ready before accessing connections.
    """
    logger.info("Initializing SAP connection...")
    max_retries = 10
    retry_delay = 2

    for attempt in range(max_retries):
        try:
            SAP_GUI_AUTO = win32com.client.GetObject('SAPGUI')

            if isinstance(SAP_GUI_AUTO, win32com.client.CDispatch):
                application = SAP_GUI_AUTO.GetScriptingEngine

                # Check if we have any connections yet
                if application.Children.Count == 0:
                    if attempt < max_retries - 1:
                        logger.warning(f"No SAP connections available yet, retrying in {retry_delay}s... (attempt {attempt + 1}/{max_retries})")
                        time.sleep(retry_delay)
                        continue
                    else:
                        raise Exception("SAP opened but no connection established. Check SAP login.")

                connection = application.Children(0)
                logger.info(f"✓ SAP connection established (found {application.Children.Count} connection(s))")
                return connection

        except Exception as e:
            if attempt < max_retries - 1:
                logger.warning(f"SAP connection attempt {attempt + 1} failed: {e}, retrying...")
                time.sleep(retry_delay)
            else:
                logger.error(f"Could not initialize SAP connection after {max_retries} attempts: {e}")
                raise

def Open_SAP(username, password):
    """
    Open SAP GUI and log in with the provided username and password.
    Once Logged in Successfully, creates 3 SAP session windows.
    """
    from config import get_sap_config
    sap_cfg = get_sap_config()

    exe_path = sap_cfg["saplogon_path"]
    process = subprocess.Popen(exe_path)
    time.sleep(7)

    sapshcut_path = sap_cfg["sapshcut_path"]
    system = sap_cfg["system"]
    client = sap_cfg["client"]
    language = sap_cfg.get("language", "EN")
    # Note: os.system is used here with config-controlled arguments for SAP shortcut launch
    command = f'"{sapshcut_path}" -system={system} -client={client} -user={username} -pw={password} -language={language}'
    os.system(command)
    time.sleep(5)

    connection = SAP_Init()

    # Wait for the first session to be fully available
    max_attempts = 15
    session = None
    for attempt in range(max_attempts):
        try:
            if connection.Children.Count > 0:
                session = connection.Children(0)
                # Try to access the session to verify it's ready
                session.findById("wnd[0]")
                logger.info("✓ Initial SAP session is ready")
                break
            else:
                logger.warning(f"Waiting for first SAP session... (attempt {attempt + 1}/{max_attempts})")
                time.sleep(2)
        except Exception as e:
            if attempt < max_attempts - 1:
                logger.warning(f"Session not ready yet, retrying... (attempt {attempt + 1}/{max_attempts})")
                time.sleep(2)
            else:
                raise Exception(f"Failed to get initial SAP session after {max_attempts} attempts: {e}")

    if session is None:
        raise Exception("Could not establish initial SAP session")

    try:
        logger.info("Connecting to Initial SAP session...")
        session.findById("wnd[0]").maximize()

        logger.info("Creating two additional SAP sessions...")
        session.findById("wnd[0]").sendVKey(74) # Create new session through the intial session
        time.sleep(3)  # Wait for session to be created

        # Verify second session was created
        if connection.Children.Count < 2:
            raise Exception("Failed to create second SAP session")
        logger.info(f"✓ Second SAP session created (Total: {connection.Children.Count})")

        session.findById("wnd[0]").sendVKey(74)
        time.sleep(3)  # Wait for session to be created

        # Verify third session was created
        if connection.Children.Count < 3:
            raise Exception("Failed to create third SAP session")
        logger.info(f"✓ Third SAP session created (Total: {connection.Children.Count})")

    except Exception as e:
        logger.error(f"Error creating SAP sessions: {e}")
        raise

def wait_for_element(driver, by, value, total_wait=480, check_interval=10):
    try:
        element = WebDriverWait(driver, total_wait, check_interval).until(EC.presence_of_element_located((by, value)))
        return element
    except TimeoutException:
        logger.error(f"Timeout waiting for element by {by} with value {value}")
        return None