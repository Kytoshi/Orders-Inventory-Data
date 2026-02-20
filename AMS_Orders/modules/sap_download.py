from helpers import *
from threading import Thread
import warnings
from file_utils import find_and_copy_file
from logger import logger
from config import get_sap_config
import psutil
from excel_manager import excel_manager

# Thread timeout in seconds (10 minutes)
THREAD_TIMEOUT = 600


def close_sap():
    """Close SAP using proper escalation: graceful → gentle → force"""
    with com_context():
        # STEP 1: Graceful close (best)
        try:
            logger.info("Attempting graceful SAP close...")
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)

            # Close all sessions
            session_count = connection.Children.Count
            logger.info(f"Closing {session_count} SAP session(s)...")

            for i in range(session_count - 1, -1, -1):  # Close in reverse order
                try:
                    session = connection.Children(i)
                    session.findById("wnd[0]").close()
                    time.sleep(0.5)

                    # Handle logoff confirmation dialog if it appears
                    try:
                        # Check if confirmation window exists
                        confirm_window = session.findById("wnd[1]")
                        if confirm_window:
                            # Press "Yes" or "OK" button to confirm logoff
                            try:
                                # Try different possible button IDs for confirmation
                                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()  # Yes button
                                logger.info(f"✓ Confirmed logoff for session {i}")
                            except Exception as e:
                                logger.debug(f"OPTION1 button not found: {e}")
                                try:
                                    session.findById("wnd[1]/tbar[0]/btn[0]").press()  # OK button
                                    logger.info(f"✓ Confirmed logoff for session {i}")
                                except Exception as e:
                                    logger.debug(f"btn[0] button not found: {e}")
                            time.sleep(0.5)
                    except Exception as e:
                        logger.debug(f"No confirmation dialog for session {i}: {e}")

                    logger.info(f"✓ Closed session {i}")
                except Exception as e:
                    logger.warning(f"Could not close session {i}: {e}")

            # Close connection
            connection.CloseConnection()
            logger.info("✓ SAP closed gracefully")
            return True

        except Exception as e:
            logger.warning(f"Graceful close failed: {e}")

        # STEP 2: Force close SAP
        try:
            logger.warning("Force closing SAP...")
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] and 'sap' in proc.info['name'].lower():
                    proc.terminate()  # Try gentle first
                    try:
                        proc.wait(timeout=3)
                        logger.info(f"✓ Terminated: {proc.info['name']}")
                    except psutil.TimeoutExpired:
                        proc.kill()  # Force kill if needed
                        logger.info(f"✓ Force killed: {proc.info['name']}")
            return True

        except Exception as e:
            logger.error(f"Force close failed: {e}")
            return False

def close_excel():
    with com_context():
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            for wb in excel.workbooks:
                wb.Close(SaveChanges=True)
            logger.info("Excel Workbooks closed successfully.")
        except Exception as e:
            logger.warning(f"Error closing Excel: {e}")

def MO_Backorders(today_str):
    with com_context():
        logger.info("Starting MO BACKORDERS Transaction...")
        sap_cfg = get_sap_config()

        connection = SAP_Init()
        session = connection.Children(0)

        session.findById("wnd[0]").iconify()
        # session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "MB25"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/tbar[1]/btn[17]").press()

        session.findById("wnd[1]/usr/txtV-LOW").text = "MO CHECKER"
        session.findById("wnd[1]/usr/txtENAME-LOW").text = sap_cfg["variant_username"]
        session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 10

        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/usr/ctxtBDTER-HIGH").text = today_str
        session.findById("wnd[0]/usr/ctxtBDTER-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtBDTER-HIGH").caretPosition = 8

        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = get_current_dir()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MB25 Backorders.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        logger.info("MO BACKORDERS (MB25) transaction completed.")

def MB51(today_str, yesterday_str):
    with com_context():
        logger.info("Starting MB51 Transaction...")
        sap_cfg = get_sap_config()

        connection = SAP_Init()
        session = connection.Children(1)

        session.findById("wnd[0]").iconify()
        # session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "MB51"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/tbar[1]/btn[17]").press()

        session.findById("wnd[1]/usr/txtV-LOW").text = "MB51 CHECKER"
        session.findById("wnd[1]/usr/txtENAME-LOW").text = sap_cfg["variant_username"]
        session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 12

        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = yesterday_str
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = today_str
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 6

        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = get_current_dir()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MB51.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        logger.info("MB51 transaction completed.")

def DAILY_MO_MB25(today_str, yesterday_str):
    with com_context():
        logger.info("Starting Daily MO MB25 Transaction...")
        sap_cfg = get_sap_config()

        connection = SAP_Init()
        session = connection.Children(2)

        session.findById("wnd[0]").iconify()
        # session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "MB25"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/tbar[1]/btn[17]").press()

        session.findById("wnd[1]/usr/txtV-LOW").text = "DAILY MO MB25"
        session.findById("wnd[1]/usr/txtENAME-LOW").text = sap_cfg["variant_username"]
        session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 13

        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/usr/ctxtBDTER-LOW").text = yesterday_str
        session.findById("wnd[0]/usr/ctxtBDTER-HIGH").text = today_str
        session.findById("wnd[0]/usr/ctxtBDTER-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtBDTER-HIGH").caretPosition = 6

        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = get_current_dir()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "DAILY MO MB25.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13

        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        time.sleep(5)

        logger.info("Daily MO MB25 (MB25) transaction completed.")

        logger.info("Backing up Daily MO MB25 file...")
        find_and_copy_file(
            source_folder=get_current_dir(),
            destination_folder=os.path.join(get_current_dir(), "Backup"),
            file_prefix="DAILY MO MB25"
        )

def main(username, password):
    with com_context():
        warnings.filterwarnings("ignore", category=ResourceWarning)

        today = today_date()
        today_str = today.strftime("%m/%d/%Y")
        prev_date = subtract_one_business_day(today)
        prev_date_str = prev_date.strftime("%m/%d/%Y")

        logger.info(f"Today's Date: {today_str}")
        logger.info(f"Last Workday: {prev_date_str}")

        sap_process = None

        sap_cfg = get_sap_config()
        try:
            # Start SAP (no Excel needed!)
            sap_process = subprocess.Popen(
                sap_cfg["saplogon_path"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            time.sleep(5)

            # Login
            Open_SAP(username, password)
            logger.info("Waiting for SAP to fully initialize...")
            time.sleep(5)

            # Verify all 3 sessions are ready
            connection = SAP_Init()
            if connection.Children.Count < 3:
                raise Exception(f"Expected 3 SAP sessions, but only found {connection.Children.Count}. Cannot proceed.")
            logger.info(f"✓ All 3 SAP sessions verified and ready")

            # Run threads
            thread1 = Thread(target=MO_Backorders, args=(today_str,), name="MO_Backorders")
            thread2 = Thread(target=MB51, args=(today_str, prev_date_str), name="MB51")
            thread3 = Thread(target=DAILY_MO_MB25, args=(today_str, prev_date_str), name="DAILY_MO_MB25")

            for t in [thread1, thread2, thread3]:
                t.start()

            # Wait for threads with timeout
            timed_out_threads = []
            for t in [thread1, thread2, thread3]:
                t.join(timeout=THREAD_TIMEOUT)
                if t.is_alive():
                    logger.error(f"Thread {t.name} timed out after {THREAD_TIMEOUT}s")
                    timed_out_threads.append(t.name)

            if timed_out_threads:
                raise TimeoutError(f"Threads did not complete: {', '.join(timed_out_threads)}")

            logger.info("✓ All SAP transactions completed")
            time.sleep(5)

            # Files are saved by SAP directly - no Excel interaction needed!
            logger.info("✓ SAP files exported successfully")

            # Close any Excel windows that SAP might have opened
            try:
                close_excel()
                # Use thread-safe Excel manager to close any Excel instance
                excel_manager.release_excel(force_quit=True)
                logger.info("✓ Closed Excel instance")
            except Exception as e:
                logger.warning(f"Excel cleanup note: {e}")

            # Close SAP
            close_sap()

            logger.info("✓ All cleanup completed successfully")

        except Exception as e:
            logger.error(f"Error in main: {e}")
            raise
        finally:
            # Close SAP process
            try:
                if sap_process and sap_process.poll() is None:
                    sap_process.terminate()
                    sap_process.wait(timeout=5)
            except Exception as e:
                logger.debug(f"Error terminating SAP process: {e}")
                try:
                    if sap_process:
                        sap_process.kill()
                except Exception as e:
                    logger.debug(f"Error killing SAP process: {e}")