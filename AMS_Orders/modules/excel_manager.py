"""
Thread-safe Excel COM object manager to prevent conflicts when
multiple threads need to interact with Excel simultaneously.
"""
import win32com.client
import pythoncom
import threading
import time
from logger import logger


class ExcelManager:
    """
    Thread-safe Excel manager that serializes all Excel operations.

    This prevents COM object conflicts when multiple threads need to
    interact with Excel files simultaneously (e.g., web_download and
    sap_download running in parallel).

    Uses a single lock to ensure only one thread can perform Excel
    operations at a time, preventing race conditions and crashes.
    """
    _instance = None
    _lock = threading.Lock()
    _operation_lock = threading.RLock()  # Serializes all Excel operations

    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
        return cls._instance

    def convert_xls_to_xlsx(self, xls_path, xlsx_path, timeout=60):
        """
        Thread-safe conversion of XLS to XLSX file.

        This method handles COM initialization, file conversion, and cleanup
        all within a locked section to prevent concurrent Excel access.

        Args:
            xls_path: Path to source .xls file
            xlsx_path: Path to destination .xlsx file
            timeout: Maximum time to wait for operation lock (seconds)

        Returns:
            True if successful, False otherwise
        """
        acquired = self._operation_lock.acquire(timeout=timeout)

        if not acquired:
            logger.error(f"Timeout waiting for Excel lock ({timeout}s)")
            return False

        excel = None
        wb = None

        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()

            logger.info(f"Converting {xls_path} to {xlsx_path}...")

            # Create Excel instance
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            # Open and convert file
            wb = excel.Workbooks.Open(xls_path)
            wb.SaveAs(xlsx_path, FileFormat=51)  # 51 = xlsx format
            wb.Close(SaveChanges=False)
            wb = None

            time.sleep(0.5)  # Give Excel time to release file

            logger.info(f"✓ Conversion complete: {xlsx_path}")
            return True

        except Exception as e:
            logger.error(f"Error converting file: {e}")
            return False

        finally:
            # Clean up Excel objects
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass

            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass

            # Uninitialize COM
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

            # Release lock
            self._operation_lock.release()

    def release_excel(self, force_quit=False):
        """
        Force close any Excel instances (for cleanup in sap_download).

        Args:
            force_quit: If True, attempts to close all Excel instances
        """
        if not force_quit:
            return

        acquired = self._operation_lock.acquire(timeout=10)
        if not acquired:
            logger.warning("Could not acquire lock to force close Excel")
            return

        try:
            pythoncom.CoInitialize()

            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                excel.Quit()
                logger.info("✓ Force closed Excel instance")
            except Exception:
                logger.debug("No Excel instance found to close")

        except Exception as e:
            logger.debug(f"Excel cleanup: {e}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._operation_lock.release()


# Global singleton instance
excel_manager = ExcelManager()
