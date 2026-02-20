"""
AOMOSO Download Manager - PySide6 Desktop App
High Contrast Modern UI - Variant 2
"""

import sys
import threading
import logging
import io
import time
import os
import ctypes
from contextlib import redirect_stdout, redirect_stderr
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QGroupBox, QFormLayout,
    QProgressBar, QMessageBox, QFrame
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QIcon

# Set Windows AppUserModelID so the taskbar shows our icon instead of Python's default
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("SMC.AOMOSODownloadManager")

# Lazy-loaded at first use to speed up app startup
web_download = None
sap_download = None
excel_report = None


def _ensure_imports():
    """Import heavy modules on first use (selenium, win32com, etc.)."""
    global web_download, sap_download, excel_report
    if web_download is None:
        import web_download as _wd
        web_download = _wd
    if sap_download is None:
        import sap_download as _sd
        sap_download = _sd
    if excel_report is None:
        import excel_report as _er
        excel_report = _er


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller (onedir)."""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# High Contrast Theme - Nord inspired with better readability
class Theme:
    # Darker, richer backgrounds
    BACKGROUND = "#0d1117"        # GitHub dark bg
    SURFACE = "#161b22"           # Elevated surface
    SURFACE_LIGHT = "#21262d"     # Input backgrounds

    # Vibrant accent colors
    PRIMARY = "#B6CEB4"           # Bright blue
    PRIMARY_DARK = "#96A78D"      # Darker blue
    SUCCESS = "#3fb950"           # Vibrant green
    WARNING = "#f2cc60"           # Bright yellow
    ERROR = "#f85149"             # Bright red

    # High contrast text
    TEXT = "#f0f6fc"              # Almost white
    TEXT_DIM = "#8b949e"          # Medium gray
    TEXT_MUTED = "#6e7681"        # Muted gray

    BORDER = "#30363d"            # Subtle borders
    ACCENT = "#bc8cff"            # Purple accent


class LogFileMonitor(threading.Thread):
    """Monitor a log file and emit new lines to Qt Signal"""
    def __init__(self, log_file_path, signal):
        super().__init__(daemon=True)
        self.log_file_path = log_file_path
        self.signal = signal
        self.running = True
        self.last_position = 0

    def run(self):
        """Monitor log file for new content"""
        while self.running and not os.path.exists(self.log_file_path):
            time.sleep(0.1)

        if not self.running:
            return

        try:
            with open(self.log_file_path, 'r', encoding='utf-8', errors='ignore') as f:
                f.seek(0, 2)
                self.last_position = f.tell()

                while self.running:
                    line = f.readline()
                    if line:
                        line = line.strip()
                        if line:
                            self.signal.emit(line)
                        self.last_position = f.tell()
                    else:
                        time.sleep(0.1)

        except Exception as e:
            self.signal.emit(f"Log monitor error: {str(e)}")

    def stop(self):
        """Stop monitoring"""
        self.running = False


class StreamCapture(io.StringIO):
    """Capture stdout/stderr and emit to Qt Signal"""
    def __init__(self, signal):
        super().__init__()
        self.signal = signal

    def write(self, text):
        if text and text.strip():
            self.signal.emit(text.strip())
        return len(text)


class QtLogHandler(logging.Handler):
    """Custom logging handler that emits logs to Qt Signal"""
    def __init__(self, signal):
        super().__init__()
        self.signal = signal

    def emit(self, record):
        try:
            msg = self.format(record)
            self.signal.emit(msg)
        except Exception:
            pass


class WorkerThread(QThread):
    """Background thread to run scripts without freezing UI"""
    finished = Signal(bool, str)
    progress = Signal(str)

    def __init__(self, script_type, username, password, sap_username='', sap_password='', log_file_path=None):
        super().__init__()
        self.script_type = script_type
        self.username = username
        self.password = password
        self.sap_username = sap_username
        self.sap_password = sap_password
        self.log_file_path = log_file_path
        self.log_monitor = None

    def run(self):
        """Run the script in background"""
        _ensure_imports()

        if self.log_file_path:
            self.log_monitor = LogFileMonitor(self.log_file_path, self.progress)
            self.log_monitor.start()

        log_handler = QtLogHandler(self.progress)
        log_handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(message)s')
        log_handler.setFormatter(formatter)

        root_logger = logging.getLogger()
        root_logger.addHandler(log_handler)
        root_logger.setLevel(logging.INFO)

        stdout_capture = StreamCapture(self.progress)
        stderr_capture = StreamCapture(self.progress)

        try:
            with redirect_stdout(stdout_capture), redirect_stderr(stderr_capture):
                if self.script_type == 'both':
                    self.progress.emit("Starting both scripts in parallel...")

                    results = {'website': None, 'sap': None}
                    errors = {'website': None, 'sap': None}

                    def run_website():
                        try:
                            self.progress.emit("→ Website download started...")
                            web_download.main(self.username, self.password)
                            results['website'] = type('obj', (object,), {'returncode': 0})()
                            self.progress.emit("✓ Website download completed!")
                        except Exception as script_error:
                            results['website'] = type('obj', (object,), {'returncode': 1})()
                            errors['website'] = str(script_error)
                            self.progress.emit(f"✗ Website download failed: {script_error}")

                    def run_sap():
                        try:
                            time.sleep(2)
                            self.progress.emit("→ SAP extraction started...")
                            sap_download.main(self.sap_username, self.sap_password)
                            results['sap'] = type('obj', (object,), {'returncode': 0})()
                            self.progress.emit("✓ SAP extraction completed!")
                        except Exception as script_error:
                            results['sap'] = type('obj', (object,), {'returncode': 1})()
                            errors['sap'] = str(script_error)
                            self.progress.emit(f"✗ SAP extraction failed: {script_error}")

                    thread1 = threading.Thread(target=run_website)
                    thread2 = threading.Thread(target=run_sap)

                    thread1.start()
                    thread2.start()

                    thread1.join()
                    thread2.join()

                    website_success = results['website'] and results['website'].returncode == 0
                    sap_success = results['sap'] and results['sap'].returncode == 0

                    if website_success and sap_success:
                        self.finished.emit(True, "Both scripts completed successfully!")
                    elif website_success or sap_success:
                        msg = "Partial success: "
                        if website_success:
                            msg += "Website ✓, SAP ✗"
                        else:
                            msg += "Website ✗, SAP ✓"
                        self.finished.emit(False, msg)
                    else:
                        self.finished.emit(False, "Both scripts failed!")

                elif self.script_type == 'website':
                    self.progress.emit("Starting website script...")
                    try:
                        web_download.main(self.username, self.password)
                        self.finished.emit(True, "PDBS Files Downloaded Successfully!")
                    except Exception as e:
                        self.progress.emit(f"Error: {str(e)}")
                        self.finished.emit(False, f"Error: {str(e)}")

                elif self.script_type == 'sap':
                    self.progress.emit("Starting SAP script...")
                    try:
                        sap_download.main(self.sap_username, self.sap_password)
                        self.finished.emit(True, "SAP Files Downloaded Successfully!")
                    except Exception as e:
                        self.progress.emit(f"Error: {str(e)}")
                        self.finished.emit(False, f"Error: {str(e)}")

                elif self.script_type == 'excel_report':
                    self.progress.emit("Starting Excel report engine...")
                    try:
                        excel_report.main(
                            progress_callback=lambda pct, stage: self.progress.emit(f"[{pct}%] {stage}")
                        )
                        self.finished.emit(True, "Excel Report completed successfully!")
                    except Exception as e:
                        self.progress.emit(f"Error: {str(e)}")
                        self.finished.emit(False, f"Error: {str(e)}")

                elif self.script_type == 'all':
                    # Phase 1: downloads in parallel
                    self.progress.emit("Starting downloads (PDBS + SAP)...")

                    results = {'website': None, 'sap': None}
                    errors = {'website': None, 'sap': None}

                    def run_website():
                        try:
                            self.progress.emit("→ Website download started...")
                            web_download.main(self.username, self.password)
                            results['website'] = type('obj', (object,), {'returncode': 0})()
                            self.progress.emit("✓ Website download completed!")
                        except Exception as script_error:
                            results['website'] = type('obj', (object,), {'returncode': 1})()
                            errors['website'] = str(script_error)
                            self.progress.emit(f"✗ Website download failed: {script_error}")

                    def run_sap():
                        try:
                            time.sleep(2)
                            self.progress.emit("→ SAP extraction started...")
                            sap_download.main(self.sap_username, self.sap_password)
                            results['sap'] = type('obj', (object,), {'returncode': 0})()
                            self.progress.emit("✓ SAP extraction completed!")
                        except Exception as script_error:
                            results['sap'] = type('obj', (object,), {'returncode': 1})()
                            errors['sap'] = str(script_error)
                            self.progress.emit(f"✗ SAP extraction failed: {script_error}")

                    thread1 = threading.Thread(target=run_website)
                    thread2 = threading.Thread(target=run_sap)
                    thread1.start()
                    thread2.start()
                    thread1.join()
                    thread2.join()

                    website_ok = results['website'] and results['website'].returncode == 0
                    sap_ok = results['sap'] and results['sap'].returncode == 0

                    if not (website_ok and sap_ok):
                        msg = "Downloads failed — "
                        if not website_ok and not sap_ok:
                            msg += "both PDBS and SAP failed."
                        elif not website_ok:
                            msg += "PDBS failed, SAP succeeded."
                        else:
                            msg += "PDBS succeeded, SAP failed."
                        self.finished.emit(False, msg + " Excel report skipped.")
                        return

                    # Phase 2: excel report (sequential, needs download outputs)
                    self.progress.emit("→ Starting Excel report engine...")
                    try:
                        excel_report.main(
                            progress_callback=lambda pct, stage: self.progress.emit(f"[{pct}%] {stage}")
                        )
                        self.finished.emit(True, "All tasks completed successfully!")
                    except Exception as e:
                        self.progress.emit(f"✗ Excel report failed: {e}")
                        self.finished.emit(False, f"Downloads succeeded but Excel report failed: {e}")

        except Exception as e:
            self.finished.emit(False, f"Error: {str(e)}")
        finally:
            root_logger.removeHandler(log_handler)
            if self.log_monitor:
                self.log_monitor.stop()
                self.log_monitor.join(timeout=2)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AOMOSO Download Manager")
        self.setFixedSize(1100, 700)

        # Set window icon (works for both dev and PyInstaller)
        # In dev mode: look in parent directory (root folder)
        # In PyInstaller: look in _MEIPASS (bundled resources)
        icon_path = get_resource_path("AMSO Logo v2.ico")
        if not os.path.exists(icon_path):
            # Fallback: try root directory relative to this file
            icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "AMSO Logo v2.ico")

        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.log_file_path = os.path.join(os.getcwd(), "logs", "ams_orders.txt")

        # Main widget with horizontal split layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        root_layout = QHBoxLayout(main_widget)
        root_layout.setContentsMargins(28, 28, 28, 28)
        root_layout.setSpacing(24)

        # ── Left panel: controls ──
        left_panel = QVBoxLayout()
        left_panel.setSpacing(20)

        # Header
        header_container = QVBoxLayout()
        header_container.setSpacing(6)

        header = QLabel("AOMOSO Download Manager")
        header.setObjectName("header")
        header_container.addWidget(header)

        subtitle = QLabel("Automated data extraction for AOMOSO reporting")
        subtitle.setObjectName("subtitle")
        header_container.addWidget(subtitle)

        left_panel.addLayout(header_container)

        # Credentials Section
        creds_group = self.create_credentials_section()
        left_panel.addWidget(creds_group)

        # Action Buttons
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(12)

        self.website_btn = QPushButton("PDBS Download")
        self.website_btn.setObjectName("secondaryButton")
        self.website_btn.setFixedHeight(44)
        self.website_btn.clicked.connect(self.run_website_script)

        self.sap_btn = QPushButton("SAP Download")
        self.sap_btn.setObjectName("secondaryButton")
        self.sap_btn.setFixedHeight(44)
        self.sap_btn.clicked.connect(self.run_sap_script)

        self.run_both_btn = QPushButton("PDBS + SAP")
        self.run_both_btn.setObjectName("primaryButton")
        self.run_both_btn.setFixedHeight(44)
        self.run_both_btn.clicked.connect(self.run_both_scripts)

        buttons_layout.addWidget(self.website_btn)
        buttons_layout.addWidget(self.sap_btn)
        buttons_layout.addWidget(self.run_both_btn)

        left_panel.addLayout(buttons_layout)

        # Second row: Excel Report + Run All
        buttons_layout2 = QHBoxLayout()
        buttons_layout2.setSpacing(12)

        self.excel_report_btn = QPushButton("Excel Report")
        self.excel_report_btn.setObjectName("secondaryButton")
        self.excel_report_btn.setFixedHeight(44)
        self.excel_report_btn.clicked.connect(self.run_excel_report)

        self.run_all_btn = QPushButton("Run Full Download + Excel")
        self.run_all_btn.setObjectName("primaryButton")
        self.run_all_btn.setFixedHeight(44)
        self.run_all_btn.clicked.connect(self.run_all)

        buttons_layout2.addWidget(self.excel_report_btn)
        buttons_layout2.addWidget(self.run_all_btn)

        left_panel.addLayout(buttons_layout2)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(3)
        left_panel.addWidget(self.progress_bar)

        left_panel.addStretch()

        # Wrap left panel in a widget with fixed width
        left_widget = QWidget()
        left_widget.setLayout(left_panel)
        left_widget.setFixedWidth(480)
        root_layout.addWidget(left_widget)

        # ── Right panel: activity log ──
        right_panel = QVBoxLayout()
        right_panel.setSpacing(10)

        log_label = QLabel("Activity Log")
        log_label.setObjectName("sectionLabel")
        right_panel.addWidget(log_label)

        self.log_console = QTextEdit()
        self.log_console.setReadOnly(True)
        self.log_console.setObjectName("logConsole")
        right_panel.addWidget(self.log_console, 1)

        root_layout.addLayout(right_panel, 1)

        # Apply theme
        self.apply_theme()

    def create_credentials_section(self):
        """Create credentials input section"""
        group = QFrame()
        group.setObjectName("credentialsFrame")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(20)

        # PDBS Section
        pdbs_label = QLabel("PDBS CREDENTIALS")
        pdbs_label.setObjectName("sectionLabel")
        layout.addWidget(pdbs_label)

        pdbs_layout = QVBoxLayout()
        pdbs_layout.setSpacing(10)

        self.website_username = QLineEdit()
        self.website_username.setPlaceholderText("Username")
        self.website_username.setFixedHeight(42)
        self.website_password = QLineEdit()
        self.website_password.setPlaceholderText("Password")
        self.website_password.setFixedHeight(42)
        self.website_password.setEchoMode(QLineEdit.EchoMode.Password)

        pdbs_layout.addWidget(self.website_username)
        pdbs_layout.addSpacing(10)
        pdbs_layout.addWidget(self.website_password)
        layout.addLayout(pdbs_layout)

        # Divider
        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setObjectName("divider")
        layout.addWidget(divider)

        # SAP Section
        sap_label = QLabel("SAP CREDENTIALS")
        sap_label.setObjectName("sectionLabel")
        layout.addWidget(sap_label)

        sap_layout = QVBoxLayout()
        sap_layout.setSpacing(10)

        self.sap_username = QLineEdit()
        self.sap_username.setPlaceholderText("Username")
        self.sap_username.setFixedHeight(42)
        self.sap_password = QLineEdit()
        self.sap_password.setPlaceholderText("Password")
        self.sap_password.setFixedHeight(42)
        self.sap_password.setEchoMode(QLineEdit.EchoMode.Password)

        sap_layout.addWidget(self.sap_username)
        sap_layout.addSpacing(10)
        sap_layout.addWidget(self.sap_password)
        layout.addLayout(sap_layout)

        return group

    def apply_theme(self):
        """Apply high contrast modern theme"""
        self.setStyleSheet(f"""
            * {{
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            }}

            QMainWindow {{
                background-color: {Theme.BACKGROUND};
            }}

            #header {{
                color: {Theme.TEXT};
                font-size: 28px;
                font-weight: 700;
                letter-spacing: -0.5px;
                margin-bottom: 0px;
            }}

            #subtitle {{
                color: {Theme.TEXT_DIM};
                font-size: 14px;
                font-weight: 400;
                margin-bottom: 4px;
            }}

            #sectionLabel {{
                color: {Theme.TEXT_DIM};
                font-size: 11px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: 1px;
                margin-top: 2px;
                margin-bottom: 4px;
            }}

            #credentialsFrame {{
                background-color: {Theme.SURFACE};
                border: 1px solid {Theme.BORDER};
                border-radius: 10px;
            }}

            #divider {{
                background-color: {Theme.BORDER};
                max-height: 1px;
                border: none;
                margin: 6px 0px;
            }}

            QLineEdit {{
                background-color: {Theme.SURFACE_LIGHT};
                border: 1.5px solid {Theme.BORDER};
                border-radius: 7px;
                padding: 0px 14px;
                color: {Theme.TEXT};
                font-size: 14px;
                font-weight: 500;
            }}

            QLineEdit:focus {{
                border: 1.5px solid {Theme.PRIMARY};
                background-color: {Theme.BACKGROUND};
            }}

            QLineEdit::placeholder {{
                color: {Theme.TEXT_MUTED};
            }}

            QPushButton {{
                border: none;
                border-radius: 7px;
                padding: 0px 20px;
                font-size: 14px;
                font-weight: 600;
                letter-spacing: 0.2px;
            }}

            #primaryButton {{
                background-color: {Theme.PRIMARY};
                color: #1a1a1a;
            }}

            #primaryButton:hover {{
                background-color: {Theme.PRIMARY_DARK};
            }}

            #primaryButton:pressed {{
                background-color: {Theme.PRIMARY_DARK};
            }}

            #primaryButton:disabled {{
                background-color: {Theme.SURFACE_LIGHT};
                color: {Theme.TEXT_MUTED};
            }}

            #secondaryButton {{
                background-color: {Theme.SURFACE};
                border: 1.5px solid {Theme.BORDER};
                color: {Theme.TEXT};
            }}

            #secondaryButton:hover {{
                background-color: {Theme.SURFACE_LIGHT};
                border-color: {Theme.TEXT_DIM};
            }}

            #secondaryButton:pressed {{
                background-color: {Theme.SURFACE_LIGHT};
            }}

            #secondaryButton:disabled {{
                background-color: {Theme.SURFACE};
                border-color: {Theme.BORDER};
                color: {Theme.TEXT_MUTED};
            }}

            QProgressBar {{
                background-color: {Theme.SURFACE_LIGHT};
                border: none;
                border-radius: 1.5px;
            }}

            QProgressBar::chunk {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {Theme.PRIMARY}, stop:1 {Theme.ACCENT});
                border-radius: 1.5px;
            }}

            #logConsole {{
                background-color: {Theme.SURFACE};
                border: 1.5px solid {Theme.BORDER};
                border-radius: 10px;
                color: {Theme.TEXT};
                font-family: 'SF Mono', 'Monaco', 'Consolas', 'Courier New', monospace;
                font-size: 12px;
                line-height: 1.6;
                padding: 14px;
                selection-background-color: {Theme.PRIMARY};
            }}
        """)

    def add_log(self, message):
        """Add message to log console with better formatting"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")

        # Color code based on message type
        if "✓" in message or "SUCCESS" in message.upper() or "completed" in message.lower():
            color = Theme.SUCCESS
            icon = "✓"
        elif "✗" in message or "ERROR" in message.upper() or "FAILED" in message.upper():
            color = Theme.ERROR
            icon = "✗"
        elif "→" in message or "WARNING" in message.upper() or "started" in message.lower():
            color = Theme.WARNING
            icon = "→"
        else:
            color = Theme.TEXT
            icon = "•"

        # Better formatting with icon
        formatted_msg = f'''
            <div style="margin: 2px 0;">
                <span style="color: {Theme.TEXT_MUTED}; font-size: 11px;">{timestamp}</span>
                <span style="color: {color}; font-weight: 600; margin: 0 6px;">{icon}</span>
                <span style="color: {color};">{message}</span>
            </div>
        '''
        self.log_console.append(formatted_msg.strip())

    def disable_all_buttons(self):
        """Disable all run buttons"""
        self.website_btn.setEnabled(False)
        self.sap_btn.setEnabled(False)
        self.run_both_btn.setEnabled(False)
        self.excel_report_btn.setEnabled(False)
        self.run_all_btn.setEnabled(False)

    def enable_all_buttons(self):
        """Enable all run buttons"""
        self.website_btn.setEnabled(True)
        self.sap_btn.setEnabled(True)
        self.run_both_btn.setEnabled(True)
        self.excel_report_btn.setEnabled(True)
        self.run_all_btn.setEnabled(True)

    def run_website_script(self):
        """Run website download script"""
        username = self.website_username.text().strip()
        password = self.website_password.text().strip()

        if not username or not password:
            self.show_error("Missing Information", "Please enter both username and password")
            return

        self.add_log("Starting website download...")
        self.disable_all_buttons()
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

        self.worker = WorkerThread('website', username, password, log_file_path=self.log_file_path)
        self.worker.progress.connect(self.add_log)
        self.worker.finished.connect(self.on_website_finished)
        self.worker.start()

    def run_sap_script(self):
        """Run SAP extraction script"""
        username = self.sap_username.text().strip()
        password = self.sap_password.text().strip()

        if not username or not password:
            self.show_error("Missing Information", "Please enter both username and password")
            return

        self.add_log("Starting SAP data extraction...")
        self.disable_all_buttons()
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

        self.worker = WorkerThread('sap', '', '', sap_username=username, sap_password=password, log_file_path=self.log_file_path)
        self.worker.progress.connect(self.add_log)
        self.worker.finished.connect(self.on_sap_finished)
        self.worker.start()

    def run_both_scripts(self):
        """Run both website and SAP scripts in parallel"""
        web_username = self.website_username.text().strip()
        web_password = self.website_password.text().strip()
        sap_username = self.sap_username.text().strip()
        sap_password = self.sap_password.text().strip()

        if not web_username or not web_password:
            self.show_error("Missing Information", "Please enter website username and password")
            return

        if not sap_username or not sap_password:
            self.show_error("Missing Information", "Please enter SAP username and password")
            return

        self.add_log("Starting both scripts in parallel...")
        self.disable_all_buttons()
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

        self.worker = WorkerThread('both', web_username, web_password, sap_username, sap_password, log_file_path=self.log_file_path)
        self.worker.progress.connect(self.add_log)
        self.worker.finished.connect(self.on_both_finished)
        self.worker.start()

    def run_excel_report(self):
        """Run Excel report engine (no credentials needed)"""
        self.add_log("Starting Excel report engine...")
        self.disable_all_buttons()
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

        self.worker = WorkerThread('excel_report', '', '', log_file_path=self.log_file_path)
        self.worker.progress.connect(self.add_log)
        self.worker.finished.connect(self.on_generic_finished)
        self.worker.start()

    def run_all(self):
        """Run downloads (parallel) then Excel report (sequential)"""
        web_username = self.website_username.text().strip()
        web_password = self.website_password.text().strip()
        sap_username = self.sap_username.text().strip()
        sap_password = self.sap_password.text().strip()

        if not web_username or not web_password:
            self.show_error("Missing Information", "Please enter website username and password")
            return

        if not sap_username or not sap_password:
            self.show_error("Missing Information", "Please enter SAP username and password")
            return

        self.add_log("Starting full pipeline (Downloads + Excel Report)...")
        self.disable_all_buttons()
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

        self.worker = WorkerThread('all', web_username, web_password, sap_username, sap_password, log_file_path=self.log_file_path)
        self.worker.progress.connect(self.add_log)
        self.worker.finished.connect(self.on_generic_finished)
        self.worker.start()

    def on_generic_finished(self, success, message):
        """Handle completion for excel_report and all modes"""
        self.enable_all_buttons()
        self.progress_bar.setVisible(False)
        self.add_log(message)

        if success:
            self.show_success("Success", message)
        else:
            self.show_error("Error", message)

    def on_website_finished(self, success, message):
        """Handle website script completion"""
        self.enable_all_buttons()
        self.progress_bar.setVisible(False)
        self.add_log(message)

        if success:
            self.show_success("Success", message)
        else:
            self.show_error("Error", message)

    def on_sap_finished(self, success, message):
        """Handle SAP script completion"""
        self.enable_all_buttons()
        self.progress_bar.setVisible(False)
        self.add_log(message)

        if success:
            self.show_success("Success", message)
        else:
            self.show_error("Error", message)

    def on_both_finished(self, success, message):
        """Handle both scripts completion"""
        self.enable_all_buttons()
        self.progress_bar.setVisible(False)
        self.add_log(message)

        if success:
            self.show_success("Success", message)
        else:
            self.show_error("Error", message)

    def show_success(self, title, message):
        """Show success message box"""
        msg = QMessageBox(self)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setStyleSheet(f"""
            QMessageBox {{
                background-color: {Theme.SURFACE};
            }}
            QMessageBox QLabel {{
                color: {Theme.TEXT};
                font-size: 13px;
            }}
            QMessageBox QPushButton {{
                background-color: {Theme.SUCCESS};
                color: #ffffff;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: 600;
                min-width: 80px;
            }}
        """)
        msg.exec()

    def show_error(self, title, message):
        """Show error message box"""
        msg = QMessageBox(self)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.setIcon(QMessageBox.Icon.Critical)
        msg.setStyleSheet(f"""
            QMessageBox {{
                background-color: {Theme.SURFACE};
            }}
            QMessageBox QLabel {{
                color: {Theme.TEXT};
                font-size: 13px;
            }}
            QMessageBox QPushButton {{
                background-color: {Theme.ERROR};
                color: #ffffff;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: 600;
                min-width: 80px;
            }}
        """)
        msg.exec()


def main():
    app = QApplication(sys.argv)

    # Set application-level icon (needed for taskbar on Windows)
    icon_path = get_resource_path("AMSO Logo v2.png")
    if not os.path.exists(icon_path):
        icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "AMSO Logo v2.png")
    if not os.path.exists(icon_path):
        # Fall back to .ico if PNG not found
        icon_path = get_resource_path("AMSO Logo v2.ico")
        if not os.path.exists(icon_path):
            icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "AMSO Logo v2.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
