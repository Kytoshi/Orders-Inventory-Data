# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# collect_all grabs the full package: modules, data files, and binaries.
# send2trash uses platform-conditional imports that PyInstaller can't trace.
# NOTE: Run PyInstaller from the .venv so all packages are discoverable:
#   .venv\Scripts\pyinstaller App.spec
s2t_datas, s2t_binaries, s2t_hiddenimports = collect_all('send2trash')
# Ensure submodules are always listed even if collect_all finds nothing
s2t_hiddenimports = list(set(s2t_hiddenimports + [
    'send2trash', 'send2trash.win', 'send2trash.win.modern',
    'send2trash.win.legacy', 'send2trash.win.IFileOperationProgressSink',
    'send2trash.compat', 'send2trash.util', 'send2trash.exceptions',
]))

a = Analysis(
    ['AMS_Orders/modules/App.py'],  # Point to nested App.py
    pathex=['AMS_Orders/modules'],  # Add modules to path
    binaries=[
        ('chromedriver.exe', '.'),  # Bundle chromedriver
    ],
    datas=[
        ('AMSO Logo v2.ico', '.'),  # Bundle icon file for window/taskbar
        ('AMSO Logo v2.png', '.'),  # Bundle high-res PNG for Qt rendering
    ] + s2t_datas,
    hiddenimports=[
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.chrome',
        'selenium.webdriver.chrome.service',
        'win32com.shell',          # Required by send2trash.win.modern
        'win32com.shell.shell',
        'win32com.shell.shellcon',
        'webdriver_manager',
        'webdriver_manager.chrome',
        'PySide6',
        'PySide6.QtCore',
        'PySide6.QtWidgets',
        'PySide6.QtGui',
        'win32com.client',
        'win32timezone',           # Required by win32com for COM DATE/datetime handling
        'pythoncom',               # Required for COM thread initialization
        'pywintypes',              # Required by win32com for COM type conversions
        'threading',
        'multiprocessing',
        'helpers',  # Your custom modules
        'file_utils',
        'logger',
        'web_download',
        'sap_download',
        'excel_report',            # Migrated Excel report engine module
    ] + s2t_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter', '_tkinter',       # Not used — PySide6 is the GUI
        'unittest',                   # Test framework not needed at runtime
        'pytest',
        'PySide6.QtNetwork',         # Unused Qt modules
        'PySide6.QtQml',
        'PySide6.QtQuick',
        'PySide6.QtSvg',
        'PySide6.QtMultimedia',
        'PySide6.QtWebEngine',
        'PySide6.QtWebEngineWidgets',
        'PySide6.Qt3DCore',
        'PySide6.Qt3DRender',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AMSOrderDownloadManager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='AMSO Logo v2.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AMSOrderDownloadManager',
)