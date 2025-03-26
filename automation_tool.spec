# -*- mode: python ; coding: utf-8 -*-
import platform
from PyInstaller.utils.hooks import collect_submodules

platform_name = platform.system()

our_datas = [
    ('resource', 'resource'),
    ('LICENSE', 'LICENSE'),
    ('src/task/web', 'src/task/web'),
    ('src/task/desktop', 'src/task/desktop'),
    ('src/task/arbitrary', 'src/task/arbitrary'),
    ('src', 'src')
]
our_hidden_imports = [
    'selenium.webdriver.chrome',
    'selenium.webdriver.support.expected_conditions',
    'selenium.webdriver.support.wait',
    'requests',
    'wget',
    'xlwings',
    'pdfplumber',
    'PyPDF2',
    'pyautogui',
    'pywinauto',
    'openpyxl',
    'PyQt5',
    'PyQt5.QtWidgets',
    'PyQt5.QtCore',
    'pywin32',
    'psutil',
    'fpdf',
    'certifi',
    'xvfbwrapper',
    'pikepdf',
    'pywin32',
    'pyuac',
]
if platform_name == 'Windows':
    our_hidden_imports.extend(collect_submodules('comtypes'))

our_hidden_imports.extend(collect_submodules('src'))

a = Analysis(
    ['src/gui/GUIApp.py'],
    pathex=[],
    binaries=[],
    datas=our_datas,
    hiddenimports=our_hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'numpy'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='automation_tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)