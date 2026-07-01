# -*- mode: python ; coding: utf-8 -*-
"""
激活码生成器 — PyInstaller 打包配置
轻量级：仅包含 PyQt5 + license_manager，不需要 pandas/sqlalchemy
"""
import os

project_root = os.path.dirname(os.path.abspath(SPEC))

block_cipher = None

# 只打包必要的资源
datas = [
    (os.path.join(project_root, 'data', 'icons'), os.path.join('data', 'icons')),
]

hiddenimports = [
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',
    'PyQt5.sip',
    'app',
    'app.core',
    'app.core.license_manager',
    'app.models',
    'app.models.path_config',
]

excludes = [
    'pandas',
    'openpyxl',
    'xlsxwriter',
    'python_calamine',
    'calamine',
    'sqlalchemy',
    'matplotlib',
    'scipy',
    'IPython',
    'notebook',
    'jupyter',
    'pytest',
    'tkinter',
]

a = Analysis(
    ['app/utils/keygen_gui.py'],
    pathex=[project_root],
    datas=datas,
    hiddenimports=hiddenimports,
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

icon_path = os.path.join(project_root, 'data', 'icons', 'dasheng.ico')
if not os.path.exists(icon_path):
    icon_path = None

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='激活码生成器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path,
)
