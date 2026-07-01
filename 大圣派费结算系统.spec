# -*- mode: python ; coding: utf-8 -*-
"""
大圣·快递物流派费结算系统 — PyInstaller 打包配置
- onefile 模式：单文件 exe
- windowed 模式：无控制台窗口
- 自动打包图标/默认配置等资源
"""
import os
import sys

# 项目根目录（spec 文件所在目录）
project_root = os.path.dirname(os.path.abspath(SPEC))

block_cipher = None

# ========== datas：需要打包的静态资源 ==========
datas = [
    ('app', 'app'),
    (os.path.join(project_root, 'data', 'config', 'fee_rules.json'), os.path.join('data', 'config')),
    (os.path.join(project_root, 'data', 'config', 'default_settings.json'), os.path.join('data', 'config')),
    (os.path.join(project_root, 'data', 'icons'), os.path.join('data', 'icons')),
]

# ========== binaries：二进制依赖（通常留空，pyinstaller 自动收集）==========
binaries = []

# ========== hiddenimports：PyInstaller 静态分析遗漏的模块 ==========
hiddenimports = [
    # PyQt5 基础组件
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',
    'PyQt5.QtNetwork',
    'PyQt5.sip',

    # pandas & 数据处理
    'pandas',
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.timezones',
    'pandas._libs.tslibs.conversion',
    'pandas._libs.lib',
    'pandas._libs.reshape',
    'openpyxl',
    'openpyxl.reader.excel',
    'openpyxl.writer.excel',
    'xlsxwriter',

    # SQLAlchemy + SQLite
    'sqlalchemy',
    'sqlalchemy.sql.default_comparator',
    'sqlalchemy.dialects.sqlite',
    'sqlalchemy.ext.declarative',
    'sqlalchemy.orm',

    # PIL 图像处理
    'PIL',
    'PIL.Image',
    'PIL.ImageQt',

    # win32com
    'win32com',
    'win32com.client',
    'pythoncom',

    # numpy
    'numpy',
    'numpy.core._methods',
    'numpy.lib.format',

    # 其他必要依赖
    'datetime',
    'pytz',
    'ctypes',
    'ctypes.wintypes',
    'uuid',
    'hashlib',
    'hmac',
    'itertools',
    'functools',
    'collections',
    'collections.abc',
    'json',
    'csv',
    'io',
    'shutil',
    'tempfile',
    'pathlib',
    'time',
    're',
    'math',
    'decimal',
    'logging',

    # 业务模块
    'app',
    'app.core',
    'app.core.utils',
    'app.core.settlement',
    'app.core.license_manager',
    'app.models',
    'app.models.database',
    'app.models.fee_record',
    'app.models.fee_detail',
    'app.models.station',
    'app.models.courier',
    'app.models.commission_rule',
    'app.models.column_mapping',
    'app.models.customer_store',
    'app.models.user',
    'app.models.path_config',
    'app.services',
    'app.services.calculate_service',
    'app.services.customer_service',
    'app.services.export_service',
    'app.services.column_matcher',
    'app.services.rule_service',
    'app.services.excel_parser',
    'app.ui',
    'app.ui.main_window',
    'app.ui.login_window',
    'app.ui.activation_window',
    'app.ui.widgets',
    'app.ui.workers',
    'app.utils',
]

# ========== hookspath：自定义 hooks 目录（如有）==========
hookspath = [os.path.join(project_root, 'hooks')] if os.path.isdir(os.path.join(project_root, 'hooks')) else []

# ========== runtime_hooks ==========
runtime_hooks = [
    os.path.join(project_root, 'rth_fix_pkgres.py'),
]

# ========== excludes：排除不相关的大包（减小体积）==========
excludes = [
    # 不使用的 GUI 框架
    'tkinter',
    '_tkinter',
    'Tkinter',
    # 不使用的大型库
    'matplotlib',
    'scipy',
    'IPython',
    'pytest',
    'PySide2',
    'PySide6',
    'PyQt4',
    'PyQt6',
    # 不使用的 Qt 模块
    'PyQt5.QtQuick',
    'PyQt5.QtQuickWidgets',
    'PyQt5.QtQml',
    'PyQt5.QtQmlModels',
    'PyQt5.QtWebEngineCore',
    'PyQt5.QtWebEngineWidgets',
    'PyQt5.QtWebSockets',
    'PyQt5.QtDBus',
    'PyQt5.QtSvg',
    'PyQt5.QtSql',
    'PyQt5.QtXml',
    'PyQt5.QtXmlPatterns',
    'PyQt5.QtMultimedia',
    'PyQt5.QtMultimediaWidgets',
    'PyQt5.Qt3DCore',
    'PyQt5.Qt3DRender',
    'PyQt5.Qt3DInput',
    'PyQt5.Qt3DAnimation',
    'PyQt5.Qt3DExtras',
    'PyQt5.QtTest',
    'PyQt5.QtPrintSupport',
    'PyQt5.QtOpenGL',
    'PyQt5.uic',
    # 不使用的 Python 内置
    'unittest',
    'doctest',
    'lib2to3',
    # 排除 pkg_resources/setuptools
    'pkg_resources',
    'setuptools',
    'setuptools._vendor',
    'setuptools.extern',
]

a = Analysis(
    ['main.py'],
    pathex=[project_root],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=hookspath,
    runtime_hooks=runtime_hooks,
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# ========== 体积精简：过滤掉打包进来但实际未使用的 DLL ==========
exclude_keywords = [
    # Qt 未使用模块（.dll / QtQuick 资源文件）
    'Qt5Quick', 'Qt5Qml', 'Qt5WebSockets', 'Qt5WebEngine',
    'Qt5DBus', 'Qt5Svg', 'Qt5Sql', 'Qt5Xml',
    'Qt5Multimedia', 'Qt5Test', 'Qt5PrintSupport', 'Qt5OpenGL',
    'Qt53D', 'Qt5Positioning', 'Qt5Sensors', 'Qt5Serial',
    'Qt5Designer', 'Qt5Location', 'Qt5Bluetooth',
    'Qt5Help', 'Qt5Network', 'Qt5XmlPatterns',
    'qt5_qml', 'qml', 'QtQuick',
    # Qt 翻译文件（不需要中文外的语言包）
    '/translations/qt_',
    '/translations/qtbase_',
    '/translations/qtmultimedia_',
    '/translations/qt_help_',
    '/translations/assistant_',
    '/translations/designer_',
    '/translations/linguist_',
    # Tcl/Tk
    'tcl8', 'tk8', '_tkinter', 'tcl/tcl',
    # Python 标准库开发工具
    'lib2to3',
    # OpenGL 相关驱动（DLL 很大，不需要）
    'libEGL', 'libGLES', 'd3dcompiler', 'opengl32sw',
    # 不需要的 SSL 库（PyQt5 自带了一套，Python 自己也有）
    'libcrypto-1_1', 'libssl-1_1', 'libeay32', 'ssleay32',
]

def _should_exclude(path, keywords):
    """判断路径是否应该被排除（大小写不敏感）"""
    low = str(path).lower().replace('\\', '/')
    for kw in keywords:
        if kw.lower() in low:
            return True
    return False

# 过滤 datas
a.datas = [(dest, src, typ) for dest, src, typ in a.datas
           if not _should_exclude(dest, exclude_keywords)]

# 过滤 binaries（DLL 等）
a.binaries = [(dest, src, typ) for dest, src, typ in a.binaries
              if not _should_exclude(dest, exclude_keywords)]

# 过滤 zipfiles
a.zipfiles = [(dest, src, typ) for dest, src, typ in a.zipfiles
              if not _should_exclude(dest, exclude_keywords)]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 应用图标路径
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
    name='大圣派费结算系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path,
)
