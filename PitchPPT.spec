# -*- mode: python ; coding: utf-8 -*-

import sys
import os

# 项目根目录
project_root = os.path.abspath('.')

block_cipher = None

# 分析主程序
a = Analysis(
    ['src\\main.py'],
    pathex=[project_root],
    binaries=[],
    datas=[
        # 包含资源文件
        ('resources', 'resources'),
        # 包含配置文件
        ('config.json', '.'),
        # 包含文档
        ('README.md', '.'),
        ('README.zh-CN.md', '.'),
        ('LICENSE', '.'),
    ],
    hiddenimports=[
        'win32com.client',
        'win32com.gen_py',
        'pythoncom',
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
        'PyQt5.sip',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'src.core',
        'src.core.converter',
        'src.core.win32_converter',
        'src.core.smart_optimizer_v4',
        'src.core.smart_optimizer_v5',
        'src.core.smart_optimizer_v6',
        'src.core.progress_tracker',
        'src.ui',
        'src.ui.main_window',
        'src.ui.smart_config_widget',
        'src.ui.batch_conversion_worker',
        'src.ui.smart_optimization_worker_v4',
        'src.utils',
        'src.utils.logger',
        'src.utils.config_manager',
        'src.utils.history_manager',
        'src.utils.error_handler',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'tkinter',
        'unittest',
        'pytest',
        'pydoc',
        'email',
        'http',
        'xmlrpc',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 去除重复项
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 配置可执行文件
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PitchPPT',
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
    # 图标文件
    icon='resources\\LOGO_256x256.ico',
    # 版本信息文件
    version='version.txt',
)
