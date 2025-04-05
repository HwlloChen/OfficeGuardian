from PyInstaller.building.api import PYZ, EXE
from PyInstaller.building.build_main import Analysis
import sys
import os

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[os.path.abspath(SPECPATH)],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pycaw.pycaw',
        'comtypes',
        'win32com.client',
        'numpy',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,           # 添加二进制文件
    a.zipfiles,          # 添加压缩文件
    a.datas,             # 添加数据文件
    [],
    exclude_binaries=False,  # 改为False以包含所有文件
    name='OfficeGuardian',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    runtime_tmpdir=None,    # 添加这行以确保在临时目录中运行
    # icon='icon.ico',
)
