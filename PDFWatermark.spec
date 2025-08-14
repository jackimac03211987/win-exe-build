# PDFWatermark.spec
# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from PyInstaller.utils.hooks import collect_all, collect_submodules, collect_data_files

# 收集所有 pandas 和 openpyxl 的依赖
datas = []
binaries = []
hiddenimports = []

# 收集 pandas 的所有内容
tmp_ret = collect_all('pandas')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# 收集 openpyxl 的所有内容
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# 添加额外的隐藏导入
hiddenimports += [
    'openpyxl',
    'openpyxl.cell._writer',
    'openpyxl.styles.stylesheet',
    'pandas.io.excel._openpyxl',
    'pandas.io.formats.excel',
    'pdf2image',
    'pdf2image.pdf2image',
]

# 添加 Poppler 二进制文件
if sys.platform == 'win32':
    # Windows 平台
    poppler_path = 'poppler'
    if os.path.exists(poppler_path):
        # 添加所有 Poppler 二进制文件
        for file in os.listdir(poppler_path):
            if file.endswith('.exe') or file.endswith('.dll'):
                binaries.append((os.path.join(poppler_path, file), '.'))

a = Analysis(
    ['src/app_main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PDFWatermark',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # 不显示控制台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PDFWatermark',
)
