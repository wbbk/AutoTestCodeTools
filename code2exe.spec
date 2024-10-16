# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules, collect_data_files
import os
import sys
sys.setrecursionlimit(sys.getrecursionlimit() * 5)

block_cipher = None

# 获取脚本所在的目录
path = os.path.dirname(sys.argv[0])

# 收集需要的子模块
hiddenimports = (
    collect_submodules('openpyxl') +
    collect_submodules('poco') +
    collect_submodules('airtest') +
    ['tkinter']
)

# 收集数据文件
datas = (
    collect_data_files('openpyxl') +
    collect_data_files('poco') +
    collect_data_files('airtest') +
    [
        (os.path.join(path, 'tpl1727690063231.png'), 'tpl1727690063231.png'),
        (os.path.join(path, 'tpl1727667630169.png'), 'tpl1727667630169.png'),
        (os.path.join(path, 'tpl1724235301671.png'), 'tpl1724235301671.png')
    ]  # 添加模板图片
)

a = Analysis(
    ['code2exe.py'],
    pathex=[os.path.join(path)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[os.path.join(path, 'hooks')],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False)

pyz = PYZ(a.pure, a.zipped_data,cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    name='code2exe',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # 启用UPX压缩
    upx_exclude=[],  # 排除不需要压缩的文件
    runtime_tmpdir=None,
    console=True,  # 根据你的应用程序是否为控制台应用程序来设置
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)