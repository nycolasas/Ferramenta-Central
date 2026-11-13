# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\extrator_saldo.py', '.'), ('C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\atualizar_bruta.py', '.'), ('C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\gerador_conf.py', '.'), ('C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\gerador_atrasos.py', '.'), ('C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\atualizar_conf.py', '.'), ('C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\atualizar_atr.py', '.'), ('C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\portable_paths.py', '.'), ('C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\conferente_atra.py', '.'), ('C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\Feriados.xlsx', '.'), ('C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\planilha_bruta.xlsx', '.'), ('C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\dados', 'dados')]
binaries = []
hiddenimports = ['selenium.webdriver.chrome.service', 'selenium.webdriver.chrome.options', 'selenium.webdriver.common.by', 'selenium.webdriver.support.ui', 'selenium.webdriver.support.expected_conditions', 'selenium.webdriver.remote.client_config', 'openpyxl.styles', 'numpy', 'pandas', 'json', 'unicodedata', 'requests']
tmp_ret = collect_all('selenium')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pandas')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('numpy')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('requests')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['C:\\Users\\Administrador\\Desktop\\pandas\\codigo\\interface_executor.py'],
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
    name='Painel_Automacao_v5',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Painel_Automacao_v5',
)
