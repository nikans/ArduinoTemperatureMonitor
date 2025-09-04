# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['temperature_monitor.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('lang', 'lang'),
        ('config.ini', '.'),
        ('localization.py', '.'),
    ],
    hiddenimports=[
        'serial',
        'serial.tools.list_ports',
        'matplotlib.backends.backend_tkagg',
        'matplotlib.figure',
        'numpy',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'csv',
        'os',
        'threading',
        'time',
        'datetime',
        'yaml',
        'configparser',
        'locale',
        'localization'
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

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='TemperatureMonitor',
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
    # icon='icon.ico'  # We'll create this
)
