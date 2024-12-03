# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

added_files = [
    ('C:\\Users\\mt.sales\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python312\\site-packages\\numpy', 'numpy'),
    ('C:\\Users\\mt.sales\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python312\\site-packages\\pandas', 'pandas'),
    ('C:\\Users\\mt.sales\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python312\\site-packages\\matplotlib', 'matplotlib')
]

a = Analysis(
    ['final.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        'numpy',
        'pandas',
        'matplotlib',
        'openpyxl',
        'matplotlib.backends.backend_tkagg',
        'tkinter',
        'tkinter.ttk',
        'PIL',
        '_tkinter',
        'packaging.version',
        'packaging.specifiers',
        'packaging.requirements'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Sales Analysis Tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='icon.ico',
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)