# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],  # Your main script
    pathex=['/Users/user-1/DispatchGenerator'],  # Adjust to your project path
    binaries=[],
    datas=[
        ('CreateDataModel.py', '.'),  # Include necessary Python files
        ('VRP.py', '.'),
        ('format_solution.py', '.'),
        ('createDispatch.py', '.'),
        ('MasterTemplate.py', '.'),
        ('loadTemplate.py', '.'),
        ('FileDropTarget.py', '.'),
        ('MainMenu.py', '.'),
        ('createDistanceMatrix.py', '.'),
        ('format_solution.py', '.'),
        ('saveTemplate.py', '.'),
        ('styleTemplate.py', '.'),
    ],
    hiddenimports=[
        'openpyxl',
        'wx',
        'ortools',
        'sqlite3',
        'ortools.constraint_solver.routing_parameters_pb2',  # Ensure this module is included
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EasyDispatch',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to False to suppress the console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EasyDispatch',
)

app = BUNDLE(
    coll,
    name='EasyDispatch.app',
    icon='officialLogo.ico',
    bundle_identifier='com.yourcompany.easydispatch',  # Change this to your bundle identifier
    info_plist={
        'CFBundleName': 'EasyDispatch',
        'CFBundleDisplayName': 'EasyDispatch',
        'CFBundleIdentifier': 'com.yourcompany.easydispatch',  # Change this to your bundle identifier
        'CFBundleVersion': '0.1.0',
        'CFBundleShortVersionString': '0.1.0',
    },
    console=False,  # Ensure this is set to False
)
