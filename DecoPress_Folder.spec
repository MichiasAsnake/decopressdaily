# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('DecoPressLogo.jpg', '.'),
        ('DecoPressLogo.ico', '.'),
        ('PackingSlipTemplate.xlsx', '.')
    ],
    hiddenimports=[
        'win32com.client',
        'playwright.sync_api',
        'playwright._impl._api_structures',
        'playwright._impl._api_types',
        'playwright._impl._connection',
        'playwright._impl._driver',
        'playwright._impl._errors',
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

# Add the Playwright CLI to the binaries
import os
import site
from pathlib import Path
import inspect
try:
    import playwright._impl._driver
    playwright_dir = os.path.dirname(inspect.getfile(playwright._impl._driver))
    if os.path.exists(os.path.join(playwright_dir, 'driver', 'package', 'cli.js')):
        # Add all files from the driver directory
        driver_dir = os.path.join(playwright_dir, 'driver')
        for root, dirs, files in os.walk(driver_dir):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, playwright_dir)
                target_dir = os.path.dirname(os.path.join('playwright', rel_path))
                a.datas.append((os.path.join('playwright', rel_path), full_path, 'DATA'))
except (ImportError, AttributeError):
    pass

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,  # Exclude binaries from the EXE
    name='DecoPress',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,  # Set to True for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='DecoPressLogo.ico',
)

# Create the folder structure
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DecoPress_Folder',
) 