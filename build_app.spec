# -*- mode: python ; coding: utf-8 -*-
import os
import sys
import glob as _glob
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect all playwright data (crucial for browsers and dependencies)
playwright_data = collect_data_files('playwright')
playwright_submodules = collect_submodules('playwright')

# Collect the Playwright driver binary (node + playwright CLI)
# This is critical - without it, `compute_driver_executable()` won't find the driver
import playwright
pw_package_dir = os.path.dirname(playwright.__file__)
pw_driver_dir = os.path.join(pw_package_dir, 'driver')
pw_driver_binaries = []
if os.path.isdir(pw_driver_dir):
    for root, dirs, files in os.walk(pw_driver_dir):
        for f in files:
            src = os.path.join(root, f)
            # Compute relative destination path within the bundle
            rel = os.path.relpath(root, pw_package_dir)
            dest = os.path.join('playwright', rel)
            pw_driver_binaries.append((src, dest))

# Additional data for other packages
pptx_data = collect_data_files('pptx')
PIL_data = collect_data_files('PIL')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=pw_driver_binaries,
    datas=playwright_data + pptx_data + PIL_data + [
        ('requirements.txt', '.'),
    ],
    hiddenimports=[
        'playwright.async_api',
        'playwright._impl._browser',
        'playwright._impl._browser_context',
        'playwright._impl._page',
        'playwright._impl._driver',
        'pandas',
        'xlsxwriter',
        'pptx',
        'customtkinter',
        'darkdetect',
        'PIL.Image',
        'PIL.PngImagePlugin',
    ] + playwright_submodules,
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
    name='GmailMailerPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True, # KEEP TRUE for first build to debug, change to False when confirmed working
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None # Add 'icon.ico' here if you have one
)

# For Mac App creation
app = BUNDLE(
    exe,
    name='GmailMailerPro.app',
    icon=None,
    bundle_identifier='com.rox.gmailmailer',
)
