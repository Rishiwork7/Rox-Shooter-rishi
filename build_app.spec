# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Collect all playwright data (crucial for browsers)
playwright_data = collect_data_files('playwright')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=playwright_data + [
        ('requirements.txt', '.'),
    ],
    hiddenimports=[
        'playwright.async_api',
        'pandas',
        'xlsxwriter',
        'pptx',
        'customtkinter',
        'darkdetect',
        'PIL.Image'
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
    name='GmailMailerPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False, # Set to True if you want to see terminal logs for debugging
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
