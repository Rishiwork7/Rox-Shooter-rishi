# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect all playwright data (needed for the async_api to work)
playwright_data = collect_data_files('playwright')
playwright_submodules = collect_submodules('playwright')

# Additional data for other packages
pptx_data = collect_data_files('pptx')
PIL_data = collect_data_files('PIL')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
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
        'supabase',
        'requests',
        'postgrest',
        'gotrue',
        'storage3',
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
