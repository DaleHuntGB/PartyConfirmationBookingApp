# booking_confirmation_app.spec
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['BookingConfirmationApp.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('BookingData.json', '.'),
        ('PartyBookingConfirmationTemplate.docx', '.'),
        ('Media/FL_Logo.ico', 'Media')
    ],
    hiddenimports=['tkinter', 'tkcalendar', 'ttkthemes', 'docx', 'babel.numbers'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='BookingConfirmationApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='Media/FL_Logo.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='BookingConfirmationApp'
)
