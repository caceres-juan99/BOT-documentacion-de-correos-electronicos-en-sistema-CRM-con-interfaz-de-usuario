# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Interfaz.py'],
    pathex=[],
    binaries=[],
    datas=[('C:/Users/cacer/OneDrive - unimilitar.edu.co/NECTIM COLOMBIA S.A.S/BPO/BOT BOTONES DE PROYECTO/Interfaz de usuario/Iconos/Fondo de interfaz.png', 'Icono')],
    hiddenimports=[],
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
    a.binaries,
    a.datas,
    [],
    name='Interfaz',
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
    icon=['C:\\Users\\cacer\\OneDrive - unimilitar.edu.co\\NECTIM COLOMBIA S.A.S\\BPO\\BOT BOTONES DE PROYECTO\\Interfaz de usuario\\Iconos\\Icono de ventana.ico'],
)
