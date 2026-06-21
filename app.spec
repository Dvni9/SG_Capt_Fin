# -*- mode: python ; coding: utf-8 -*-
# Spec de PyInstaller para "Captación SG 1.0.0"
# Genera una CARPETA (onedir) con todo lo necesario para llevar en USB.
# config.py se incluye como dato externo para poder editar AGENT_ID en cada PC.

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('.env', '.'),                   # Configuración de Netelip SIP
        ('config.py', '.'),              # config editable fuera del exe
        ('shared_excel_access.py', '.'), # por si acaso, también como dato
    ],
    hiddenimports=[
        'docx',               # python-docx, se importa dinámicamente en _leer_docx
        'PyQt5.QtWebEngineWidgets',
    ],
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
    exclude_binaries=True,      # onedir: binarios van aparte en la carpeta
    name='CaptacionSG',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,              # sin ventana de consola
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='Captacion.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CaptaciónSG 1.1',   # nombre de la carpeta de salida para el USB
)
