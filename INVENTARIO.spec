# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['INVENTARIO.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
a.datas += [
            ("./formulario_inventario.ui", "formulario_inventario.ui", "DATA"),
            ("./Base_dat.db", "Base_dat.db", "DATA"),
            ("./benceno.ico", "benceno.ico", "DATA"),
            ("./sony.ico", "sony.ico", "DATA"),
            ("./navidad-ani.gif", "navidad-ani.gif", "DATA"),
            ("./INCIDENCIA.py", "INCIDENCIA.py", "DATA"),
            ("./INCIDENCIAS_SE.ui", "INCIDENCIAS_SE.ui", "DATA"),
            ("./acercade.py", "acercade.py", "DATA"),
            ("./acercade.ui", "acercade.ui", "DATA")
            ]
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='INVENTARIO',
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
    icon='base-de-datos.ico',
)
