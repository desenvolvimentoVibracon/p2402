# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['principal.py'],
    pathex=[
        'C:/Users/ADV/Documents/vibraconService/projetos/p2303/p2303sw/_biblioteca/codigos',
        'C:/Users/ADV/Documents/vibraconService/projetos/p2303/p2303sw/_biblioteca/arte/logos',
        'C:/Users/ADV/Documents/vibraconService/projetos/p2303/p2303sw/_biblioteca/arte/botoes'
    ],
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
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='principal',
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
    icon='C:/Users/ADV/Documents/vibraconService/projetos/p2303/p2303sw/_biblioteca/arte/logos/iconeVibracon1.ico'
)
