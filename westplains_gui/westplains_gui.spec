# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['westplains_gui.py'],
             pathex=[],
             binaries=[],
             datas=[('tabula-1.0.5-jar-with-dependencies.jar', './tabula/')],
             hiddenimports=['tkcalendar', 'babel.numbers', 'xlwings', 'tabula'],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts, 
          [],
          exclude_binaries=True,
          name='westplains_gui',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None , icon='biourjaLogo.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas, 
               strip=False,
               upx=True,
               upx_exclude=[],
               name='westplains_gui')
