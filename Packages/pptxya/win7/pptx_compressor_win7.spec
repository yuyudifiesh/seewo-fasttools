# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

hidden_imports_list = [
    'PIL',
    'PIL.Image',
    'PIL.ImageFile',
    'zipfile',
    'io',
    'tkinter',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.ttk',
    'os',
    'sys',
    'threading',
]

a = Analysis(['y_win7.py'],
             pathex=['路径'],
             binaries=[],
             datas=[],
             hiddenimports=hidden_imports_list,
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='PPT压缩工具_Win7专用版',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True,  # 保持控制台开启
          icon=None,
          version=1,
          uac_admin=False)
