# -*- mode: python -*-

block_cipher = None


a = Analysis(['Custum_complains.py'],
             pathex=['C:\\Users\\admin\\Documents\\share\\项目记录系统\\custum_complains'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='Custum_complains',
          debug=False,
          strip=False,
          upx=True,
          console=False , icon='110.ico')
