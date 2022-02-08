# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.building.toc_conversion import Tree
from PyInstaller.building.api import EXE, PYZ
from PyInstaller.building.build_main import Analysis, COLLECT
from PyInstaller.building.osx import BUNDLE

block_cipher = None


a = Analysis(['../src/main.py'],
             pathex=['../src'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

a.datas += Tree('../src', prefix='.', excludes=['main.py'])

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

options = [('u', None, 'OPTION'), ('v', None, 'OPTION'), ('w', None, 'OPTION')]

exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='main',
          debug=True,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True,
          icon='../src/gooey-images/program_icon.ico')

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='main')

info_plist = {'addition_prop': 'additional_value'}

app = BUNDLE(coll,
             name='main.app',
             icon='../src/gooey-images/program_icon.ico',
             bundle_identifier=None,
             info_plist=info_plist
            )
