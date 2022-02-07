# -*- mode: python ; coding: utf-8 -*-

from PyInstaller import compat

from PyInstaller.building.api import EXE, PYZ
from PyInstaller.building.build_main import Analysis
from PyInstaller.building.osx import BUNDLE

block_cipher = None

a = Analysis(['../src/main.py'],
             pathex=['../src'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None,
             )

a.datas += Tree('../src', prefix='.')

pyz = PYZ(a.pure)

options = [('u', None, 'OPTION'), ('v', None, 'OPTION'), ('w', None, 'OPTION')]


exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          options,
          name='Trendy',
          debug=False,
          strip=None,
          upx=True,
          console=False,
          icon='../src/gooey-images/program_icon.ico')

info_plist = {'addition_prop': 'additional_value'}
app = BUNDLE(exe,
             name='Trendy.app',
             bundle_identifier=None,
             info_plist=info_plist
            )