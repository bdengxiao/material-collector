# -*- mode: python ; coding: utf-8 -*-
a = Analysis(['src/app.py'], pathex=['src'], binaries=[], datas=[], hiddenimports=[], hookspath=[], hooksconfig={}, runtime_hooks=[], excludes=[], noarchive=False, optimize=0)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='MaterialCollector', debug=False, bootloader_ignore_signals=False, strip=False, upx=False, console=False, disable_windowed_traceback=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name='MaterialCollector')
