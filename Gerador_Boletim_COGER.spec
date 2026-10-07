# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
from pathlib import Path
project = Path(SPECPATH)
ctk_datas, ctk_binaries, ctk_hidden = collect_all("customtkinter")
tkcal_datas, tkcal_binaries, tkcal_hidden = collect_all("tkcalendar")
babel_datas, babel_binaries, babel_hidden = collect_all("babel")
datas = [(str(project / "assets"), "assets"), (str(project / "vendor" / "tesseract"), "tesseract")] + ctk_datas
hiddenimports = ["win32com.client", "pythoncom", "pywintypes"] + ctk_hidden + tkcal_hidden + babel_hidden
a = Analysis(["main.py"], pathex=[str(project)], binaries=ctk_binaries, datas=datas, hiddenimports=hiddenimports, hookspath=[], hooksconfig={}, runtime_hooks=[], excludes=[], noarchive=False)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="Gerador_Boletim_COGER", debug=False, bootloader_ignore_signals=False, strip=False, upx=True, console=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=True, upx_exclude=[], name="Gerador_Boletim_COGER")
