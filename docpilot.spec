# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the DocPilot desktop build.

Layout matters: every module derives its asset paths from __file__ walked up
to the project root, and PyInstaller maps frozen modules under sys._MEIPASS
with their package path intact - so shipping 'assets' and 'frontend/dist' at
the bundle root makes every existing path resolve with no code changes.
"""

datas = [
    ('assets', 'assets'),                    # templates, brand images, bundled fonts
    ('frontend/dist', 'frontend/dist'),      # the built UI
]

hiddenimports = [
    'win32com', 'win32com.client', 'pythoncom', 'pywintypes',
    'uvicorn.logging', 'uvicorn.loops', 'uvicorn.loops.auto',
    'uvicorn.protocols', 'uvicorn.protocols.http', 'uvicorn.protocols.http.auto',
    'uvicorn.protocols.websockets', 'uvicorn.protocols.websockets.auto',
    'uvicorn.lifespan', 'uvicorn.lifespan.on',
    'backend.app', 'backend.jobs', 'backend.drive', 'backend.categories',
    'backend.fonts',
    'backend.pipeline.daily_ca', 'backend.pipeline.daily_ca_hindi',
    'backend.pipeline.prelims_ab', 'backend.pipeline.prelims_qb',
    'backend.pipeline.mains_ab', 'backend.pipeline.fitter',
    'backend.pipeline.daily_ca_xylem', 'backend.pipeline.cover_xylem',
    'backend.pipeline.cover', 'backend.pipeline.cover_hindi',
    'backend.pipeline.styler', 'backend.pipeline.template',
    'backend.pipeline.postprocess', 'backend.pipeline.convert',
    'backend.pipeline.docx_tools', 'backend.pipeline.ooxml',
    # Daily Editorial Analysis editions (2026-09-21)
    'backend.pipeline.dea_common', 'backend.pipeline.dea_english',
    'backend.pipeline.dea_hindi', 'backend.pipeline.combo_english',
    'backend.pipeline.combo_hindi', 'backend.pipeline.combo_hindi_layout', 'backend.pipeline.cover_kit',
    'segno', 'zxingcpp',
]

a = Analysis(
    ['desktop_launcher.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy.testing', 'PIL.ImageQt'],   # tkinter IS the UI
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name='DocPilot',
    icon='assets/icon/docpilot.ico',
    version='tools/file_version.txt',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,                # desktop_launcher shows a real status window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe, a.binaries, a.datas,
    strip=False, upx=False, upx_exclude=[],
    name='DocPilot-2.0-beta',
)
