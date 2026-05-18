# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for WatchNext
#
# Build commands:
#   Windows:  pyinstaller watchnext.spec
#   macOS:    pyinstaller watchnext.spec
#
# Output:
#   Windows:  dist/WatchNext.exe
#   macOS:    dist/WatchNext

from pathlib import Path
import playwright as _pw
_pw_driver = str(Path(_pw.__file__).parent / 'driver')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Bundle all .py source files as data so runpy.run_path can execute them.
        # These land in sys._MEIPASS at runtime (the temp extraction directory).
        ('scraper.py',    '.'),
        ('consolidate.py', '.'),
        ('cleaners',      'cleaners'),
        ('scrapers',      'scrapers'),
        ('recommender',   'recommender'),
        # Playwright driver (Node.js binary + CLI) needed to install browsers at runtime.
        (_pw_driver,      'playwright/driver'),
    ],
    hiddenimports=[
        'playwright',
        'playwright.sync_api',
        'playwright.async_api',
        'playwright._impl._driver',
        'playwright._impl._browser_type',
        'openpyxl',
        'openpyxl.styles',
        'thefuzz',
        'thefuzz.fuzz',
        'thefuzz.process',
        'requests',
        'requests.adapters',
        'urllib3',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='WatchNext',
    debug=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    argv_emulation=False,
)
