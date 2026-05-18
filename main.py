#!/usr/bin/env python3
"""WatchNext — watch history scraper and recommendation engine."""

import os
import runpy
import subprocess
import sys
from pathlib import Path

# ── Project root ───────────────────────────────────────────────────────────────
# In a PyInstaller bundle, write output to a user-writable location, not the temp dir.
# Must be set before any local module imports (they read this env var at import time).
_BUNDLE = hasattr(sys, '_MEIPASS')
if _BUNDLE:
    if sys.platform == 'darwin':
        # Binary installed to /usr/local/bin (not writable) — use ~/WatchNext instead
        _data_dir = Path.home() / 'WatchNext'
    else:
        # Windows: binary lives in a user-chosen folder; write data alongside it
        _data_dir = Path(sys.executable).parent
    os.environ.setdefault('WATCHNEXT_HOME', str(_data_dir))
    _data_dir.mkdir(parents=True, exist_ok=True)
    # Pin browser cache to a fixed location so install and runtime agree on where they are
    os.environ.setdefault('PLAYWRIGHT_BROWSERS_PATH', str(_data_dir / '.browsers'))

_BASE = Path(sys._MEIPASS) if _BUNDLE else Path(__file__).parent


# ── Playwright browser check ───────────────────────────────────────────────────
def _ensure_browsers():
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch(headless=True)
            b.close()
    except Exception:
        print("First-time setup: downloading Chromium browser (~150 MB).")
        print("This takes a minute or two and only happens once.\n")
        try:
            from playwright._impl._driver import compute_driver_executable
            node, cli = compute_driver_executable()
            result = subprocess.run([str(node), str(cli), "install", "chromium"])
            if result.returncode != 0:
                print("\nERROR: Could not install Chromium. Check your internet connection.")
                sys.exit(1)
            print("\nChromium installed.\n")
        except Exception as e:
            print(f"\nERROR installing Chromium: {e}")
            sys.exit(1)


# ── Script runner ──────────────────────────────────────────────────────────────
def _run(rel_path: str, argv: list = ()):
    if argv:
        sys.argv = ['WatchNext', *argv]
    runpy.run_path(str(_BASE / rel_path), run_name='__main__')


# ── Phase 1 ───────────────────────────────────────────────────────────────────
def _prime(extra):
    _ensure_browsers()
    _run('scraper.py', extra)

def _netflix(extra):
    _ensure_browsers()
    _run('scrapers/netflix_scraper.py', extra)

def _process():
    _run('cleaners/amazon_cleaner.py')
    print()
    _run('cleaners/netflix_cleaner.py')
    print()
    _run('consolidate.py')

def _all(extra):
    print("=" * 60)
    print("  Step 1 of 3 — Amazon Prime Video")
    print("=" * 60)
    _prime(extra)
    print()
    print("=" * 60)
    print("  Step 2 of 3 — Netflix")
    print("=" * 60)
    _netflix(extra)
    print()
    print("=" * 60)
    print("  Step 3 of 3 — Clean and Consolidate")
    print("=" * 60)
    _process()


# ── Phase 2 ───────────────────────────────────────────────────────────────────
def _match(extra):
    _run('recommender/01_match_tmdb.py', extra)

def _recommend(extra):
    _run('recommender/02_fetch_recs.py', extra)
    print()
    _run('recommender/03_aggregate.py')

def _rate():
    _run('recommender/04_rate_and_refine.py')


# ── Phase 3 ───────────────────────────────────────────────────────────────────
def _scrape_platform(extra):
    _ensure_browsers()
    _run('recommender/05_scrape_platform_recs.py', extra)

def _platform_recs(extra):
    _scrape_platform(extra)
    print()
    _run('recommender/06_aggregate_platform.py')


# ── CLI routing (mirrors run.bat) ─────────────────────────────────────────────
def _cli(cmd: str, extra: list):
    dispatch = {
        'prime':           lambda: _prime(extra),
        'netflix':         lambda: _netflix(extra),
        'process':         _process,
        'match':           lambda: _match(extra),
        'recommend':       lambda: _recommend(extra),
        'rate':            _rate,
        'scrapeplatform':  lambda: _scrape_platform(extra),
        'platformrecs':    lambda: _platform_recs(extra),
    }
    fn = dispatch.get(cmd.lower())
    if fn is None:
        print(f"Unknown command: {cmd}")
        print("Run with no arguments to see the interactive menu.")
        sys.exit(1)
    fn()


# ── Interactive menu ──────────────────────────────────────────────────────────
_MENU = """\
============================================================
  WatchNext
============================================================

  Phase 1 — Watch History
    0  Run all three steps  (recommended)
    1  Scrape Amazon Prime Video
    2  Scrape Netflix
    3  Clean and consolidate

  Phase 2 — TMDB Recommendations  (requires TMDB API key)
    4  Match titles to TMDB
    5  Fetch and rank recommendations
    6  Re-rank using your ratings

  Phase 3 — Platform Recommendations
    7  Scrape platform carousels
    8  Aggregate platform recommendations

  q  Quit

Choice: """

def _menu():
    choice = input(_MENU).strip().lower()
    print()
    actions = {
        '0': lambda: _all([]),
        '1': lambda: _prime([]),
        '2': lambda: _netflix([]),
        '3': _process,
        '4': lambda: _match([]),
        '5': lambda: _recommend([]),
        '6': _rate,
        '7': lambda: _scrape_platform([]),
        '8': lambda: _platform_recs([]),
        'q': lambda: sys.exit(0),
    }
    fn = actions.get(choice)
    if fn is None:
        print(f"Unrecognised choice: {choice!r}")
        sys.exit(1)
    fn()


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == '__main__':
    args = sys.argv[1:]
    if args:
        _cli(args[0], args[1:])
    else:
        _menu()

    if _BUNDLE:
        print()
        input("Press Enter to close...")
