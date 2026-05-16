#!/usr/bin/env python3
"""
Netflix Watch History Scraper

Opens a real browser window, waits for you to log in once, then scrapes
your entire viewing history (clicking through all "Load more" pages) and
saves it to a CSV file including per-title Netflix URLs.

Usage:
    python scrapers/netflix_scraper.py           # normal run
    python scrapers/netflix_scraper.py --debug   # dump page HTML for troubleshooting
    python scrapers/netflix_scraper.py --fresh   # clear saved login and re-login
"""

import argparse
import asyncio
import csv
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

# ── paths ──────────────────────────────────────────────────────────────────────
PROJECT_DIR = Path(__file__).parent.parent
OUTPUT_DIR  = PROJECT_DIR / "output"
SESSION_DIR = PROJECT_DIR / ".session" / "netflix"

# ── URLs ───────────────────────────────────────────────────────────────────────
LOGIN_URL    = "https://www.netflix.com/login"
ACTIVITY_URL = "https://www.netflix.com/viewingactivity"

# JavaScript that extracts title links and dates from the viewing activity page.
# Relies on href patterns and DOM proximity rather than obfuscated CSS classes.
_JS_EXTRACT = """
() => {
    // (?!\d) prevents matching a partial number: "26" in "26365 Days" fails
    // because "26" is followed by "3", but "26" at end-of-field passes.
    const DATE_RE = /(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\\s+\\d{1,2},?\\s+\\d{4}|\\d{1,2}\\/\\d{1,2}\\/\\d{2,4}(?!\\d)/i;

    const seen  = new Set();
    const items = [];

    const anchors = document.querySelectorAll('a[href*="/title/"], a[href*="/watch/"]');

    for (const a of anchors) {
        const href = a.getAttribute('href') || '';
        if (!href || href === '#') continue;

        const url = (href.startsWith('http') ? href : 'https://www.netflix.com' + href)
                    .replace(/[?#].*$/, '');

        if (seen.has(url)) continue;
        seen.add(url);

        const title = (a.textContent || a.getAttribute('aria-label') || '').trim();
        if (!title || title.length < 2) continue;

        let date = '';
        let el   = a.parentElement;
        for (let depth = 0; depth < 8 && el && el.tagName !== 'BODY'; depth++) {
            const parent = el.parentElement;
            if (!parent) break;
            for (const sib of Array.from(parent.children)) {
                if (sib.contains(a)) continue;
                const text = sib.textContent.trim();
                const m    = text.match(DATE_RE);
                if (m) { date = m[0]; break; }
            }
            if (date) break;
            el = parent;
        }

        items.push({ url, title, date });
    }

    return items;
}
"""

# Diagnostic: for the first 3 anchors, show text at each DOM level and in
# direct siblings of the anchor itself, so we can pinpoint where the date lives.
_JS_DIAG = """
() => {
    // (?!\d) prevents matching a partial number: "26" in "26365 Days" fails
    // because "26" is followed by "3", but "26" at end-of-field passes.
    const DATE_RE = /(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\\s+\\d{1,2},?\\s+\\d{4}|\\d{1,2}\\/\\d{1,2}\\/\\d{2,4}(?!\\d)/i;
    const seen = new Set();
    const out  = [];

    for (const a of document.querySelectorAll('a[href*="/title/"], a[href*="/watch/"]')) {
        const href = a.getAttribute('href') || '';
        if (!href || href === '#') continue;
        const url = (href.startsWith('http') ? href : 'https://www.netflix.com' + href).replace(/[?#].*$/, '');
        if (seen.has(url)) continue;
        seen.add(url);

        const title = (a.textContent || a.getAttribute('aria-label') || '').trim();
        if (!title || title.length < 2) continue;

        // Direct siblings of the anchor
        const aSiblings = [];
        for (const sib of (a.parentElement?.children || [])) {
            if (sib === a) continue;
            aSiblings.push(sib.textContent.trim().substring(0, 80));
        }

        // Ancestor levels with sibling text and match info
        const levels = [];
        let el = a.parentElement;
        for (let depth = 0; depth < 10 && el && el.tagName !== 'BODY'; depth++) {
            const parent = el.parentElement;
            if (!parent) break;
            const sibTexts = [];
            for (const sib of parent.children) {
                if (sib.contains(a)) continue;
                const t = sib.textContent.trim().substring(0, 100);
                const m = t.match(DATE_RE);
                sibTexts.push({ text: t, match: m ? m[0] : null });
            }
            levels.push(sibTexts);
            el = parent;
        }

        out.push({ title: title.substring(0, 60), aSiblings, levels });
        if (out.length >= 3) break;
    }
    return out;
}
"""

# JavaScript that finds and clicks the "Load more" pagination button.
# Restricted to button and role=button elements with an EXACT text match to
# avoid accidentally clicking "see more" / "show more" links on title cards.
# Returns true if the button was found and clicked, false otherwise.
_JS_CLICK_LOAD_MORE = """
() => {
    const targets = new Set(['load more', 'show more']);
    for (const el of document.querySelectorAll('button, [role="button"]')) {
        const text = (el.textContent || el.getAttribute('aria-label') || '').trim().toLowerCase();
        if (targets.has(text)) {
            el.scrollIntoView({ block: 'center' });
            el.click();
            return true;
        }
    }
    return false;
}
"""


# ── helpers ────────────────────────────────────────────────────────────────────

async def js_diagnose(page):
    """Print per-item DOM structure to locate where dates live."""
    try:
        samples = await page.evaluate(_JS_DIAG)
    except Exception as e:
        print(f"  Diagnose error: {e}")
        return
    print("\n── Netflix date diagnostic (first 3 items) ───────────────")
    for i, s in enumerate(samples, 1):
        print(f"\n  [{i}] {s['title']}")
        print(f"      Direct <a> siblings: {s['aSiblings']}")
        for j, sibs in enumerate(s['levels']):
            for sib in sibs:
                flag = f"  ← DATE MATCH: {sib['match']!r}" if sib['match'] else ""
                print(f"      Level {j+1:>2} sib: {sib['text']!r}{flag}")
    print("──────────────────────────────────────────────────────────\n")


async def is_logged_in(page) -> bool:
    return "login" not in page.url.lower() and "netflix.com" in page.url.lower()


async def wait_for_login(page):
    print("\n  A browser window has opened. Please log in to your Netflix account.")
    print("  The scraper will continue automatically once you are signed in.\n")
    try:
        await page.wait_for_url(
            lambda u: "netflix.com" in u and "login" not in u,
            timeout=300_000,
        )
    except PlaywrightTimeoutError:
        print("  Login timed out after 5 minutes. Exiting.")
        sys.exit(1)
    await page.wait_for_load_state("domcontentloaded", timeout=15_000)
    print("  Login detected! Continuing...\n")


# ── extraction ─────────────────────────────────────────────────────────────────

def _normalize_date(date: str) -> str:
    """Expand 2-digit year to 4-digit: '5/13/26' → '5/13/2026'."""
    import re
    return re.sub(r'^(\d{1,2}/\d{1,2}/)(\d{2})$', lambda m: m.group(1) + '20' + m.group(2), date)


async def js_extract(page) -> list[dict]:
    try:
        items = await page.evaluate(_JS_EXTRACT)
        for item in items:
            if item.get('date'):
                item['date'] = _normalize_date(item['date'])
        return items
    except Exception:
        return []


# ── load all items (with incremental CSV flush) ────────────────────────────────

async def load_all_and_collect(page, output_csv: Path) -> int:
    """
    Click "Load more" until exhausted, writing new items to disk after each
    batch so the output file can be monitored while the scraper runs.

    Returns the total number of unique titles collected.
    """
    # Counts <a> tags for /title/ and /watch/ currently in the DOM.
    # Used to detect when Netflix has injected the next batch of items.
    _DOM_COUNT = "() => document.querySelectorAll('a[href*=\"/title/\"], a[href*=\"/watch/\"]').length"

    collected:    dict[str, dict] = {}   # all seen items keyed by url-or-title
    written_keys: set[str]        = set()  # keys already flushed to disk

    OUTPUT_DIR.mkdir(exist_ok=True)
    csv_file = open(output_csv, "w", newline="", encoding="utf-8")
    writer   = csv.DictWriter(csv_file, fieldnames=["title", "date", "url"])
    writer.writeheader()
    csv_file.flush()

    run_start = time.monotonic()

    def merge(batch: list[dict]) -> int:
        """Add new items to collected; return count of newly added."""
        added = 0
        for item in batch:
            key = item["url"] or item["title"]
            if key not in collected:
                collected[key] = item
                added += 1
        return added

    def flush_new():
        """Write any collected items not yet on disk; flush the file."""
        new_rows = []
        for key, item in collected.items():
            if key not in written_keys:
                new_rows.append(item)
                written_keys.add(key)
        if new_rows:
            writer.writerows(new_rows)
            csv_file.flush()
        return len(new_rows)

    step        = 0
    stall_count = 0

    while stall_count < 3:
        step += 1
        step_start = time.monotonic()

        added = merge(await js_extract(page))
        flushed = flush_new()

        elapsed   = time.monotonic() - run_start
        tag_parts = []

        dom_links_before = await page.evaluate(_DOM_COUNT)
        clicked = await page.evaluate(_JS_CLICK_LOAD_MORE)

        if clicked:
            tag_parts.append("btn")
            new_links = False
            for _ in range(25):
                await page.wait_for_timeout(200)
                if await page.evaluate(_DOM_COUNT) > dom_links_before:
                    stall_count = 0
                    new_links = True
                    break
            if not new_links:
                stall_count += 1
                tag_parts.append(f"no new links — stall {stall_count}/3")
        else:
            tag_parts.append("scroll")
            await page.evaluate("window.scrollTo(0, document.documentElement.scrollHeight)")
            await page.wait_for_timeout(1_500)
            added += merge(await js_extract(page))
            flush_new()

            if added == 0:
                stall_count += 1
                tag_parts.append(f"stall {stall_count}/3")
            else:
                stall_count = 0

        step_s = time.monotonic() - step_start
        print(
            f"  Step {step:>4} | {len(collected):>4} titles (+{flushed:<3})"
            f" | {', '.join(tag_parts)}"
            f" | step {step_s:.1f}s  elapsed {elapsed/60:.1f}min"
        )

    # Final harvest after the last batch settles.
    merge(await js_extract(page))
    flushed = flush_new()
    csv_file.close()

    total_min = (time.monotonic() - run_start) / 60
    print(f"\n  Collection complete — {len(collected)} unique titles"
          f" ({flushed} in final flush) in {total_min:.1f} min")
    print(f"  Written to: {output_csv}")
    return len(collected)


# ── debug helpers ──────────────────────────────────────────────────────────────

async def dump_debug(page, label: str):
    debug_dir = OUTPUT_DIR / "debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%H%M%S")
    html_path = debug_dir / f"netflix_{label}_{ts}.html"
    html_path.write_text(await page.content(), encoding="utf-8")
    shot_path = debug_dir / f"netflix_{label}_{ts}.png"
    await page.screenshot(path=str(shot_path), full_page=True)
    print(f"  Debug dump: {html_path.name} + {shot_path.name}")


# ── main ───────────────────────────────────────────────────────────────────────

async def run(debug: bool, fresh: bool, diagnose: bool = False):
    if fresh and SESSION_DIR.exists():
        shutil.rmtree(SESSION_DIR)
        print("  Cleared saved Netflix session. You will need to log in again.\n")

    SESSION_DIR.mkdir(parents=True, exist_ok=True)

    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_csv = OUTPUT_DIR / f"netflix_raw_{timestamp}.csv"

    print("Netflix Watch History Scraper")
    print("=" * 35)

    async with async_playwright() as pw:
        context = await pw.chromium.launch_persistent_context(
            user_data_dir=str(SESSION_DIR),
            headless=False,
            viewport={"width": 1280, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            args=["--disable-blink-features=AutomationControlled"],
        )

        page = await context.new_page()

        # ── step 1: ensure logged in ───────────────────────────────────────────
        print("\n[1/4] Checking login status...")
        await page.goto("https://www.netflix.com", wait_until="domcontentloaded", timeout=30_000)
        await page.wait_for_timeout(2_000)

        if not await is_logged_in(page):
            await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=30_000)
            await wait_for_login(page)
        else:
            print("  Already logged in (session restored).\n")

        # ── step 2: profile selection ──────────────────────────────────────────
        print("[2/4] Profile selection...")
        await page.goto("https://www.netflix.com/profiles", wait_until="domcontentloaded", timeout=30_000)
        await page.wait_for_timeout(2_000)

        if "login" in page.url:
            print("\n  Session expired. Please log in again.")
            await wait_for_login(page)
            await page.goto("https://www.netflix.com/profiles", wait_until="domcontentloaded", timeout=30_000)
            await page.wait_for_timeout(2_000)

        print("\n  Please click your profile in the browser window.")
        print("  The scraper will continue automatically once you've chosen.\n")
        try:
            await page.wait_for_url(
                lambda u: "netflix.com/profiles" not in u,
                timeout=300_000,
            )
        except PlaywrightTimeoutError:
            print("  Timed out waiting for profile selection. Exiting.")
            await context.close()
            return

        # ── step 3: navigate to viewing activity ───────────────────────────────
        print("[3/4] Opening viewing activity page...")
        await page.goto(ACTIVITY_URL, wait_until="commit", timeout=30_000)
        await page.wait_for_load_state("domcontentloaded", timeout=15_000)
        await page.wait_for_timeout(2_000)
        print(f"  Landed on: {page.url}\n")

        if debug:
            await dump_debug(page, "before_load")

        if diagnose:
            await js_diagnose(page)
            await context.close()
            return

        # ── step 4: click through all pages and scrape ────────────────────────
        print("[4/4] Scrolling and scraping...")
        print(f"  Output file: {output_csv}\n")
        total = await load_all_and_collect(page, output_csv)

        if debug:
            await dump_debug(page, "after_load")

        if total == 0:
            print(
                "\n  No viewing history items could be extracted.\n"
                "  This usually means Netflix changed their page layout.\n"
                "  Run with --debug to capture the page HTML for analysis."
            )
            await dump_debug(page, "no_items_found")

        await context.close()


def main():
    parser = argparse.ArgumentParser(description="Scrape Netflix viewing history to CSV")
    parser.add_argument("--debug",    action="store_true", help="Save page HTML + screenshot for troubleshooting")
    parser.add_argument("--fresh",    action="store_true", help="Clear saved login session and re-login")
    parser.add_argument("--diagnose", action="store_true", help="Print DOM structure for first 3 items to debug date extraction, then exit")
    args = parser.parse_args()

    asyncio.run(run(debug=args.debug, fresh=args.fresh, diagnose=args.diagnose))


if __name__ == "__main__":
    main()
