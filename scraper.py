#!/usr/bin/env python3
"""
Amazon Prime Video Watch History Scraper

Opens a real browser window, waits for you to log in once, then scrapes
your entire watch history and saves it to a CSV file.

Usage:
    python scraper.py              # normal run
    python scraper.py --debug      # dump page HTML for troubleshooting
    python scraper.py --fresh      # clear saved login and re-login
"""

import argparse
import asyncio
import csv
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

# ── paths ──────────────────────────────────────────────────────────────────────
PROJECT_DIR = Path(os.environ['WATCHNEXT_HOME']) if 'WATCHNEXT_HOME' in os.environ else Path(__file__).parent
OUTPUT_DIR = PROJECT_DIR / "output"
SESSION_DIR = PROJECT_DIR / ".session"   # stores login cookies

# ── Amazon URLs ────────────────────────────────────────────────────────────────
LOGIN_URL = "https://www.amazon.com/gp/sign-in.html"
WATCH_HISTORY_URLS = [
    "https://www.primevideo.com/settings/watch-history",
    "https://www.amazon.com/gp/yourstore/iyr/",
]

# JavaScript injected into the page to find watch history items.
# Works regardless of Amazon's obfuscated/hashed CSS class names by relying
# on URL patterns, aria-label, img[alt], and DOM structure.
_JS_EXTRACT = """
() => {
    const DATE_RE = /(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\\s+\\d{1,2},?\\s+\\d{4}|\\d{1,2}[\\/-]\\d{1,2}[\\/-]\\d{2,4}|(?:Today|Yesterday|\\d+\\s+days?\\s+ago)/i;

    const seen  = new Set();
    const items = [];

    // Collect every <a> that links to a content detail page
    const anchors = document.querySelectorAll(
        'a[href*="/detail/"], a[href*="/dp/"], a[href*="/gp/video/detail"]'
    );

    for (const a of anchors) {
        const href = a.href || '';
        if (!href) continue;

        // Deduplicate on the path segment that identifies the title
        const key = href.replace(/[?#].*$/, '').replace(/\\/ref=.*$/, '');
        if (seen.has(key)) continue;
        seen.add(key);

        // ── title: prefer aria-label > img[alt] > link text ──────────────────
        let title = (a.getAttribute('aria-label') || '').trim();
        if (!title) {
            const img = a.querySelector('img');
            title = img ? (img.getAttribute('alt') || '').trim() : '';
        }
        if (!title) title = a.textContent.trim();
        if (!title || title.length < 2) continue;

        // ── card container: walk up until we find a node with ≥2 children ────
        let card = a.parentElement;
        for (let i = 0; i < 8 && card && card.tagName !== 'BODY'; i++) {
            if (card.children.length >= 2 || card.tagName === 'LI') break;
            card = card.parentElement;
        }

        // ── date: Amazon groups items under a date section heading that is a
        // preceding sibling of the item container, not inside the card itself.
        // Walk up the DOM and at each level check sibling elements for a date.
        let date = '';
        let el = a.parentElement;
        for (let i = 0; i < 12 && el && el.tagName !== 'BODY'; i++) {
            const parent = el.parentElement;
            if (parent) {
                for (const sib of parent.children) {
                    if (sib === el || sib.contains(a)) continue;
                    const m = sib.textContent.trim().match(DATE_RE);
                    if (m) { date = m[0]; break; }
                }
            }
            if (date) break;
            el = el.parentElement;
        }

        // ── "Episodes Watched" button presence → series, absence → movie ──────
        const cardEl = card || a;
        const hasEpisodesBtn = Array.from(
            cardEl.querySelectorAll('button, [role="button"], a')
        ).some(el =>
            /episodes?\\s+watched/i.test(
                (el.textContent || el.getAttribute('aria-label') || '').trim()
            )
        );
        const contentType = hasEpisodesBtn ? 'TV Show' : '';

        items.push({
            title:        title.substring(0, 300),
            content_type: contentType,
            date_watched: date,
            url:          href,
        });
    }

    return items;
}
"""

# Same as _JS_EXTRACT but walks much further up the DOM and also checks
# preceding sibling section headers (Amazon sometimes groups items under
# a date heading that is a sibling of the list, not inside each card).
_JS_EXTRACT_DIAG = """
() => {
    const DATE_RE = /(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\\s+\\d{1,2},?\\s+\\d{4}|\\d{1,2}[\\/-]\\d{1,2}[\\/-]\\d{2,4}|(?:Today|Yesterday|\\d+\\s+days?\\s+ago|\\d+\\s+(?:week|month|year)s?\\s+ago)/i;

    const seen  = new Set();
    const diag  = [];
    const anchors = document.querySelectorAll(
        'a[href*="/detail/"], a[href*="/dp/"], a[href*="/gp/video/detail"]'
    );

    for (const a of anchors) {
        const href = a.href || '';
        if (!href) continue;
        const key = href.replace(/[?#].*$/, '').replace(/\\/ref=.*$/, '');
        if (seen.has(key)) continue;
        seen.add(key);

        let title = (a.getAttribute('aria-label') || a.querySelector('img')?.getAttribute('alt') || a.textContent || '').trim();
        if (!title || title.length < 2) continue;

        // Walk up 12 levels (much deeper than normal extract)
        let card = a.parentElement;
        const levels = [];
        for (let i = 0; i < 12 && card && card.tagName !== 'BODY'; i++) {
            levels.push(card.textContent.trim().substring(0, 200));
            card = card.parentElement;
        }

        // Check preceding siblings at each ancestor level for date headers
        let siblingText = '';
        let el = a.parentElement;
        for (let i = 0; i < 12 && el && el.tagName !== 'BODY'; i++) {
            const parent = el.parentElement;
            if (parent) {
                for (const sib of parent.children) {
                    if (sib === el || sib.contains(a)) continue;
                    const t = sib.textContent.trim();
                    if (DATE_RE.test(t)) { siblingText = t.substring(0, 200); break; }
                }
            }
            if (siblingText) break;
            el = el?.parentElement;
        }

        diag.push({ title: title.substring(0, 100), levels, siblingText });
        if (diag.length >= 5) break;   // only need a few samples
    }
    return diag;
}
"""


# ── helpers ────────────────────────────────────────────────────────────────────

async def is_logged_in(page) -> bool:
    try:
        await page.wait_for_selector(
            "#nav-link-accountList, [data-nav-role='signin'], .nav-line-1",
            timeout=5_000,
        )
        content = (await page.text_content("#nav-link-accountList") or "")
        return "Hello" in content or "Account" in content
    except PlaywrightTimeoutError:
        return False


async def wait_for_login(page):
    print("\n  A browser window has opened. Please log in to your Amazon account.")
    print("  The scraper will automatically continue once you are signed in.\n")
    try:
        await page.wait_for_function(
            """() => {
                const el = document.querySelector('#nav-link-accountList');
                return el && (el.innerText.includes('Hello') || el.innerText.includes('Account'));
            }""",
            timeout=300_000,
        )
    except PlaywrightTimeoutError:
        print("  Login timed out after 5 minutes. Exiting.")
        sys.exit(1)
    print("  Login detected! Continuing...\n")


# ── extraction ─────────────────────────────────────────────────────────────────

async def js_extract(page) -> list[dict]:
    """Run the JS extractor and return deduplicated items."""
    try:
        return await page.evaluate(_JS_EXTRACT)
    except Exception:
        return []


async def js_diagnose(page):
    """Print card-text samples to diagnose missing date extraction."""
    try:
        samples = await page.evaluate(_JS_EXTRACT_DIAG)
    except Exception as e:
        print(f"  Diagnose error: {e}")
        return
    print("\n── Card text diagnostic (first 5 items) ──────────────────")
    for i, s in enumerate(samples, 1):
        print(f"\n  [{i}] {s['title']}")
        print(f"      Sibling date text: {s['siblingText']!r}")
        for j, lvl in enumerate(s['levels']):
            snippet = lvl.replace('\n', ' ')[:120]
            print(f"      Level {j+1:>2}: {snippet!r}")
    print("──────────────────────────────────────────────────────────\n")


# ── scrolling + progressive collection ────────────────────────────────────────

async def scroll_and_collect(page, debug: bool = False) -> list[dict]:
    """
    Scroll to the bottom while harvesting items on every step.
    Progressive collection handles virtual-DOM trimming: items removed from
    the DOM while scrolling are captured before they disappear.
    """
    collected: dict[str, dict] = {}   # keyed by URL to deduplicate

    def merge(batch: list[dict]):
        for item in batch:
            collected.setdefault(item["url"] or item["title"], item)

    prev_height = -1
    stall_count = 0
    step = 0

    while stall_count < 4:
        step += 1

        # Harvest whatever is in the DOM right now
        merge(await js_extract(page))
        print(f"  Step {step:>3} — {len(collected)} unique titles so far", end="\r")

        # Click any load-more / see-more button
        for btn_text in ["Load more", "See more", "Show more", "Load More", "See More"]:
            btn = page.locator(f"button:has-text('{btn_text}'), a:has-text('{btn_text}')")
            if await btn.count() > 0:
                try:
                    await btn.first.click()
                    await page.wait_for_timeout(2_000)
                    if debug:
                        print(f"\n  Clicked '{btn_text}' button")
                except Exception:
                    pass

        try:
            await page.evaluate("window.scrollTo(0, document.documentElement.scrollHeight)")
            await page.wait_for_timeout(1_500)
            height = await page.evaluate("document.documentElement.scrollHeight")
        except Exception:
            break   # page navigated mid-scroll

        if height == prev_height:
            stall_count += 1
        else:
            stall_count = 0
            prev_height = height

    # Final harvest after scroll settles
    merge(await js_extract(page))
    print(f"\n  Scroll complete — {len(collected)} unique titles found.")
    return list(collected.values())


# ── CSV export ─────────────────────────────────────────────────────────────────

def save_csv(items: list[dict], path: Path):
    OUTPUT_DIR.mkdir(exist_ok=True)
    fieldnames = ["title", "content_type", "date_watched", "url"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(items)
    print(f"\n  Saved {len(items)} items → {path}")


# ── debug helpers ──────────────────────────────────────────────────────────────

async def dump_debug(page, label: str):
    debug_dir = OUTPUT_DIR / "debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%H%M%S")
    html_path = debug_dir / f"{label}_{ts}.html"
    html_path.write_text(await page.content(), encoding="utf-8")
    screenshot_path = debug_dir / f"{label}_{ts}.png"
    await page.screenshot(path=str(screenshot_path), full_page=True)
    print(f"  Debug dump: {html_path.name} + {screenshot_path.name}")


# ── main ───────────────────────────────────────────────────────────────────────

async def run(debug: bool, fresh: bool, diagnose: bool = False):
    if fresh and SESSION_DIR.exists():
        shutil.rmtree(SESSION_DIR)
        print("  Cleared saved session. You will need to log in again.\n")

    SESSION_DIR.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_csv = OUTPUT_DIR / f"amazon_watch_history_{timestamp}.csv"

    print("Amazon Prime Video Watch History Scraper")
    print("=" * 45)

    async with async_playwright() as pw:
        # persistent context keeps you logged in between runs
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
        print("\n[1/3] Checking login status...")
        await page.goto("https://www.amazon.com", wait_until="domcontentloaded", timeout=30_000)

        if not await is_logged_in(page):
            await page.goto(LOGIN_URL, wait_until="domcontentloaded")
            await wait_for_login(page)
        else:
            print("  Already logged in (session restored).\n")

        # ── step 2: navigate to watch history ─────────────────────────────────
        print("[2/3] Opening watch history page...")
        for url in WATCH_HISTORY_URLS:
            try:
                # "commit" fires as soon as the server responds (before JS runs),
                # avoiding ERR_ABORTED from redirects or SPA navigation aborts.
                await page.goto(url, wait_until="commit", timeout=30_000)
                await page.wait_for_load_state("domcontentloaded", timeout=15_000)
                await page.wait_for_timeout(2_000)
                print(f"  Landed on: {page.url}")
                break
            except Exception as e:
                print(f"  Could not load {url}: {type(e).__name__} — trying next...")

        # primevideo.com uses a separate SSO session — detect the sign-in redirect
        # and wait for the user to complete it before continuing.
        if "signin" in page.url or "ap/signin" in page.url:
            print(
                "\n  Prime Video requires you to log in separately.\n"
                "  Please complete the sign-in in the browser window.\n"
                "  The scraper will continue automatically once you're done..."
            )
            try:
                await page.wait_for_url(
                    lambda u: "signin" not in u and "ap/signin" not in u,
                    timeout=300_000,
                )
                await page.wait_for_load_state("domcontentloaded", timeout=15_000)
                await page.wait_for_timeout(2_000)
                print(f"  Signed in. Now on: {page.url}\n")
            except PlaywrightTimeoutError:
                print("  Login timed out. Exiting.")
                await context.close()
                return

        # If we still aren't on the watch history page, ask the user to navigate there
        if "watch-history" not in page.url and "yourstore" not in page.url:
            print(
                f"\n  Current page: {page.url}\n"
                "  Could not reach the watch history page automatically.\n"
                "  Please navigate to it manually in the browser, then press Enter..."
            )
            input("  > Press Enter when ready: ")

        if debug:
            await dump_debug(page, "before_scroll")

        if diagnose:
            await js_diagnose(page)
            await context.close()
            return

        # ── step 3: scroll and scrape ──────────────────────────────────────────
        print("\n[3/3] Scrolling and scraping...")
        items = await scroll_and_collect(page, debug=debug)

        if debug:
            await dump_debug(page, "after_scroll")

        if not items:
            print(
                "\n  No watch history items could be extracted.\n"
                "  This usually means Amazon changed their page layout.\n"
                "  Run with --debug to capture the page HTML for analysis."
            )
            await dump_debug(page, "no_items_found")
        else:
            save_csv(items, output_csv)
            print(f"  Done. {len(items)} titles scraped.")

        await context.close()


def main():
    parser = argparse.ArgumentParser(description="Scrape Amazon Prime Video watch history to CSV")
    parser.add_argument("--debug",    action="store_true", help="Save page HTML + screenshot for troubleshooting")
    parser.add_argument("--fresh",    action="store_true", help="Clear saved login session and re-login")
    parser.add_argument("--diagnose", action="store_true", help="Print card-text samples to debug date extraction, then exit")
    args = parser.parse_args()

    asyncio.run(run(debug=args.debug, fresh=args.fresh, diagnose=args.diagnose))


if __name__ == "__main__":
    main()
