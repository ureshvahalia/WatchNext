#!/usr/bin/env python3
"""
Amazon Watch History Cleaner

Reads the latest amazon_watch_history_*.csv from output/ and applies three steps:
  1. Deduplicate by title (keep first occurrence)
  2. Add Type column: "Series" if content_type is "TV Show" (scraped from "Episodes Watched"
     button presence), else "Movie"
  3. Add Name column: title up to " - Season" or " Season" for Series, else title

Writes output/amazon_clean.csv
"""

import csv
import re
import os
import sys
from pathlib import Path

PROJECT_DIR = Path(os.environ['WATCHNEXT_HOME']) if 'WATCHNEXT_HOME' in os.environ else Path(__file__).parent.parent
OUTPUT_DIR  = PROJECT_DIR / "output"


def find_latest_raw() -> Path | None:
    candidates = sorted(
        OUTPUT_DIR.glob("amazon_watch_history_*.csv"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def get_name(title: str, type_: str) -> str:
    if type_ == "Movie":
        return title
    # Strip the " - Season N" or " Season N" suffix to get the bare show name.
    m = re.search(r"\s*-?\s*Season\s+\d+\s*$", title, re.IGNORECASE)
    if m:
        return title[: m.start()].strip().rstrip(",.")
    return title


def clean(input_path: Path, output_path: Path) -> int:
    with open(input_path, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    # Step 1: deduplicate by title, keeping first occurrence (preserves order)
    seen_titles: set[str] = set()
    deduped = []
    for row in rows:
        title = (row.get("title") or "").strip()
        if title and title not in seen_titles:
            seen_titles.add(title)
            deduped.append(row)

    # Steps 2 & 3: classify and extract name; dedup again by Name so multiple
    # seasons of the same show (e.g. "Suits Season 1" / "Suits, Season 9")
    # collapse to a single entry (keep the first one encountered).
    seen_names: set[str] = set()
    result = []
    for row in deduped:
        title = (row.get("title") or "").strip()
        content_type = (row.get("content_type") or "").strip()
        if content_type == "TV Show":
            type_ = "Series"
        elif re.search(r"\s*-?\s*Season\s+\d+\s*$", title, re.IGNORECASE):
            type_ = "Series"  # fallback when button wasn't detected
        else:
            type_ = "Movie"
        name  = get_name(title, type_)
        if name.lower() in seen_names:
            continue
        seen_names.add(name.lower())
        result.append({
            "Type":         type_,
            "Name":         name,
            "title":        title,
            "date_watched": (row.get("date_watched") or "").strip(),
            "url":          (row.get("url") or "").strip(),
        })

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["Type", "Name", "title", "date_watched", "url"])
        writer.writeheader()
        writer.writerows(result)

    return len(result)


def main():
    raw = find_latest_raw()
    if not raw:
        print(
            "ERROR: No Amazon raw CSV found in output/.\n"
            "       Run the Amazon scraper first (run.bat Prime)."
        )
        sys.exit(1)

    out = OUTPUT_DIR / "amazon_clean.csv"
    print(f"Amazon cleaner: reading {raw.name}")
    count = clean(raw, out)
    print(f"  -> {count} unique titles -> {out.name}")


if __name__ == "__main__":
    main()
