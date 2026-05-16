#!/usr/bin/env python3
"""
Netflix Watch History Cleaner

Reads the latest netflix_raw_*.csv from output/ and:
  1. Deduplicates by Netflix URL (/title/{id} is show-level, so all episodes of
     the same series collapse to one entry)
  2. Adds a secondary name-based dedup for any remaining duplicates
  3. Classifies Type: "Series" if title has 2+ colons, else "Movie"
  4. Extracts Name via 3 rules: (1) identical first two colon-parts → first part,
     (2) second part contains Season/Episode/Limited Series → first part,
     (3) otherwise → everything up to the second colon

Writes output/netflix_clean.csv
"""

import csv
import re
import sys
from pathlib import Path

PROJECT_DIR = Path(__file__).parent.parent
OUTPUT_DIR  = PROJECT_DIR / "output"


def find_latest_raw() -> Path | None:
    candidates = sorted(
        OUTPUT_DIR.glob("netflix_raw_*.csv"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def classify_netflix(title: str) -> tuple[str, str]:
    """Return (type, name) for a Netflix title using colon-count + 3 name rules."""
    parts = title.split(":")
    if len(parts) < 3:
        return "Movie", title.strip()

    p0 = parts[0].strip()
    p1 = parts[1].strip()

    # Rule 1: first two elements are identical → first element is the series name.
    if p0 == p1:
        return "Series", p0

    # Rule 2: second element contains "Season", "Episode", or "Limited Series".
    if re.search(r"\b(Season|Episode|Limited Series)\b", p1, re.IGNORECASE):
        return "Series", p0

    # Rule 3: name is everything up to the second colon.
    return "Series", (parts[0] + ":" + parts[1]).strip()


def clean(input_path: Path, output_path: Path) -> int:
    with open(input_path, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    # The raw CSV is in newest-first order (matches Netflix page order).
    # Primary dedup: by Netflix URL — every episode of a series shares the same
    #   /title/{id} URL, so this collapses the series to one entry.
    # Secondary dedup: by extracted name — catches entries that lack a URL or
    #   appear under slightly different episode-level /watch/ URLs.
    seen_urls:  set[str] = set()
    seen_names: set[str] = set()
    result = []

    for row in rows:
        url   = (row.get("url")   or "").strip().rstrip("/")
        title = (row.get("title") or "").strip()
        if not title:
            continue

        type_, name = classify_netflix(title)
        name_key = name.lower()

        if url:
            if url in seen_urls:
                continue
            seen_urls.add(url)

        if name_key in seen_names:
            continue
        seen_names.add(name_key)

        result.append({
            "Type":  type_,
            "Name":  name,
            "title": title,
            "date":  (row.get("date") or "").strip(),
            "url":   url,
        })

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["Type", "Name", "title", "date", "url"])
        writer.writeheader()
        writer.writerows(result)

    return len(result)


def main():
    raw = find_latest_raw()
    if not raw:
        print(
            "ERROR: No Netflix raw CSV found in output/.\n"
            "       Run the Netflix scraper first (run.bat Netflix)."
        )
        sys.exit(1)

    out = OUTPUT_DIR / "netflix_clean.csv"
    print(f"Netflix cleaner: reading {raw.name}")
    count = clean(raw, out)
    print(f"  -> {count} unique titles -> {out.name}")


if __name__ == "__main__":
    main()
