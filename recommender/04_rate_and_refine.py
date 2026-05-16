"""
Step 4 -- Re-weight recommendations by user ratings.

Reads the 'Your Rating' column (values 1-5) from watch_history.xlsx.
Applies weight = rating / 3.0 to each source title's contributions:
  5 stars -> 1.67x  (strong signal: find me more like this)
  3 stars -> 1.00x  (neutral, same as unrated)
  1 star  -> 0.33x  (weak signal: soft-suppress similar titles)
Unrated titles default to weight 1.0.

Rewrites the Recommendations sheet in output/recommendations.xlsx,
adding a 'Weighted Score' column. The Unmatched sheet is preserved.

Usage:
  python recommender/04_rate_and_refine.py
"""
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

sys.path.insert(0, str(Path(__file__).parent))
from _common import (
    WATCH_FILE, OUT_FILE, GENRE_MAP, TMDB_BASE_URL,
    load_matches, load_recs, watched_title_set, is_already_watched,
)


def load_ratings() -> dict[str, float]:
    wb  = openpyxl.load_workbook(WATCH_FILE)
    ws  = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col = {h: i for i, h in enumerate(headers)}
    if "Your Rating" not in col:
        print("No 'Your Rating' column found in watch_history.xlsx.")
        print("Run  run.bat Match  first to add the column, then enter ratings in Excel.")
        return {}
    name_idx   = col["Name"]
    rating_idx = col["Your Rating"]
    ratings: dict[str, float] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        name   = row[name_idx]
        rating = row[rating_idx]
        if name and rating is not None:
            try:
                val = float(rating)
                if 1.0 <= val <= 5.0:
                    ratings[str(name)] = val
            except (ValueError, TypeError):
                pass
    return ratings


def main():
    matches = load_matches()
    recs    = load_recs()
    ratings = load_ratings()
    watched = watched_title_set()

    if not ratings:
        print("No valid ratings found (1-5 in the 'Your Rating' column). Nothing to refine.")
        sys.exit(0)

    rated_count = sum(1 for n in matches if n in ratings and matches[n].get("matched"))
    print(f"Applying {rated_count} ratings to re-rank recommendations...")

    watched_ids = {str(m["tmdb_id"]) for m in matches.values() if m.get("matched")}
    weighted: dict[str, dict] = defaultdict(lambda: {"weighted_score": 0.0, "count": 0, "data": None})

    for source_name, m in matches.items():
        if not m.get("matched"):
            continue
        source_id = str(m["tmdb_id"])
        rating    = ratings.get(source_name, 3.0)   # unrated = neutral weight
        weight    = rating / 3.0
        seen: set[str] = set()
        for rec in recs.get(source_id, []):
            tid = str(rec["tmdb_id"])
            if tid in seen:
                continue
            seen.add(tid)
            if tid in watched_ids:
                continue
            if is_already_watched(rec["title"], watched):
                continue
            weighted[tid]["weighted_score"] += weight
            weighted[tid]["count"]          += 1
            if weighted[tid]["data"] is None:
                weighted[tid]["data"] = rec

    ranked = sorted(
        [(tid, v) for tid, v in weighted.items() if v["data"]],
        key=lambda x: (-(x[1]["weighted_score"]), -(x[1]["data"].get("vote_average") or 0)),
    )

    # Preserve the Unmatched sheet from the existing file if it exists.
    wb = openpyxl.load_workbook(OUT_FILE) if OUT_FILE.exists() else openpyxl.Workbook()
    if "Recommendations" in wb.sheetnames:
        del wb["Recommendations"]
    ws = wb.create_sheet("Recommendations", 0)

    headers = ["Rank", "Title", "Type", "Score", "Weighted Score", "TMDB Rating", "Genres", "Year", "TMDB Link"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for rank, (tid, v) in enumerate(ranked, 1):
        d      = v["data"]
        genres = ", ".join(GENRE_MAP.get(gid, str(gid)) for gid in d.get("genre_ids", []))
        mtype  = d.get("type", "")
        seg    = "movie" if mtype == "Movie" else "tv"
        link   = f"{TMDB_BASE_URL}/{seg}/{tid}"
        ws.append([
            rank, d["title"], mtype, v["count"],
            round(v["weighted_score"], 2),
            round(d.get("vote_average") or 0, 1),
            genres, d.get("year", ""), link,
        ])

    try:
        wb.save(OUT_FILE)
    except PermissionError:
        print(f"ERROR: Cannot write {OUT_FILE.name} -- close it in Excel and retry.")
        sys.exit(1)

    print(f"Re-ranked {len(ranked)} recommendations saved to {OUT_FILE}")


if __name__ == "__main__":
    main()
