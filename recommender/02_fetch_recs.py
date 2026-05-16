"""
Step 2 -- Fetch TMDB recommendations and similar titles for each matched entry.

For each matched title, calls two TMDB endpoints:
  /{type}/{id}/recommendations  -- editorially curated "if you liked X"
  /{type}/{id}/similar          -- genre/tag-based similarity

Results include both movies and TV series regardless of the source type
(cross-type recommendations are intentionally included).

REC_PAGES controls pages fetched per endpoint (20 results/page):
  --pages 1  (default) -- up to 20 results per endpoint  = up to 40 per title
  --pages 2             -- up to 40 results per endpoint  = up to 80 per title

Caches results in recommender/cache/recs_raw.json; re-running skips already-
cached titles unless --reset is passed.

Usage:
  python recommender/02_fetch_recs.py
  python recommender/02_fetch_recs.py --pages 2
  python recommender/02_fetch_recs.py --reset
"""
import sys
import time
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).parent))
from _common import RECS_FILE, require_api_key, load_matches, load_recs, save_json, tmdb_get

TMDB_BASE = "https://api.themoviedb.org/3"


# ── arg parsing ────────────────────────────────────────────────────────────────

def parse_args() -> tuple[int, bool]:
    pages = 1
    reset = False
    args  = sys.argv[1:]
    for i, a in enumerate(args):
        if a == "--pages" and i + 1 < len(args):
            try:
                pages = max(1, int(args[i + 1]))
            except ValueError:
                pass
        if a == "--reset":
            reset = True
    return pages, reset


# ── TMDB helpers ───────────────────────────────────────────────────────────────

def fetch_page(tmdb_id: int, media_type: str, endpoint: str, page: int, api_key: str) -> list:
    mtype = "movie" if media_type == "Movie" else "tv"
    r = tmdb_get(
        f"{TMDB_BASE}/{mtype}/{tmdb_id}/{endpoint}",
        api_key,
        params={"page": page},
    )
    return r.json().get("results", [])


def normalize_rec(r: dict) -> dict:
    is_movie = "title" in r
    return {
        "tmdb_id":      r["id"],
        "title":        r.get("title") or r.get("name", ""),
        "type":         "Movie" if is_movie else "Series",
        "year":         (r.get("release_date") or r.get("first_air_date") or "")[:4],
        "vote_average": r.get("vote_average", 0),
        "genre_ids":    r.get("genre_ids", []),
    }


def fetch_all_recs(tmdb_id: int, media_type: str, rec_pages: int, api_key: str) -> list:
    results = []
    for endpoint in ("recommendations", "similar"):
        for page in range(1, rec_pages + 1):
            try:
                batch = fetch_page(tmdb_id, media_type, endpoint, page, api_key)
                results.extend(batch)
                time.sleep(0.25)
                if len(batch) < 20:   # reached the last page early
                    break
            except requests.HTTPError as e:
                if e.response is not None and e.response.status_code == 404:
                    break  # no recommendations available for this title
                print(f"    WARN: {endpoint} p{page}: {e}")
            except Exception as e:
                print(f"    WARN: {endpoint} p{page}: {e}")
    return results


# ── main ───────────────────────────────────────────────────────────────────────

def main():
    rec_pages, reset = parse_args()
    api_key  = require_api_key()
    matches  = load_matches()
    recs     = {} if reset else load_recs()

    matched_titles = [(name, m) for name, m in matches.items() if m.get("matched")]
    total          = len(matched_titles)

    print(f"Fetching recommendations for {total} matched titles "
          f"({rec_pages} page(s) per endpoint, up to {rec_pages * 20} results each)...")

    fetched = 0
    for i, (name, m) in enumerate(matched_titles, 1):
        source_id = str(m["tmdb_id"])
        if source_id in recs and not reset:
            continue

        print(f"[{i}/{total}] {name} (TMDB {source_id})...", end=" ", flush=True)
        raw  = fetch_all_recs(m["tmdb_id"], m["media_type"], rec_pages, api_key)
        recs[source_id] = [normalize_rec(r) for r in raw]
        save_json(RECS_FILE, recs)
        fetched += 1
        print(f"{len(recs[source_id])} recs")

    print()
    print(f"Done.  Fetched: {fetched}  |  Already cached: {total - fetched}")


if __name__ == "__main__":
    main()
