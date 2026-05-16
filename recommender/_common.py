"""Shared utilities for the recommendation pipeline."""
import json
import os
import sys
from pathlib import Path

import openpyxl
import requests
from thefuzz import fuzz

# ── Paths ──────────────────────────────────────────────────────────────────────
ROOT           = Path(__file__).parent.parent
WATCH_FILE     = ROOT / "output" / "watch_history.xlsx"
MATCHES_FILE   = ROOT / "recommender" / "cache" / "tmdb_matches.json"
RECS_FILE      = ROOT / "recommender" / "cache" / "recs_raw.json"
OUT_FILE       = ROOT / "output" / "recommendations.xlsx"

# ── Config ─────────────────────────────────────────────────────────────────────
FILTER_THRESHOLD = 90    # min fuzzy score to flag a candidate as already watched
TMDB_BASE_URL    = "https://www.themoviedb.org"

# ── TMDB genre ID → name (movie + TV combined) ────────────────────────────────
GENRE_MAP = {
    28: "Action",        12: "Adventure",     16: "Animation",   35: "Comedy",
    80: "Crime",         99: "Documentary",   18: "Drama",    10751: "Family",
    14: "Fantasy",       36: "History",       27: "Horror",   10402: "Music",
  9648: "Mystery",    10749: "Romance",      878: "Sci-Fi",  10770: "TV Movie",
    53: "Thriller",   10752: "War",           37: "Western",
 10759: "Action & Adventure", 10762: "Kids",   10763: "News",  10764: "Reality",
 10765: "Sci-Fi & Fantasy",  10766: "Soap",   10767: "Talk", 10768: "War & Politics",
}


def get_api_key() -> str:
    """Return TMDB API key from TMDB_API_KEY env var or recommender/.env file."""
    key = os.environ.get("TMDB_API_KEY", "").strip()
    if key:
        return key
    env_file = Path(__file__).parent / ".env"
    if env_file.exists():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("TMDB_API_KEY="):
                return line.split("=", 1)[1].strip()
    return ""


def require_api_key() -> str:
    """Return the API key or print setup instructions and exit."""
    key = get_api_key()
    if not key:
        print("ERROR: TMDB API key not found.")
        print()
        print("To get a free key:")
        print("  1. Create an account at  https://www.themoviedb.org/signup")
        print("  2. Go to                 https://www.themoviedb.org/settings/api")
        print("  3. Click 'Create' and select Developer / personal use")
        print("  4. Copy either value — both are accepted:")
        print("       API Key (v3 auth)         short alphanumeric string")
        print("       Read Access Token (v4)    long string starting with eyJ...")
        print()
        print("Then store your key in ONE of these ways:")
        print()
        print("  Option A — environment variable (current session only):")
        print("    set TMDB_API_KEY=your_key_here")
        print()
        print("  Option B — file (persists across sessions):")
        print(r"    Create the file  recommender\.env  containing:")
        print("    TMDB_API_KEY=your_key_here")
        sys.exit(1)
    return key


def tmdb_get(url: str, api_key: str, params: dict | None = None) -> requests.Response:
    """GET a TMDB endpoint, handling both v3 API key and v4 Read Access Token.

    TMDB issues two credential types:
      - API Key (v3): short alphanumeric, passed as ?api_key=...
      - Read Access Token (v4): JWT starting with eyJ, sent as Bearer header.
    Both authenticate v3 endpoints; the token just travels differently.
    """
    params = params or {}
    if api_key.startswith("eyJ"):
        r = requests.get(url, params=params,
                         headers={"Authorization": f"Bearer {api_key}"},
                         timeout=10)
    else:
        r = requests.get(url, params={"api_key": api_key, **params}, timeout=10)
    r.raise_for_status()
    return r


def load_watch_history() -> tuple[dict, list]:
    """Return (col_index_map, rows).

    col_index_map: {header_name: 0-based index into the values tuple}
    rows: list of raw value tuples (one per data row)
    """
    wb = openpyxl.load_workbook(WATCH_FILE)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col = {h: i for i, h in enumerate(headers)}
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    return col, rows


def watched_title_set() -> set:
    """Return set of lowercased watched title names."""
    col, rows = load_watch_history()
    name_idx = col.get("Name", 1)
    return {str(row[name_idx]).lower().strip() for row in rows if row[name_idx]}


def is_already_watched(title: str, watched_set: set) -> bool:
    t = title.lower().strip()
    if t in watched_set:
        return True
    return any(fuzz.token_sort_ratio(t, wt) >= FILTER_THRESHOLD for wt in watched_set)


def load_matches() -> dict:
    if not MATCHES_FILE.exists():
        return {}
    return json.loads(MATCHES_FILE.read_text(encoding="utf-8"))


def load_recs() -> dict:
    if not RECS_FILE.exists():
        return {}
    return json.loads(RECS_FILE.read_text(encoding="utf-8"))


def save_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
