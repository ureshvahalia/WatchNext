# Stream History Recommender — Project Plan

## Goal

A self-contained, packaged tool that anyone can run to:
1. Extract their watch history from Amazon Prime Video and Netflix
2. Clean and consolidate it into a single structured list
3. Generate ranked recommendations based on what similar viewers watched, refined by personal ratings

---

## Status

### Phase 1 — Data Extraction & Cleaning  ✅ COMPLETE

| Module | File | Status |
|---|---|---|
| Amazon scraper | `scraper.py` | Done — Playwright, persistent login, JS DOM extraction |
| Netflix scraper | `scrapers/netflix_scraper.py` | Done — Playwright, profile selection, `/viewingactivity` page, captures Netflix title URLs |
| Amazon cleaner | `cleaners/amazon_cleaner.py` | Done — deduplicates by title then by extracted name, classifies Movie/Series, strips season suffix |
| Netflix cleaner | `cleaners/netflix_cleaner.py` | Done — deduplicates by Netflix URL then by name, classifies by colon-count + 3 name-extraction rules |
| Consolidator | `consolidate.py` | Done — fuzzy-merges both lists (token_sort_ratio ≥ 90, type must match), outputs `watch_history.xlsx` with clickable URLs |
| Runner (dev) | `run.bat Prime / Netflix / Process / (none)` | Done |
| Runner (packaged) | `main.py` — interactive menu + `full` end-to-end mode | Done |
| Setup (dev) | `setup.bat` | Done |
| Packaging | PyInstaller + GitHub Actions → `WatchNext.exe` / `WatchNext-mac.pkg` | Done |

**Output:** `output/watch_history.xlsx`
Columns: `Type | Name | source | url (hyperlink) | last_watched`
~250–1100 rows depending on how much of each platform has been scraped.

---

### Phase 2 — TMDB Recommendation Engine  ✅ COMPLETE

| Step | File | Status |
|---|---|---|
| Step 1 — TMDB Matching | `recommender/01_match_tmdb.py` | Done |
| Step 2 — Fetch Recs | `recommender/02_fetch_recs.py` | Done |
| Step 3 — Aggregate & Rank | `recommender/03_aggregate.py` | Done |
| Step 4 — Rate & Refine | `recommender/04_rate_and_refine.py` | Done |

**Output:** `output/recommendations.xlsx`
Sheets: `Recommendations | Audit | Unmatched`

**Open decisions from original plan — all resolved:**

| # | Question | Decision |
|---|---|---|
| 1 | Rating UX | Option B — `Your Rating` column (1–5) in `watch_history.xlsx`; user fills in Excel |
| 2 | Unmatched titles | Reported in a separate `Unmatched` sheet in `recommendations.xlsx` |
| 3 | Cross-type recommendations | Included — a movie can recommend a TV series and vice versa |
| 4 | Recommendation cap per source | 20 per endpoint (1 page default); `--pages N` flag for more |

---

### Phase 3 — Platform Recommendation Scraping  ✅ COMPLETE

*(Not in original plan — added after Phase 2 to cross-reference TMDB recs with what the platforms themselves surface.)*

| Step | File | Status |
|---|---|---|
| Step 5 — Scrape platform carousels | `recommender/05_scrape_platform_recs.py` | Done — Playwright, scrapes Netflix & Prime rec carousels per watched title |
| Step 6 — Aggregate platform recs | `recommender/06_aggregate_platform.py` | Done |

**Output:** `output/platform_recommendations.xlsx`
Sheets: `Recommendations | Audit | No_Recs`

`run.bat` commands: `ScrapePlatform` (step 5 only) | `PlatformRecs` (steps 5+6)

---

## Current Scoring Design

### Phase 2 (TMDB-based)

**Base score (`03_aggregate.py`, columns: Freq + Score):**
```
freq[candidate]  = count of watched titles that recommended it
score[candidate] = freq * log(1 + tmdb_vote_average)

Titles with no TMDB rating (0 or None) use 5.0 as a neutral midpoint
so they are not zeroed out — absent data ≠ bad quality.
```

**Weighted score (`04_rate_and_refine.py`, columns: Freq + Wtd Score + Score):**
```
wtd_score[candidate] = Σ (your_rating_of_source / 3.0)
                         for each source that recommended the candidate
weight mapping: 5★ → 1.67x  |  3★ → 1.0x (neutral)  |  1★ → 0.33x
unrated titles default to 3★ (weight 1.0)

score[candidate] = wtd_score * log(1 + tmdb_vote_average)
```

Both pipelines sort by `score` descending. `Freq`/`Wtd Score` are shown for transparency.

**To rerun with new scoring (no need to re-fetch from TMDB):**
```
run.bat Recommend   ← reruns Step 3; Step 2 skips instantly (all cached)
run.bat Rate        ← reruns Step 4 (only if Your Rating column is filled)
```

### Phase 3 (Platform-based)

```
score[candidate] = count of watched titles whose platform page surfaced it
sort key: -score  (TMDB metadata not always available for platform titles)
```

---

## Completed Scoring Improvements

### ✅ Idea 4 — TMDB rating as multiplicative factor (2026-05-18)

**Decision:** Option A — `score = freq * log(1 + vote_average)`.
Titles with 0/None rating use 5.0 neutral fallback instead of being zeroed.
Implemented in `_common.py` (`tmdb_score_factor()`), `03_aggregate.py`, `04_rate_and_refine.py`.

### ✅ Idea 6 — Source-weighted voting (2026-05-18)

**Decision:** `weight = rating / 3.0`, unrated = 3★ (weight 1.0), scale 1–5.
3★ is the neutral baseline; higher ratings amplify, lower ones discount.
Combined with Idea 4: final `score = wtd_score * log(1 + tmdb_vote_average)`.
Implemented in `04_rate_and_refine.py`.

---

## Deferred Ideas

### Idea 8 — Collaborative filtering hybrid (deferred — needs new data source)

True collaborative filtering requires a user-item rating matrix (many users × many titles). The TMDB API does not expose individual user ratings in bulk.

**Practical path if revisited:** Download the [MovieLens](https://grouplens.org/datasets/movielens/) dataset (free, ~33M ratings, includes TMDB IDs). Map watched TMDB movie IDs into MovieLens space, find taste-similar users, pull their highly-rated unwatched titles.

**Limitations:** MovieLens is movies-only — TV series are not covered. Only worth implementing if a significant portion of the watch history is movies and the user has rated enough titles to form a meaningful taste vector.

**Prerequisites before implementing:**
1. User has rated their watched titles (in `watch_history.xlsx`)
2. Decision on whether movie-only coverage is acceptable

---

## Data Flow (Current — All Phases)

```
watch_history.xlsx
      │
      ▼
[Step 1] recommender/01_match_tmdb.py
      │  Search each title on TMDB; cache → recommender/cache/tmdb_matches.json
      │
      ▼
[Step 2] recommender/02_fetch_recs.py
      │  /recommendations + /similar per matched title
      │  Cache → recommender/cache/recs_raw.json
      │
      ▼
[Step 3] recommender/03_aggregate.py
      │  score = freq * log(1 + tmdb_vote_average)
      │  → output/recommendations.xlsx  (Phase 2a, base score)
      │
      ▼  (optional, requires Your Rating column filled in)
[Step 4] recommender/04_rate_and_refine.py
      │  score = Σ(rating/3.0) * log(1 + tmdb_vote_average)
      │  → output/recommendations.xlsx  (Phase 2b, overwrites Recommendations sheet)
      │
      │
      └─── (independent pipeline) ──────────────────────────────────────────────
[Step 5] recommender/05_scrape_platform_recs.py
      │  Playwright — scrapes Netflix & Prime rec carousels per watched title
      │  Cache → recommender/cache/platform_recs.json
      │
      ▼
[Step 6] recommender/06_aggregate_platform.py
         Frequency score from platform carousel data
         → output/platform_recommendations.xlsx
```

---

## File Layout

```
WatchNext/
├── main.py                             # Packaged entry point — interactive menu + CLI
├── watchnext.spec                      # PyInstaller build spec
├── entitlements.plist                  # macOS hardened runtime entitlements
├── scraper.py                          # Amazon Prime scraper
├── consolidate.py                      # Merge + output watch_history.xlsx
├── run.bat                             # Dev entry point for all steps
├── setup.bat                           # Dev one-time venv + pip install
├── requirements.txt
├── .github/
│   └── workflows/
│       └── build.yml                   # CI: builds + signs + notarizes + publishes releases
├── scrapers/
│   └── netflix_scraper.py
├── cleaners/
│   ├── amazon_cleaner.py
│   └── netflix_cleaner.py
├── recommender/
│   ├── _common.py                      # Shared utilities + paths
│   ├── 01_match_tmdb.py
│   ├── 02_fetch_recs.py
│   ├── 03_aggregate.py
│   ├── 04_rate_and_refine.py
│   ├── 05_scrape_platform_recs.py
│   ├── 06_aggregate_platform.py
│   ├── .env                            # TMDB_API_KEY (gitignored; baked into binary at build time)
│   └── cache/
│       ├── tmdb_matches.json
│       ├── recs_raw.json
│       └── platform_recs.json
└── output/
    ├── watch_history.xlsx              # Phase 1 output; add Your Rating column here
    ├── recommendations.xlsx            # Phase 2 output
    └── platform_recommendations.xlsx  # Phase 3 output
```

### Packaged app data locations

| Platform | Data directory |
|---|---|
| Windows | Folder containing `WatchNext.exe` |
| macOS | `~/WatchNext/` |

---

*Last updated: 2026-05-18*
