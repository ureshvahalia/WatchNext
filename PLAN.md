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
| Runner | `run.bat Prime / Netflix / Process / (none)` | Done |
| Setup | `setup.bat` | Done |

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

**Base score (`03_aggregate.py`):**
```
score[candidate] = count of watched titles that recommended it
sort key: (-score, -vote_average)   ← vote_average is tiebreaker only
```

**Weighted score (`04_rate_and_refine.py`, requires user ratings):**
```
weighted_score[candidate] = Σ (rating_of_source / 3.0)
                              for each source that recommended the candidate
weight mapping: 5★ → 1.67x  |  3★ → 1.0x (neutral)  |  1★ → 0.33x
unrated titles default to weight 1.0 (as if rated 3★)
sort key: (-weighted_score, -vote_average)
```

### Phase 3 (Platform-based)

```
score[candidate] = count of watched titles whose platform page surfaced it
sort key: -score  (no TMDB rating fallback — metadata not always available)
```

---

## Upcoming — Scoring Improvements

Brainstormed on 2026-05-15. Implementing ideas 4 and 6 next.

### Idea 4 — TMDB rating as multiplicative factor (not just tiebreaker)

**Problem with current design:** TMDB `vote_average` only breaks ties. A title with count=5 and rating=9.5 ranks identically to one with count=5 and rating=4.0 — the tiebreaker only matters when counts collide.

**Proposed change:** Blend the TMDB rating multiplicatively into the base score so a consistently high-rated candidate naturally ranks above a mediocre one at the same frequency count.

**Affects:** `03_aggregate.py` (base score), `04_rate_and_refine.py` (weighted score), and `06_aggregate_platform.py` where TMDB metadata is available.

**Formula decision needed before implementing:**
- Option A: `score = count * log(1 + vote_average)`  — compresses large rating differences; titles with no rating (0.0) get score=0 regardless of count (probably too harsh)
- Option B: `score = count * (0.5 + vote_average / 20.0)` — linear blend; a 10-rated title is 1.5× a 0-rated title; a 7-rated title is ~0.85× baseline; 0-rated titles still score (at 0.5× count)
- Option C: `score = count * max(0.5, vote_average / 10.0)` — floor at 0.5× so unrated/low-rated titles are penalized but not zeroed
- Option D: Keep `vote_average` as tiebreaker but only sort by it within a ±1 count band (e.g., treat counts 5 and 6 as "close enough" and let rating decide)

### Idea 6 — Source-weighted voting (rating-weighted score)

**Current design** already implements basic version in `04_rate_and_refine.py` (`weight = rating / 3.0`, unrated = 1.0).

**Refinement to consider:** The normalization base (÷3.0) makes 3★ neutral and allows >1.0 weights. An alternative is `rating / 5.0` (max=1.0, no amplification above baseline) which is more conservative.

**Decision:** Keep `÷ 3.0` (current) or change to `÷ 5.0`?
- `÷ 3.0`: 5★ sources amplify above the unrated baseline — highly-liked titles punch above their weight. More aggressive personalization.
- `÷ 5.0`: All ratings suppress or equal the unrated baseline — a 5★ source contributes the same as an unrated one, lower ratings discount. Safer but less differentiating.

**Also consider:** Should ratings apply to the platform recs pipeline (`06_aggregate_platform.py`) as a new step, or stay TMDB-only?

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
      │  Frequency score; tiebreak by TMDB rating
      │  → output/recommendations.xlsx  (Phase 2a, base score)
      │
      ▼  (optional, requires Your Rating column filled in)
[Step 4] recommender/04_rate_and_refine.py
      │  Weighted score = Σ (your_rating / 3.0) per source
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
├── scraper.py                          # Amazon Prime scraper
├── consolidate.py                      # Merge + output watch_history.xlsx
├── run.bat                             # Entry point for all steps
├── setup.bat                           # One-time venv + pip install
├── requirements.txt
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
│   ├── .env                            # TMDB_API_KEY (gitignored)
│   └── cache/
│       ├── tmdb_matches.json
│       ├── recs_raw.json
│       └── platform_recs.json
└── output/
    ├── watch_history.xlsx              # Phase 1 output; add Your Rating column here
    ├── recommendations.xlsx            # Phase 2 output
    └── platform_recommendations.xlsx  # Phase 3 output
```

---

*Last updated: 2026-05-15*
