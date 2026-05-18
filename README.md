# WatchNext

Scrapes your Amazon Prime Video and Netflix watch histories, then generates ranked recommendations using TMDB and the platforms' own "More Like This" carousels.

## Download

Go to the [Releases](../../releases) page and download the file for your platform:

- **Windows** — `WatchNext.exe`
- **macOS** — `WatchNext-mac.zip` (unzip, then run `WatchNext` from Terminal)

> **Intel Mac users:** the macOS build is Apple Silicon native. macOS will automatically use Rosetta 2 to run it on Intel — no extra steps needed.

## First run

On the very first launch, WatchNext downloads the Chromium browser it needs (~150 MB). This is a one-time step and takes a minute or two depending on your connection.

## Setup

### TMDB API key (required for Phases 2 and 3)

1. Create a free account at [themoviedb.org](https://www.themoviedb.org/signup)
2. Go to Settings → API and request a free API key
3. Create a file called `.env` inside the `recommender/` folder next to the executable:
   ```
   TMDB_API_KEY=your_key_here
   ```

## Usage

Double-click `WatchNext.exe` (Windows) or run `./WatchNext` in Terminal (macOS). An interactive menu appears:

```
  Phase 1 — Watch History
    0  Run all three steps  (recommended)
    1  Scrape Amazon Prime Video
    2  Scrape Netflix
    3  Clean and consolidate

  Phase 2 — TMDB Recommendations
    4  Match titles to TMDB
    5  Fetch and rank recommendations
    6  Re-rank using your ratings

  Phase 3 — Platform Recommendations
    7  Scrape platform carousels
    8  Aggregate platform recommendations
```

Run Phase 1 first. It opens a browser window for each platform — log in normally and the scraper takes it from there. Your login session is saved so subsequent runs skip the login step.

**Windows:** output files are written to an `output/` folder next to the `.exe`.

**macOS:** output files are written to `~/WatchNext/output/` (a `WatchNext` folder in your home directory).

## Command-line usage

Power users can skip the menu by passing a command directly:

```
WatchNext prime
WatchNext netflix
WatchNext process
WatchNext match
WatchNext recommend
WatchNext rate
WatchNext platformrecs
```

Pass `--fresh` to any scraper step to clear the saved session and re-login. Pass `--debug` to save screenshots for troubleshooting.
