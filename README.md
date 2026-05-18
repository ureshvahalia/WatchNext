# WatchNext

Scrapes your Amazon Prime Video and Netflix watch histories, then generates ranked recommendations using TMDB and the platforms' own "More Like This" carousels.

## Download

Go to the [Releases](../../releases) page and download the file for your platform:

- **Windows** — `WatchNext.exe`
- **macOS** — `WatchNext-mac.pkg` (double-click to install, then run `watchnext` from Terminal)

> **Intel Mac users:** the macOS build is Apple Silicon native. macOS will automatically use Rosetta 2 to run it on Intel — no extra steps needed.

## First run

On the very first launch, WatchNext downloads the Chromium browser it needs (~150 MB). This is a one-time step and takes a minute or two depending on your connection.

## Usage

Double-click `WatchNext.exe` (Windows) or run `watchnext` in Terminal (macOS). An interactive menu appears:

```
  a  Run everything end-to-end  (all phases, no prompts)

  Phase 1 — Watch History
    0  Run all three steps
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

**Most users should press `a`** to run the entire pipeline end-to-end. WatchNext opens a browser window for each platform — log in normally and the scraper takes it from there. Your login session is saved so subsequent runs skip the login step.

On the very first launch, WatchNext downloads the Chromium browser it needs (~150 MB). This is a one-time step and takes a minute or two depending on your connection.

**Windows:** output files are written to an `output/` folder next to the `.exe`.

**macOS:** output files are written to `~/WatchNext/output/` (a `WatchNext` folder in your home directory).

## Ratings

After running Phase 1, open `watch_history.xlsx` and fill in the `Your Rating` column (1–5 stars) for titles you want to influence recommendations. Re-running Phase 2 step 6 ("Re-rank using your ratings") applies those weights.

## Command-line usage

```
watchnext full         # Run everything end-to-end
watchnext prime        # Amazon scraper only
watchnext netflix      # Netflix scraper only
watchnext process      # Clean and consolidate only
watchnext match        # Phase 2: match to TMDB
watchnext recommend    # Phase 2: fetch recommendations
watchnext rate         # Phase 2: re-rank with ratings
watchnext platformrecs # Phase 3: platform recommendations
```

Pass `--fresh` to any scraper step to clear the saved session and re-login. Pass `--debug` to save screenshots for troubleshooting.
