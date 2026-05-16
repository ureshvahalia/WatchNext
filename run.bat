@echo off
setlocal

:: ── activate venv if present ──────────────────────────────────────────────────
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

:: ── route to the right step ───────────────────────────────────────────────────
set CMD=%~1

if "%CMD%"==""           goto :all
if /i "%CMD%"=="Prime"     goto :prime
if /i "%CMD%"=="Netflix"   goto :netflix
if /i "%CMD%"=="Process"   goto :process
if /i "%CMD%"=="Match"          goto :match
if /i "%CMD%"=="Recommend"      goto :recommend
if /i "%CMD%"=="Rate"           goto :rate
if /i "%CMD%"=="ScrapePlatform" goto :scrapeplatform
if /i "%CMD%"=="PlatformRecs"   goto :platformrecs

echo Unknown command: %CMD%
echo.
echo Usage:  run.bat [Prime ^| Netflix ^| Process ^| Match ^| Recommend ^| Rate]
echo.
echo   Phase 1 -- Data extraction:
echo   Prime      - Scrape Amazon Prime Video watch history
echo   Netflix    - Scrape Netflix viewing history
echo   Process    - Clean and consolidate into watch_history.xlsx
echo   (none)     - Run all three Phase 1 steps in sequence
echo.
echo   Phase 2 -- TMDB recommendations (requires a TMDB API key):
echo   Match      - Match watch history titles to TMDB  [Step 1]
echo   Recommend  - Fetch and rank recommendations      [Steps 2+3]
echo   Rate       - Re-rank using your ratings          [Step 4]
echo.
echo   Phase 3 -- Platform recommendations (uses saved login sessions):
echo   ScrapePlatform - Scrape Netflix/Prime rec carousels  [Step 5]
echo   PlatformRecs   - Aggregate into platform_recommendations.xlsx [Steps 5+6]
echo.
echo   TMDB API key: set TMDB_API_KEY=your_key_here
echo   Or create the file  recommender\.env  with that line.
echo   Free key at: https://www.themoviedb.org/settings/api
echo.
echo   Extra flags for Phase 1 scrapers:
echo   --debug     Save page HTML + screenshots for troubleshooting
echo   --fresh     Clear saved login session and re-login
echo   --diagnose  Print card-text samples to debug date extraction, then exit
echo.
echo   Extra flags for Phase 2 steps:
echo   --reset     Clear cache and re-fetch everything from TMDB
echo   --pages N   Pages per endpoint for Recommend (default 1 = top 20 each)
echo.
echo   Extra flags for Phase 3 steps:
echo   --reset         Clear platform rec cache and re-scrape
echo   --platform X    Scrape only Netflix or Amazon
echo   --tabs N        Parallel browser tabs (default 3, try 4-5)
echo   --limit N       Only scrape first N titles (for testing)
echo   --debug         Save screenshots to output\debug\platform_recs\
goto :done

:: ── run all three steps in sequence ──────────────────────────────────────────
:all
echo ============================================================
echo  Step 1 of 3 -- Amazon Prime Video
echo ============================================================
call :run_prime %2 %3 %4
echo.
echo ============================================================
echo  Step 2 of 3 -- Netflix
echo ============================================================
call :run_netflix %2 %3 %4
echo.
echo ============================================================
echo  Step 3 of 3 -- Clean and Consolidate
echo ============================================================
call :run_process
goto :done

:prime
call :run_prime %2 %3 %4
goto :done

:netflix
call :run_netflix %2 %3 %4
goto :done

:process
call :run_process
goto :done

:match
call :run_match %2 %3 %4
goto :done

:recommend
call :run_recommend %2 %3 %4
goto :done

:rate
call :run_rate
goto :done

:scrapeplatform
call :run_scrapeplatform %2 %3 %4 %5 %6
goto :done

:platformrecs
call :run_scrapeplatform %2 %3 %4 %5 %6
if errorlevel 1 goto :done
echo.
call :run_aggplatform
goto :done

:: ── subroutines ───────────────────────────────────────────────────────────────
:run_prime
python scraper.py %*
goto :eof

:run_netflix
python scrapers\netflix_scraper.py %*
goto :eof

:run_process
python cleaners\amazon_cleaner.py
if errorlevel 1 echo   WARNING: Amazon cleaner reported issues -- continuing.
echo.
python cleaners\netflix_cleaner.py
if errorlevel 1 echo   WARNING: Netflix cleaner reported issues -- continuing.
echo.
python consolidate.py
goto :eof

:run_match
python recommender\01_match_tmdb.py %*
goto :eof

:run_recommend
python recommender\02_fetch_recs.py %*
if errorlevel 1 goto :eof
echo.
python recommender\03_aggregate.py
goto :eof

:run_rate
python recommender\04_rate_and_refine.py
goto :eof

:run_scrapeplatform
python recommender\05_scrape_platform_recs.py %*
goto :eof

:run_aggplatform
python recommender\06_aggregate_platform.py
goto :eof

:done
echo.
pause
